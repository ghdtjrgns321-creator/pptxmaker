"""차트·다이어그램 렌더 모듈 — pptx-visuals 스킬의 단일 출처.

build_pptx.py가 import해서 사용한다. 차트는 네이티브 차트(PPT에서 데이터 수정 가능),
다이어그램은 도형+커넥터 조합(flow/layers)으로 렌더한다.
색·폰트·크기는 인자로 받은 brand dict(brand-kit.yaml)에서만 나온다.
"""

from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


def _rgb(hex6):
    return RGBColor.from_string(hex6)


# --- 아이콘 (Lucide MIT — make_icons.py가 brand 색 PNG로 사전 변환) ---
ICON_DIR = Path(__file__).resolve().parents[2] / "pptx-build" / "assets" / "icons"


def add_icon(slide, name, x, y, size, color="primary"):
    """브랜드 색 아이콘 PNG 삽입. color: primary | accent | white.
    없는 이름이면 명시적 실패(조용한 누락 금지) — 목록은 make_icons.py ICONS."""
    path = ICON_DIR / f"{name}_{color}.png"
    if not path.exists():
        raise ValueError(f"unknown icon: {name} ({path})")
    return slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(size), Inches(size))


CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,  # 항목 간 크기 비교(카테고리명 짧을 때)
    "hbar": XL_CHART_TYPE.BAR_CLUSTERED,  # 순위·비중 비교(카테고리명 길 때 — 레이블 잘림 방지)
    "line": XL_CHART_TYPE.LINE,  # 시계열 추세
    "pie": XL_CHART_TYPE.PIE,  # 구성비(합계 100%·항목 ≤5)
    "doughnut": XL_CHART_TYPE.DOUGHNUT,  # 구성비 + 중앙 여백 활용
    "stacked_bar": XL_CHART_TYPE.COLUMN_STACKED,  # 구성 변화 비교
}


# --- 차트 ---
def add_chart(slide, spec, brand, x, y, w, h):
    """네이티브 차트 삽입 + brand 팔레트 시리즈 색 적용. GraphicFrame 반환."""
    ct = CHART_TYPES.get(spec.get("chart_type", "bar"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    data = CategoryChartData()
    data.categories = spec["categories"]
    for name, values in spec["series"].items():
        data.add_series(name, values)
    gf = slide.shapes.add_chart(ct, Inches(x), Inches(y), Inches(w), Inches(h), data)
    ch = gf.chart
    ch.has_legend = len(spec["series"]) > 1 or ct in (
        XL_CHART_TYPE.PIE,
        XL_CHART_TYPE.DOUGHNUT,
    )
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
    _apply_series_colors(ch, spec.get("chart_type", "bar"), brand)
    _apply_chart_style(ch, spec.get("chart_type", "bar"), brand)
    return gf


def _apply_chart_style(chart, kind, brand):
    """PPT 기본 차트 티 제거 — 자동 제목 삭제, 글자 축소·브랜드 폰트, 데이터 레이블."""
    c, f, s = brand["colors"], brand["fonts"], brand["sizes"]
    chart.has_title = False  # 슬라이드 제목과 중복되는 차트 자동 제목 제거
    chart.font.size = Pt(s.get("caption", 9) + 1)
    chart.font.name = f["body"]
    chart.font.color.rgb = _rgb(c["text"])
    if chart.has_legend:
        chart.legend.font.size = Pt(s.get("caption", 9) + 1)
    for axis_name in ("category_axis", "value_axis"):
        try:
            axis = getattr(chart, axis_name)
            axis.tick_labels.font.size = Pt(s.get("caption", 9) + 1)
            axis.tick_labels.font.name = f["body"]
            axis.format.line.color.rgb = _rgb(c["muted"])
            if axis.has_major_gridlines:
                axis.major_gridlines.format.line.color.rgb = _rgb(c["bg_alt"])
        except ValueError:  # pie 등 축 없는 차트
            continue
    if kind in ("bar", "hbar", "line", "stacked_bar"):
        plot = chart.plots[0]
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(s.get("caption", 9) + 1)
        plot.data_labels.font.name = f["body"]
        plot.data_labels.font.color.rgb = _rgb(c["text"])


def _apply_series_colors(chart, kind, brand):
    """Office 기본 테마색 대신 brand 팔레트(accent→primary→muted 순환)를 입힌다."""
    c = brand["colors"]
    palette = [c["accent"], c["primary"], c["muted"]]
    if kind in ("pie", "doughnut"):
        for i, pt in enumerate(chart.plots[0].series[0].points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = _rgb(palette[i % len(palette)])
        return
    for i, series in enumerate(chart.series):
        color = _rgb(palette[i % len(palette)])
        if kind == "line":
            series.format.line.color.rgb = color
        else:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = color


# --- 다이어그램 (도형+화살표 DSL) ---
def add_diagram(slide, spec, brand, x, y, w, h):
    """diagram 슬라이드 본문 렌더.
    layout: flow(좌→우 화살표) | layers(상→하 적층) | branch(1→N 분기) | timeline(마일스톤)."""
    layout = spec.get("layout", "flow")
    if layout == "layers":
        _layers(slide, spec["nodes"], brand, x, y, w, h)
    elif layout == "flow":
        _flow(slide, spec["nodes"], brand, x, y, w, h)
    elif layout == "branch":
        _branch(slide, spec, brand, x, y, w, h)
    elif layout == "timeline":
        _timeline(slide, spec["nodes"], brand, x, y, w, h)
    elif layout == "cards":
        _cards(slide, spec["nodes"], brand, x, y, w, h)
    elif layout == "from_to":
        _from_to(slide, spec["rows"], brand, x, y, w, h)
    else:
        raise ValueError(f"unknown diagram layout: {layout}")


def _from_to(slide, rows, brand, x, y, w, h):
    """From→To 전환 그리드(BCG p40 실측) — 문제 박스 →쉐브론→ 해결 박스 행 반복.
    rows: [{"from": "...", "to": "...", "label": "행 라벨(선택)"}] 3~5행 권장."""
    c, f, s = brand["colors"], brand["fonts"], brand["sizes"]
    n = len(rows)
    gap = 0.22
    head_h = 0.42  # 열 머리 영역을 내부에서 확보(외부 겹침 방지)
    rh = min(1.15, (h - head_h - gap * (n - 1)) / n)
    has_label = any(r.get("label") for r in rows)
    label_w = 1.9 if has_label else 0.0
    arrow_w = 0.55
    bw = (w - label_w - arrow_w - 0.3) / 2
    # 열 머리(From... / ...To)
    for i, (hx, htxt) in enumerate(
        ((x + label_w, "From — 기존 방식"), (x + label_w + bw + arrow_w + 0.3, "To — 본 시스템"))
    ):
        tb = slide.shapes.add_textbox(Inches(hx), Inches(y), Inches(bw), Inches(0.3))
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = htxt
        run.font.size = Pt(s["body"] - 1)
        run.font.bold = True
        run.font.name = f["head"]
        run.font.color.rgb = _rgb(c["muted"] if i == 0 else c["accent"])
    for i, r in enumerate(rows):
        ry = y + head_h + i * (rh + gap)
        if r.get("label"):
            lb = slide.shapes.add_textbox(Inches(x), Inches(ry), Inches(label_w - 0.15), Inches(rh))
            lb.text_frame.word_wrap = True
            lb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = lb.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = r["label"]
            run.font.size = Pt(s["body"] - 1)
            run.font.bold = True
            run.font.name = f["head"]
            run.font.color.rgb = _rgb(c["primary"])
        _node_box(slide, brand, x + label_w, ry, bw, rh, r["from"], fill=c["bg_alt"])
        chev = slide.shapes.add_shape(
            MSO_SHAPE.CHEVRON,
            Inches(x + label_w + bw + 0.15),
            Inches(ry + rh / 2 - 0.14),
            Inches(arrow_w - 0.15),
            Inches(0.28),
        )
        chev.fill.solid()
        chev.fill.fore_color.rgb = _rgb(c["accent"])
        chev.line.fill.background()
        _node_box(slide, brand, x + label_w + bw + arrow_w + 0.3, ry, bw, rh, r["to"], dark=True)


def _arrowhead(connector):
    """커넥터 끝에 삼각 화살촉(python-pptx 미노출 → XML 직접 부착)."""
    ln = connector.line._get_or_add_ln()
    tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
    ln.append(tail)


def _node_box(slide, brand, x, y, w, h, label, sub=None, fill=None, dark=False):
    """공용 노드 박스. dark=True면 primary 채움 + 백색 텍스트(분기 루트 등 강조용)."""
    c, f, s = brand["colors"], brand["fonts"], brand["sizes"]
    if dark:
        fill = c["primary"]
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    try:
        box.adjustments[0] = 0.08  # 과한 라운딩(AI티) 억제 — 문서형 미세 라운드
    except (IndexError, ValueError):
        pass
    box.fill.solid()
    box.fill.fore_color.rgb = _rgb(fill or c["bg_alt"])
    box.line.color.rgb = _rgb(c["primary"])
    box.line.width = Pt(1.0)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = label
    run.font.size = Pt(s["body"])
    run.font.bold = True
    run.font.name = f["head"]
    run.font.color.rgb = _rgb("FFFFFF" if dark else c["primary"])
    if sub:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run2 = p2.add_run()
        run2.text = sub
        run2.font.size = Pt(s["body"] - 2)
        run2.font.name = f["body"]
        run2.font.color.rgb = _rgb(c["bg_alt"] if dark else c["muted"])
    return box


def _branch(slide, spec, brand, x, y, w, h):
    """1→N 분기 아키텍처 — 좌측 루트(dark) → 엘보 커넥터 → 우측 자식 스택."""
    c = brand["colors"]
    root = spec.get("root", {"label": "분기"})
    nodes = spec["nodes"]
    n = len(nodes)
    rw = min(3.3, w * 0.34)
    rh = 1.5
    ry = y + h / 2 - rh / 2
    _node_box(slide, brand, x, ry, rw, rh, root.get("label", ""), root.get("sub"), dark=True)
    cx = x + rw + 1.2
    cw = w - rw - 1.2
    gap = 0.35
    ch = min(1.5, (h - gap * (n - 1)) / n)
    top = y + (h - (ch * n + gap * (n - 1))) / 2
    for i, nd in enumerate(nodes):
        cy = top + i * (ch + gap)
        _node_box(slide, brand, cx, cy, cw, ch, nd["label"], nd.get("sub"))
        conn = slide.shapes.add_connector(
            MSO_CONNECTOR.ELBOW,
            Inches(x + rw),
            Inches(ry + rh / 2),
            Inches(cx),
            Inches(cy + ch / 2),
        )
        conn.line.color.rgb = _rgb(c["accent"])
        conn.line.width = Pt(1.75)
        _arrowhead(conn)


def _timeline(slide, nodes, brand, x, y, w, h):
    """수평 타임라인 — 축선 + 넘버 마일스톤 원 + 상단 라벨/하단 설명 (로드맵용)."""
    c, f, s = brand["colors"], brand["fonts"], brand["sizes"]
    n = len(nodes)
    line_y = y + h * 0.45
    axis = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x + 0.3), Inches(line_y), Inches(x + w - 0.3), Inches(line_y)
    )
    axis.line.color.rgb = _rgb(c["primary"])
    axis.line.width = Pt(2.0)
    _arrowhead(axis)
    step = (w - 1.4) / max(n - 1, 1)
    for i, nd in enumerate(nodes):
        cx = x + 0.7 + i * step
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(cx - 0.19), Inches(line_y - 0.19), Inches(0.38), Inches(0.38)
        )
        dot.fill.solid()
        dot.fill.fore_color.rgb = _rgb(c["accent"])
        dot.line.color.rgb = _rgb(c["bg"])
        dot.line.width = Pt(1.5)
        tf = dot.text_frame
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(i + 1)
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.name = f["head"]
        run.font.color.rgb = _rgb("FFFFFF")
        # 상단 라벨(굵게) / 하단 설명(작게) — 박스를 본문 영역 안으로 클램프(양끝 잘림 방지)
        bx = max(x, cx - step / 2 + 0.1)
        bw = min(x + w, cx + step / 2 - 0.1) - bx
        lb = slide.shapes.add_textbox(Inches(bx), Inches(line_y - 0.85), Inches(bw), Inches(0.5))
        p1 = lb.text_frame.paragraphs[0]
        lb.text_frame.word_wrap = True
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = nd["label"]
        r1.font.size = Pt(s["body"])
        r1.font.bold = True
        r1.font.name = f["head"]
        r1.font.color.rgb = _rgb(c["primary"])
        if nd.get("sub"):
            sb = slide.shapes.add_textbox(
                Inches(bx), Inches(line_y + 0.32), Inches(bw), Inches(0.8)
            )
            sb.text_frame.word_wrap = True
            p2 = sb.text_frame.paragraphs[0]
            p2.alignment = PP_ALIGN.CENTER
            r2 = p2.add_run()
            r2.text = nd["sub"]
            r2.font.size = Pt(s["body"] - 2)
            r2.font.name = f["body"]
            r2.font.color.rgb = _rgb(c["muted"])


def _flow(slide, nodes, brand, x, y, w, h):
    """노드 좌→우 균등 배치 + 사이마다 accent 화살표 (프로세스 흐름)."""
    n = len(nodes)
    gap = 0.5
    bw = (w - gap * (n - 1)) / n
    bh = min(1.4, h)
    by = y + (h - bh) / 2
    for i, nd in enumerate(nodes):
        bx = x + i * (bw + gap)
        _node_box(slide, brand, bx, by, bw, bh, nd["label"], nd.get("sub"))
        if i:
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                Inches(bx - gap + 0.08),
                Inches(by + bh / 2 - 0.11),
                Inches(gap - 0.16),
                Inches(0.22),
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = _rgb(brand["colors"]["accent"])
            arrow.line.fill.background()


def _cards(slide, nodes, brand, x, y, w, h):
    """그리드 카드 — 순서·위계 없는 병렬 요소용(화살표·적층 없음). 3~6개 권장."""
    n = len(nodes)
    cols = 2 if n <= 4 else 3
    rows = (n + cols - 1) // cols
    gap = 0.3
    cw = (w - gap * (cols - 1)) / cols
    ch = min(1.7, (h - gap * (rows - 1)) / rows)
    for i, nd in enumerate(nodes):
        cx = x + (i % cols) * (cw + gap)
        cy = y + (i // cols) * (ch + gap)
        _node_box(slide, brand, cx, cy, cw, ch, nd["label"], nd.get("sub"))


def _layers(slide, nodes, brand, x, y, w, h):
    """노드 상→하 적층 (아키텍처 레이어). 줄무늬 배경으로 층 구분."""
    c = brand["colors"]
    n = len(nodes)
    gap = 0.18
    bh = min(1.0, (h - gap * (n - 1)) / n)
    for i, nd in enumerate(nodes):
        _node_box(
            slide,
            brand,
            x,
            y + i * (bh + gap),
            w,
            bh,
            nd["label"],
            nd.get("sub"),
            fill=c["bg_alt"] if i % 2 == 0 else c["bg"],
        )
