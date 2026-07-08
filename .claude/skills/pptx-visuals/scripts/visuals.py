"""차트·다이어그램 렌더 모듈 — pptx-visuals 스킬의 단일 출처.

build_pptx.py가 import해서 사용한다. 차트는 네이티브 차트(PPT에서 데이터 수정 가능),
다이어그램은 도형+커넥터 조합(flow/layers)으로 렌더한다.
색·폰트·크기는 인자로 받은 brand dict(brand-kit.yaml)에서만 나온다.
"""

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


def _rgb(hex6):
    return RGBColor.from_string(hex6)


CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
    "pie": XL_CHART_TYPE.PIE,
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
    ch.has_legend = len(spec["series"]) > 1 or ct == XL_CHART_TYPE.PIE
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
    _apply_series_colors(ch, spec.get("chart_type", "bar"), brand)
    return gf


def _apply_series_colors(chart, kind, brand):
    """Office 기본 테마색 대신 brand 팔레트(accent→primary→muted 순환)를 입힌다."""
    c = brand["colors"]
    palette = [c["accent"], c["primary"], c["muted"]]
    if kind == "pie":
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
    """diagram 슬라이드 본문 렌더. layout: flow(좌→우 화살표) | layers(상→하 적층)."""
    layout = spec.get("layout", "flow")
    if layout == "layers":
        _layers(slide, spec["nodes"], brand, x, y, w, h)
    elif layout == "flow":
        _flow(slide, spec["nodes"], brand, x, y, w, h)
    else:
        raise ValueError(f"unknown diagram layout: {layout}")


def _node_box(slide, brand, x, y, w, h, label, sub=None, fill=None):
    c, f, s = brand["colors"], brand["fonts"], brand["sizes"]
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
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
    run.font.color.rgb = _rgb(c["primary"])
    if sub:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run2 = p2.add_run()
        run2.text = sub
        run2.font.size = Pt(s["body"] - 2)
        run2.font.name = f["body"]
        run2.font.color.rgb = _rgb(c["muted"])
    return box


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
