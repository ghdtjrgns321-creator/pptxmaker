"""S8 시안 2종 — 기술 1: 용어사전 (구축→적용→실증 3구획 휴리스틱, UI 패스 — 문구 러프 허용).

재료: 00_factsheet.md §C 구축·작동 상세 + §D (원문 2_DATA-TAXONOMY §2.6).
시안 A(좌앵커형): 좌측 앵커(대형 423) + 우측 3구획 세로 스택(구축 플로우·적용 예시·분해 바).
시안 B(풀폭 플로우형): 상단 구축 셰브런 밴드 + 중단 적용 예시·다중 목적지 + 하단 분해 바.
실행: uv run python golden/s08_variants.py → golden/variants/s08_variants.pptx
"""

from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from . import grid as G
from ._variant_h import _shape_step
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit, mix, set_shape_text

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]

KICKER = "3. 기술 설명 — TECH 01 · Analyze"
SOURCE = "출처: 2_DATA-TAXONOMY.md §2.6 · 4_SEARCH-PIPELINE.md (00_factsheet.md §C·§D)"
OVERLAP = G.FLOW_H * 0.55 / 2

BREAKDOWN = [  # 등재 423 분해 (합 423 검산은 audit)
    ("자동", 316, C["primary"]),
    ("위임판단", 86, mix(C["primary"], C["muted"], 0.5)),
    ("사용자검토", 18, C["muted"]),
    ("사용자확정", 1, C["accent"]),
    ("제외", 2, mix(C["muted"], C["bg"], 0.5)),
]


def header(slide, headline):
    add_text(
        slide, G.MARGIN_L, 0.42, 8.0, 0.28, KICKER, S["caption"], F["head"], C["muted"], bold=True
    )
    add_text(
        slide,
        G.MARGIN_L,
        0.72,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.55,
        headline,
        S["section"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_box(slide, G.MARGIN_L, G.RULE_Y, G.RIGHT_EDGE - G.MARGIN_L, 0.014, fill=C["muted"])


def _subhead(slide, x, y, w, text):
    add_text(slide, x, y, w, 0.32, text, S["head"], F["head"], C["primary"], bold=True)
    add_box(slide, x, y + 0.35, w, 0.012, fill=C["muted"])


def _arrow(slide, x1, y1, x2, y2):
    from pptx.oxml.ns import qn

    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    conn.line.color.rgb = C["muted"]
    conn.line.width = Pt(1.5)
    ln = conn.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))


def bar_and_source(slide, text, right=None):
    r = right or G.RIGHT_EDGE
    bar = add_box(slide, G.MARGIN_L, G.BAR_Y, r - G.MARGIN_L, G.BAR_H, fill=C["primary"])
    tf = bar.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name, run.font.bold = F["head"], True
    run.font.size = Pt(S["body"])
    run.font.color.rgb = C["bg"]
    add_text(
        slide, G.MARGIN_L, G.SOURCE_Y, r - G.MARGIN_L, 0.3, SOURCE, S["foot"], F["body"], C["muted"]
    )


def breakdown_bar(slide, x, y, w, h):
    """등재 423 비례 스택 바 + 범례. 미세 구간은 최소 가시폭(주석으로 정직 표기)."""
    total = sum(n for _, n, _ in BREAKDOWN)
    min_w = 0.06
    seg_x = x
    for name, n, color in BREAKDOWN:
        seg_w = max(w * n / total, min_w)
        add_box(slide, seg_x, y, seg_w, h, fill=color)
        seg_x += seg_w
    leg_w = w / len(BREAKDOWN)
    for i, (name, n, color) in enumerate(BREAKDOWN):
        leg_x = x + i * leg_w
        add_box(slide, leg_x, y + h + 0.1, 0.13, 0.13, fill=color)
        add_text(
            slide,
            leg_x + 0.19,
            y + h + 0.06,
            leg_w - 0.2,
            0.22,
            f"{name} {n}",
            S["caption"],
            F["body"],
            C["muted"],
        )


def _apply_example(slide, x, y):
    """적용 예시 — 질문 속 문자가 개념 후보로 이어지는 그림."""
    q = add_box(slide, x, y, 3.4, 0.6, fill=C["bg"], line=C["muted"], line_w=1.0, shape="round")
    tf = q.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run()
    r1.text = "“리베이트 조항이 있는 계약은…”"
    r1.font.name, r1.font.size, r1.font.color.rgb = F["body"], Pt(S["caption"]), C["primary"]
    _arrow(slide, x + 3.4, y + 0.3, x + 4.0, y + 0.3)
    chip = add_box(slide, x + 4.0, y + 0.12, 1.5, 0.36, fill=C["primary"], shape="round")
    set_shape_text(chip, "변동대가", S["caption"], F["head"], C["bg"], bold=True)
    add_text(
        slide,
        x + 5.7,
        y + 0.16,
        2.6,
        0.3,
        "후보 진입점 — 확정 라우터 아님",
        S["caption"],
        F["body"],
        C["muted"],
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        slide,
        x,
        y + 0.75,
        8.2,
        0.26,
        "질문 문장에 등재 용어가 문자 그대로 나타나면 잡는다(substring) — 유사도 점수 없음. 최종 선택은 질문 문맥이 담당.",
        S["caption"],
        F["body"],
        C["muted"],
    )


def variant_a(prs):
    """A — 좌앵커(대형 423) + 우측 3구획 스택(구축 → 적용 → 실증)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    header(slide, "용어사전 — 사람이 검수한 진입 색인, AI 신규 창작은 0")
    lx, lw = G.COL_L_X, G.COL_L_W
    add_text(
        slide,
        lx,
        G.SUBHEAD_Y,
        lw,
        0.26,
        "진입 색인",
        S["caption"],
        F["head"],
        C["muted"],
        bold=True,
    )
    add_text(slide, lx, 2.08, lw, 0.7, "423", S["display"], F["head"], C["accent"], bold=True)
    add_text(slide, lx, 2.78, lw, 0.48, "등재 용어", S["title"], F["head"], C["primary"], bold=True)
    add_text(
        slide,
        lx,
        3.42,
        lw,
        1.5,
        "원천은 사람이 만든 1차 자료 3종뿐이다. AI는 초안만 내고 사람이 전수 검수했다 — 색인이 틀리면 '못 찾음'으로 안전하게 드러난다.",
        S["sub"],
        F["body"],
        C["text"],
        line_spacing=1.3,
    )
    add_text(
        slide,
        lx,
        5.2,
        lw,
        1.0,
        "뼈대(개념·간선)는 기준서에서 기계 생성 — AI 개입을 색인 하나로 한정한 신빙성 배치.",
        S["caption"],
        F["body"],
        C["muted"],
        line_spacing=1.25,
    )
    add_box(
        slide, G.V_RULE_X, G.CONTENT_TOP, 0.014, G.CONTENT_BOTTOM - G.CONTENT_TOP, fill=C["bg_alt"]
    )

    rx, rw = G.COL_R_X, G.COL_R_W
    _subhead(slide, rx, G.SUBHEAD_Y, rw, "구축 — 3원천에서 사람 전수 검수로")
    srcs = [("질의 매핑 288", 0), ("사례 제목 123", 1), ("부록A 정의 9", 2)]
    for name, i in srcs:
        b = add_box(slide, rx, 2.28 + i * 0.42, 1.9, 0.34, fill=C["bg_alt"], shape="round")
        set_shape_text(b, name, S["caption"], F["body"], C["primary"])
    steps = [("AI 초안", 6.6), ("사람 전수 검수", 8.6)]
    for name, sx in steps:
        b = add_box(slide, sx, 2.7, 1.7, 0.5, fill=C["bg_alt"], shape="round")
        set_shape_text(b, name, S["caption"], F["head"], C["primary"], bold=True)
    fin = add_box(slide, 10.9, 2.7, 1.83, 0.5, fill=C["primary"], shape="round")
    set_shape_text(fin, "등재 423", S["body"], F["head"], C["bg"], bold=True)
    for i in range(3):
        _arrow(slide, rx + 1.9, 2.45 + i * 0.42, 6.6, 2.95)
    _arrow(slide, 8.3, 2.95, 8.6, 2.95)
    _arrow(slide, 10.3, 2.95, 10.9, 2.95)
    _subhead(slide, rx, 3.75, rw, "적용 — 질문 속 문자를 그대로 잡는다")
    _apply_example(slide, rx, 4.3)
    _subhead(slide, rx, G.BOX_Y + 0.05, rw, "실증 — 등재 423의 결정 분해")
    breakdown_bar(slide, rx, G.BOX_Y + 0.45, rw, 0.24)
    bar_and_source(
        slide, "AI가 만드는 것은 색인 하나 — 틀려도 1종(놓침)으로 드러나는 자리에만 둔다"
    )
    return slide


def variant_b(prs):
    """B — 상단 풀폭 구축 셰브런 밴드 + 중단 적용·다중 목적지 + 하단 분해 바."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    header(slide, "용어사전 — 3원천, 사람 전수 검수, 신규 창작 0")
    add_text(
        slide,
        G.MARGIN_L,
        G.CONTENT_TOP,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.3,
        "질문이 그래프에 진입하는 첫 관문 — 사람이 만든 자료만으로 등재했다.",
        S["sub"],
        F["body"],
        C["text"],
    )
    band_y = 2.35
    step_w = (G.RIGHT_EDGE - G.MARGIN_L + 3 * OVERLAP) / 4
    band = [
        ("1차 자료 3종", "사람 원천"),
        ("AI 초안", "제안만"),
        ("사람 전수 검수", "애매 3건 판정"),
        ("등재 423", "창작 0"),
    ]
    for i, (head, desc) in enumerate(band):
        _shape_step(
            slide,
            G.MARGIN_L + i * (step_w - OVERLAP),
            band_y,
            step_w,
            G.FLOW_H,
            "pentagon" if i == 0 else "chevron",
            C["bg_alt"] if i < 3 else C["accent"],
            head,
            desc,
            C["primary"] if i < 3 else C["bg"],
            mix(C["primary"], C["bg"], 0.35) if i < 3 else C["bg_alt"],
        )
    caps = [
        "질의 매핑 288 · 사례 제목 123 · 부록A 정의 9",
        "후보 제안까지만 — 확정 권한 없음",
        "전건 결정 로그(누가·왜) 기록",
        "자동 316 · 위임 86 · 검토 18 · 확정 1 · 제외 2",
    ]
    for i, cap in enumerate(caps):
        add_text(
            slide,
            G.MARGIN_L + i * (step_w - OVERLAP) + 0.1,
            band_y + G.FLOW_H + 0.12,
            step_w - 0.5,
            0.55,
            cap,
            S["caption"],
            F["body"],
            C["muted"],
            line_spacing=1.2,
        )
    _subhead(slide, G.MARGIN_L, 4.15, 8.0, "적용 — substring 진입")
    _apply_example(slide, G.MARGIN_L, 4.7)
    _subhead(slide, 9.0, 4.15, G.RIGHT_EDGE - 9.0, "다중 목적지")
    term = add_box(slide, 9.0, 4.75, 1.4, 0.4, fill=C["bg_alt"], shape="round")
    set_shape_text(term, "“계약”", S["caption"], F["head"], C["primary"], bold=True)
    for i, concept in enumerate(["계약 식별", "계약변경", "계약원가"]):
        cy = 4.55 + i * 0.5
        _arrow(slide, 10.4, 4.95, 10.95, cy + 0.2)
        cb = add_box(
            slide, 10.95, cy, 1.75, 0.4, fill=C["bg"], line=C["muted"], line_w=0.75, shape="round"
        )
        set_shape_text(cb, concept, S["caption"], F["body"], C["primary"])
    add_text(
        slide,
        9.0,
        6.1,
        G.RIGHT_EDGE - 9.0,
        0.24,
        "한 용어 → 여러 개념 후보",
        S["caption"],
        F["body"],
        C["muted"],
    )
    bar_and_source(
        slide, "AI가 만드는 것은 색인 하나 — 틀려도 1종(놓침)으로 드러나는 자리에만 둔다"
    )
    return slide


# 실물 엔트리 (data/ontology/aliases.json에서 발췌 — 축약만, 창작 0)
TABLE_ROWS = [
    ("용어", "원천", "등급", "연결 개념 (어디로 진입하나)"),
    ("리베이트", "질의 매핑", "자동", "고객에게 지급할 대가"),
    ("밀어내기", "질의 매핑", "자동", "위탁약정"),
    ("볼륨디스카운트", "질의 매핑", "자동", "변동대가 · 변동대가 추정치를 제약함"),
    ("상품권", "질의 매핑", "위임판단", "고객에게 지급할 대가 · 고객이 행사하지 아니한 권리"),
    ("반품의 회계처리", "QNA 제목", "자동", "반품권이 있는 판매 · 본인 대 대리인 (QNA-SSI-38695)"),
]
JSON_LINES = [  # 상품권 엔트리 실물 축약
    '{ "term": "상품권",',
    '  "sources": ["query-mapping"],',
    '  "grade": "자동(위임판단)",',
    '  "concepts": [',
    '    "고객에게 지급할 대가",',
    '    "고객이 행사하지 아니한 권리" ],',
    '  "decision": {',
    '    "by": "AI 위임 판단",',
    '    "reason": "미행사 상품권 =',
    '        B44~47 정면 조항" } }',
]


def variant_c(prs):
    """C3 — 서사 3칼럼(왜/역할/구축) + 실물 증거(네이티브 표 + JSON 카드). 신규 어휘."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    header(slide, "용어사전 — 실무 언어를 기준서 개념에 잇는 진입 색인")
    # ── 상단: 왜 / 역할 / 구축 3칼럼 (두괄식 볼드 헤드 + 완전 문장) ──
    nar_w = (G.RIGHT_EDGE - G.MARGIN_L - 2 * 0.4) / 3
    narratives = [
        (
            "왜 필요한가",
            "실무는 '리베이트'라 말하고 기준서는 '고객에게 지급할 대가'라 쓴다. 이 언어 간극을 잇지 않으면 검색이 시작조차 되지 않는다.",
        ),
        (
            "파이프라인에서의 역할",
            "Analyze 단계의 진입 관문. 질문에 등재 용어가 문자 그대로 나타나면 연결 개념으로 그래프에 진입한다 — 유사도 점수 없이, 사람이 승인한 경로로만.",
        ),
        (
            "어떻게 만들었나",
            "사람 1차 자료 3종(질의 매핑 288 · 사례 제목 123 · 부록A 정의 9)에서 AI가 초안을 내고 사람이 전수 검수해 423개 등재 — 신규 창작 0건.",
        ),
    ]
    for i, (head, body) in enumerate(narratives):
        nx = G.MARGIN_L + i * (nar_w + 0.4)
        add_text(
            slide,
            nx,
            G.CONTENT_TOP,
            nar_w,
            0.28,
            f"0{i + 1}  {head}",
            S["head"],
            F["head"],
            C["primary"],
            bold=True,
        )
        add_text(
            slide,
            nx,
            G.CONTENT_TOP + 0.36,
            nar_w,
            0.95,
            body,
            S["caption"],
            F["body"],
            C["muted"],
            line_spacing=1.25,
        )
        if i < 2:
            add_box(slide, nx + nar_w + 0.2, G.CONTENT_TOP, 0.012, 1.2, fill=C["bg_alt"])
    # ── 하단 좌: 네이티브 표 (실물 5행) ──
    tbl_x, tbl_y, tbl_w = G.MARGIN_L, 3.35, 7.7
    n_rows = len(TABLE_ROWS)
    gf = slide.shapes.add_table(
        n_rows, 4, Inches(tbl_x), Inches(tbl_y), Inches(tbl_w), Inches(0.5 * n_rows)
    )
    tbl = gf.table
    for ci, cw in enumerate((1.65, 1.15, 1.15, 3.75)):
        tbl.columns[ci].width = Inches(cw)
    for ri, row in enumerate(TABLE_ROWS):
        tbl.rows[ri].height = Inches(0.42 if ri else 0.38)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = (
                C["primary"] if ri == 0 else (C["bg_alt"] if ri % 2 == 0 else C["bg"])
            )
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p._p.get_or_add_pPr().set("eaLnBrk", "0")
            r = p.add_run()
            r.text = val
            r.font.name = F["head"] if ri == 0 or ci == 0 else F["body"]
            r.font.size = Pt(S["caption"])
            r.font.bold = ri == 0 or ci == 0
            r.font.color.rgb = C["bg"] if ri == 0 else C["primary"]
    add_text(
        slide,
        tbl_x,
        6.0,
        tbl_w,
        0.26,
        "등재 423 중 발췌 5건 — 자동 316 · 위임판단 86 · 검토 18 · 확정 1 · 제외 2",
        S["caption"],
        F["body"],
        C["muted"],
    )
    # ── 하단 우: 실물 JSON 카드 (모노스페이스, 다크) ──
    jx, jy, jw, jh = 8.7, 3.35, G.RIGHT_EDGE - 8.7, 2.55
    card = add_box(slide, jx, jy, jw, jh, fill=C["primary"], shape="round")
    try:
        card.adjustments[0] = 0.06  # 라운드 과대 방지
    except (IndexError, ValueError):
        pass
    add_text(
        slide,
        jx + 0.25,
        jy + 0.15,
        jw - 0.5,
        0.24,
        "aliases.json — 실제 엔트리",
        S["caption"],
        F["head"],
        C["bg_alt"],
        bold=True,
    )
    tb = add_text(
        slide,
        jx + 0.25,
        jy + 0.5,
        jw - 0.5,
        jh - 0.7,
        [[(ln, {})] for ln in JSON_LINES],
        S["foot"],
        "Consolas",
        C["bg"],
        line_spacing=1.25,
    )
    add_text(
        slide,
        jx,
        6.0,
        jw,
        0.26,
        "모든 엔트리가 결정 로그(누가·왜)를 갖는다 — 전건 추적.",
        S["caption"],
        F["body"],
        C["muted"],
    )
    bar_and_source(
        slide, "AI가 만드는 것은 색인 하나 — 틀려도 1종(놓침)으로 드러나는 자리에만 둔다"
    )
    return slide


def audit(prs):
    EMU = 914400
    accent_hex = str(C["accent"])
    assert sum(n for _, n, _ in BREAKDOWN) == 423, "분해 합 검산 실패"
    fails = []
    for si, sl in enumerate(prs.slides):
        rules, solids, acc = [], [], 0
        for sh in sl.shapes:
            L, T, W, H = sh.left / EMU, sh.top / EMU, sh.width / EMU, sh.height / EMU
            if L == 0 and T == 0 and W > 13:
                continue
            if H <= 0.02 and W > 1:
                rules.append((L, T, W))
            elif H > 0.3 and W > 0.5 and sh.shape_type != 17:
                solids.append((L, T, W, H))
            if T + H > G.CONTENT_BOTTOM + 0.02 and T < G.BAR_Y - 0.02:
                fails.append((si, "bottom>6.35", round(T + H, 3)))
            if 11.5 < W < 13 and abs((L + W) - G.RIGHT_EDGE) > 0.02:
                fails.append((si, "풀폭 right≠12.733", round(L + W, 3)))
            try:
                if str(sh.fill.fore_color.rgb) == accent_hex:
                    acc += 1
            except Exception:
                pass
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if (
                            r.font.color
                            and r.font.color.rgb is not None
                            and str(r.font.color.rgb) == accent_hex
                        ):
                            acc += 1
        for rl, rt, rw in rules:
            for sl_, st, sw, shh in solids:
                if rl < sl_ + sw and rl + rw > sl_ and st - 0.005 <= rt <= st + shh - 0.005:
                    fails.append((si, "룰-채움 겹침", round(rt, 3)))
        if acc > 4:
            fails.append((si, "accent 초과", acc))
        print(f"slide{si}: accent {acc}곳")
    assert not fails, f"AUDIT FAIL {fails}"
    print("audit pass")


def main():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    variant_a(prs)
    variant_b(prs)
    variant_c(prs)
    audit(prs)
    from pathlib import Path

    out = Path(__file__).parent / "variants" / "s08_variants.pptx"
    prs.save(str(out))
    print(f"saved {out}")


if __name__ == "__main__":
    main()
