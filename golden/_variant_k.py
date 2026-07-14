# 시안 K — 2칸 재설계(사용자 지시 2026-07-14): 좌 오류 두 방향(1종/2종) · 우 같은 질문 두 결말
# golden/ 동결 원본 — goldenfab/_variant_k(c=None)와 도형 동일해야 함(compare_golden 기준).
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

import grid as G
from kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit, mix

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]

DEFAULT = {
    "kicker": "1. 문제 정의",
    "headline": "일반 LLM은 틀리는 게 아니라, 틀리는 방향이 문제다",
    "right1_head": "같은 질문, 두 개의 결말",
    "question": '"이 계약, 수익을 지금 인식해도 됩니까?"',
    "lanes": [
        {
            "name": "일반 LLM",
            "steps": [("즉시 답변", ""), ("확신형 어조", ""), ("검증 불가", "")],
            "verdict": "치명",
        },
        {
            "name": "본 시스템",
            "steps": [("근거 먼저 검색", ""), ("확정", ""), ("유보", "")],
            "verdict": "안전",
        },
    ],
    "right2_head": "오류의 두 방향 — 왜 2종만 치명적인가",
    "panels": [
        {
            "head": "1종 · 놓침 (false negative)",
            "desc": '근거가 존재하는데도 찾지 못해 "모른다"고 물러선다. 사용자는 다른 방법을 찾을 뿐 — 손실은 시간, 위험도는 낮다. 추가 검색으로 보완 가능하다.',
        },
        {
            "head": "2종 · 허위 확정 (false positive)",
            "desc": "근거가 없거나 틀렸는데 확신에 찬 결론을 준다. 그대로 재무제표에 반영되면 왜곡 — 치명적. 이 시스템이 구조로 겨냥하는 오류다.",
        },
    ],
    "bar": "확실할 때만 확정하고, 애매하면 근거를 보여주며 유보한다",
    "source": "출처: 1_OVERVIEW.md · PROJECT_OVERVIEW.md · 5_INTERFACE.md (00_factsheet.md §A·§B)",
}


def _col_head(slide, x, w, y, text):
    add_text(slide, x, y, w, 0.32, text, S["head"], F["head"], C["primary"], bold=True)
    add_box(slide, x, y + 0.36, w, 0.012, fill=C["muted"])


def _center(box, text, size, color, bold=True):
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name, r.font.bold = F["head"], bold
    r.font.size = Pt(size)
    r.font.color.rgb = color


def variant_k(prs):
    """2칸 — 좌: 오류 두 방향(1종/2종 비교), 우: 같은 질문 두 결말. 골든 기준(고정 내용)."""
    c = DEFAULT
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    add_text(
        slide,
        G.MARGIN_L,
        0.42,
        6.0,
        0.28,
        c["kicker"],
        S["caption"],
        F["head"],
        C["muted"],
        bold=True,
    )
    add_text(
        slide,
        G.MARGIN_L,
        0.72,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.55,
        c["headline"],
        S["section"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_box(slide, G.MARGIN_L, G.RULE_Y, G.RIGHT_EDGE - G.MARGIN_L, 0.014, fill=C["muted"])

    gap = 0.6
    col_w = (G.RIGHT_EDGE - G.MARGIN_L - gap) / 2
    lx = G.MARGIN_L
    rx = G.MARGIN_L + col_w + gap
    top = G.CONTENT_TOP
    add_box(
        slide, lx + col_w + gap / 2 - 0.007, top, 0.014, G.CONTENT_BOTTOM - top, fill=C["bg_alt"]
    )

    _col_head(slide, lx, col_w, top, c["right2_head"])
    pt = top + 0.62
    ph = (G.CONTENT_BOTTOM - pt - 0.25) / 2
    panel_style = [
        (C["bg_alt"], mix(C["muted"], C["bg"], 0.3), C["primary"], C["text"]),
        (C["primary"], C["accent"], C["bg"], C["bg_alt"]),
    ]
    for i, ((fill, spine, head_ink, desc_ink), panel) in enumerate(zip(panel_style, c["panels"])):
        py = pt + i * (ph + 0.25)
        add_box(slide, lx, py, col_w, ph, fill=fill, shape="round")
        add_box(slide, lx, py, 0.05, ph, fill=spine)
        add_text(
            slide,
            lx + 0.28,
            py + 0.2,
            col_w - 0.55,
            0.3,
            panel["head"],
            S["body"],
            F["head"],
            head_ink,
            bold=True,
        )
        add_text(
            slide,
            lx + 0.28,
            py + 0.64,
            col_w - 0.55,
            ph - 0.8,
            panel["desc"],
            S["caption"],
            F["body"],
            desc_ink,
            line_spacing=1.3,
        )

    _col_head(slide, rx, col_w, top, c["right1_head"])
    qt = top + 0.62
    qh = 0.72
    q = add_box(
        slide, rx, qt, col_w, qh, fill=C["bg_alt"], line=C["muted"], line_w=0.75, shape="round"
    )
    _center(q, c["question"].replace("\n", " "), S["body"], C["primary"])
    ot = qt + qh + 0.3
    oh = (G.CONTENT_BOTTOM - ot - 0.25) / 2
    lane_ink = [C["accent"], C["primary"]]
    chip_w = 1.3
    for i, (lane, vink) in enumerate(zip(c["lanes"], lane_ink)):
        oy = ot + i * (oh + 0.25)
        add_box(slide, rx, oy, col_w, oh, fill=C["bg_alt"], shape="round")
        chip = add_box(slide, rx + col_w - chip_w, oy, chip_w, oh, fill=vink, shape="round")
        _center(chip, lane["verdict"], S["title"], C["bg"])
        add_text(
            slide,
            rx + 0.28,
            oy + 0.18,
            col_w - chip_w - 0.5,
            0.3,
            lane["name"],
            S["body"],
            F["head"],
            C["primary"],
            bold=True,
        )
        steps_txt = "  ›  ".join(h for h, _ in lane["steps"])
        add_text(
            slide,
            rx + 0.28,
            oy + 0.6,
            col_w - chip_w - 0.5,
            oh - 0.75,
            steps_txt,
            S["caption"],
            F["body"],
            C["muted"],
            line_spacing=1.25,
        )

    bar = add_box(slide, G.MARGIN_L, G.BAR_Y, G.RIGHT_EDGE - G.MARGIN_L, G.BAR_H, fill=C["primary"])
    _center(bar, c["bar"], S["body"], C["bg"])
    add_text(
        slide,
        G.MARGIN_L,
        G.SOURCE_Y,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.3,
        c["source"],
        S["foot"],
        F["body"],
        C["muted"],
    )
