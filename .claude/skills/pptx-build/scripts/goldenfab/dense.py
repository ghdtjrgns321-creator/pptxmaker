"""goldenfab.dense — s06 고밀도 디자인 언어 공유 헬퍼 (2026-07-20).

s06_mid에서 확정한 규격을 골든 전 장이 공유한다(handoff 문법 승격):
- compact_header : 키커 + 14pt 헤드라인 + 굵은 룰 (골든 28pt 헤더·리드 문단 폐지)
- hero_card      : 배지·제목·태그·중앙 배너·히어로 아이콘·2단 불릿·아이콘 칩 (s06 카드)
                   — **카드는 이 컴포넌트 하나만.** 장별 경량 변형 금지(design-rules §8).
- chip_row / two_col_bullets / icon / hrule / source_line

원칙: **하단 전폭 한줄평/재서술 바 금지(형태 자체)** — 색 불문. 결론은 콘텐츠에 흡수.
Ember accent 소량, 전 텍스트 8~15pt(헤드라인 14 예외). 자세한 차단 근거는 design-rules §8.
"""

from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from . import grid as G
from .kit import ROOT, add_box, add_text, load_kit, mix, set_shape_text

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]
ICONS = ROOT / "assets/icons"

FULL_W = G.RIGHT_EDGE - G.MARGIN_L
RULE_Y = 0.66
SOURCE_Y = 7.14

line_soft = mix(C["primary"], C["bg"], 0.80)
line_mid = mix(C["primary"], C["bg"], 0.55)
ink_body = mix(C["muted"], C["primary"], 0.45)


def icon(slide, name, variant, x, y, size):
    p = ICONS / f"{name}_{variant}.png"
    return slide.shapes.add_picture(str(p), Inches(x), Inches(y), Inches(size), Inches(size))


def hrule(slide, x, y, w, color=None, weight=0.75):
    return add_box(slide, x, y, w, 0.001, fill=None, line=color or line_soft, line_w=weight)


def arrow(slide, x1, y1, x2, y2, color, w=1.5):
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    conn.line.color.rgb = color
    conn.line.width = Pt(w)
    ln = conn.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return conn


BAND_RULE_DY = 0.36


def band_head(slide, y, text):
    """밴드 소제목 + 풀폭 룰 — 한 장 안의 구획을 여백이 아니라 장치로 나눈다.

    페이지 요소다(도해가 아니다). 2층 부품은 이걸 소유하지 않는다 — 부품은 box 안 그래픽
    하나만 그리고, 구획을 나눌지는 장이 정한다.
    """
    w = G.RIGHT_EDGE - G.MARGIN_L
    add_text(slide, G.MARGIN_L, y, w, 0.32, text, S["head"], F["head"], C["primary"], bold=True)
    add_box(slide, G.MARGIN_L, y + BAND_RULE_DY, w, 0.012, fill=C["muted"])


def compact_header(slide, kicker, headline):
    """s06 압축 헤더 — 키커(8pt) + 헤드라인(14pt) + 굵은 룰. 상단 의식 ~0.7\"."""
    add_text(
        slide, G.MARGIN_L, 0.14, 8.0, 0.16, kicker, S["foot"], F["head"], C["muted"], bold=True
    )
    add_text(
        slide,
        G.MARGIN_L,
        0.32,
        FULL_W,
        0.28,
        headline,
        S["sub"],
        F["head"],
        C["primary"],
        bold=True,
    )
    hrule(slide, G.MARGIN_L, RULE_Y, FULL_W, color=C["primary"], weight=1.4)


def source_line(slide, text):
    add_text(slide, G.MARGIN_L, SOURCE_Y, FULL_W, 0.2, text, S["foot"], F["body"], C["muted"])


# [차단 2026-07-20] 하단 전폭 '한줄평/재서술 바'는 형태 자체가 금지다(handoff 결함 #2).
# 검은색이든 라이트든 accent 좌변이든 — 슬라이드 맨 밑 전폭 한 줄 결론은 만들지 않는다.
# 결론 문장은 콘텐츠(불릿·배너)에 흡수한다. takeaway_strip 헬퍼는 그래서 폐기했다.


def chip_row(slide, x, y, w, ic, em, label, *, h=0.23):
    """아이콘 칩 — 아이콘 + 강조(accent) + 라벨(muted)."""
    chip = add_box(slide, x, y, w, h, fill=C["bg_alt"])
    tf = chip.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.30)
    p = tf.paragraphs[0]
    p._p.get_or_add_pPr().set("eaLnBrk", "0")
    p.alignment = PP_ALIGN.LEFT
    for txt, color, bold in ((em, C["accent"], True), (label, C["muted"], False)):
        if not txt:
            continue
        r = p.add_run()
        r.text = txt
        r.font.name = F["head"]
        r.font.size = Pt(S["caption"])
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.language_id = MSO_LANGUAGE_ID.KOREAN
    icon(slide, ic, "primary", x + 0.08, y + (h - 0.17) / 2, 0.17)


def two_col_bullets(slide, x, y, w, items, *, step=0.29, h=0.28):
    """2단 불릿 — items: [(리드, 설명)]. 볼드 리드 + 설명(서술형 금지)."""
    for j, (lead, detail) in enumerate(items):
        add_text(
            slide,
            x,
            y + j * step,
            w,
            h,
            [
                [
                    ("· ", {"bold": True, "color": C["accent"]}),
                    (lead, {"bold": True, "color": C["primary"]}),
                    (detail, {}),
                ]
            ],
            S["caption"],
            F["body"],
            ink_body,
            anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.0,
        )


# [차단 2026-07-20] narrative_card(경량 카드) 폐기. 카드는 hero_card 하나만 쓴다 —
# 장별 경량 변형이 s06과 어긋나 "쓰레기" 카드를 낳았다(사용자 반려). 카드 = hero_card 단일.


def hero_card(slide, x, y, w, h, num, title, tag, ic, banner, items, chips, *, hero=0.48):
    """s06 표준 히어로 카드 — 골든 전 장 공용 단일 컴포넌트.

    배지→제목→태그→배너→히어로 아이콘(상단 고정) · 룰+칩(하단 고정) · 불릿(가운데 채움).
    위치를 콘텐츠·h에서 **파생**한다 — 불릿 수·카드 높이가 달라도 s06 룩 유지 + 죽은 여백 0.
    (h≥3.2 권장. 이 함수만 카드를 그린다 — 장별 즉흥 카드 금지, design-rules §8.)"""
    add_box(slide, x, y, w, h, fill=C["bg"], line=line_soft, line_w=1.0)
    badge = add_box(slide, x + 0.14, y + 0.13, 0.28, 0.28, fill=C["primary"], shape="oval")
    set_shape_text(badge, num, S["foot"], F["head"], C["bg"], bold=True)
    add_text(
        slide,
        x + 0.52,
        y + 0.11,
        w - 0.66,
        0.32,
        title,
        S["head"],
        F["head"],
        C["primary"],
        bold=True,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        slide,
        x + 0.14,
        y + 0.50,
        w - 0.28,
        0.20,
        tag,
        S["caption"],
        F["head"],
        C["accent"],
        bold=True,
    )
    ban = add_box(slide, x + 0.14, y + 0.76, w - 0.28, 0.30, fill=C["bg_alt"])
    set_shape_text(ban, banner, S["caption"], F["head"], C["primary"], bold=True)
    ban.text_frame.word_wrap = False
    icon(slide, ic, "accent", x + (w - hero) / 2, y + 1.10, hero)
    # 하단 고정: 룰 + 칩 (카드 바닥에 정렬)
    rule_y = y + h - len(chips) * 0.24 - 0.14
    hrule(slide, x + 0.14, rule_y, w - 0.28)
    for k, (cic, em, label) in enumerate(chips):
        chip_row(slide, x + 0.14, rule_y + 0.09 + k * 0.24, w - 0.28, cic, em, label, h=0.23)
    # 가운데 채움: 히어로 밑 ~ 룰 위 구간을 불릿 수로 균등 분배
    b_top = y + 1.10 + hero + 0.14
    b_bottom = rule_y - 0.06
    step = (b_bottom - b_top) / max(len(items), 1)
    two_col_bullets(slide, x + 0.15, b_top, w - 0.26, items, step=step, h=min(0.32, step))
