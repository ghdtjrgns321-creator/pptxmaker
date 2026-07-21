"""S4 문제정의 — dense 개편(2026-07-20 재작업 v2). R0/R1: 골든 variant_k 재사용, 재드로 없음.

초판(폐기)은 골든 게이트 인과도를 상단 띠로 재드로하고 한계 fan-in 수렴을 4개 hero_card로
파괴했다(§E-2 관계→목록, R2 억지 4카드). v1은 반대로 재사용했으나 검은 바 자리를 메우려
도해를 세로로 늘려(ty 1.26×) 밀도가 떨어졌다(사용자: "비어 보인다").

v2 교정:
- 헤더 → dense.compact_header (골든 28pt 헤더 폐지)
- 검은 결론 바 제거 → 원리는 우하단 해결 패널로(§8 하단 전폭 바 금지, §5 패널 어휘)
- **구획1(게이트 인과도)** = 골든 variant_k 그대로, 세로 늘림 없이 위로 이동만(순수 평행이동).
  도형·색·좌표비율 100% 골든 → 밀도 보존.
- **구획2(한계 fan-in)** = 같은 수렴 도해(N개 → 못 세운 ◇로 수렴)지만 칩을 라벨+설명 2줄로
  채워 하단 실질 밀도 확보. 카드 그리드 아님 — 수렴선이 관계를 진다(§E-2). 항목 크기는
  바뀌어도 도해 문법(fan-in→ghost◇)은 골든 그대로(P5.2).

콘텐츠는 _variant_k.DEFAULT + 팩트시트 한계 설명(§7 축약). 4카드 강제 없음.
"""

from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from . import _variant_k as VK
from . import dense as D
from . import grid as G
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]

KICKER = VK.DEFAULT["kicker"]
HEADLINE = VK.DEFAULT["headline"]
SOURCE = VK.DEFAULT["source"]

# content_contract 계약 키 — registry가 이 dense를 golden.problem_grid로 부를 때 필수 키 출처.
DEFAULT = VK.DEFAULT

# ── 구획1: 골든 게이트 인과도 — 순수 평행이동(늘리지 않음). ty(y)=y+DY ──
# 골든 band1_head(1.75)를 압축 헤더 밑 0.90으로 올린다. 세로 스케일 1.0 → 골든 밀도 그대로.
_DY = 0.90 - VK.B1_HEAD_Y  # -0.85


def ty(y):
    return y + _DY


# ── 구획2: 한계 fan-in (하단, 라벨+설명으로 채운 수렴 도해) ──
B2_HEAD_Y = 4.02
LIMITS = [
    ("근거 추적 불가", "어느 문단으로 답했는지 못 밝힌다"),
    ("재현성 없음", "같은 질문에 매번 다른 판정"),
    ("판단 출처 창작", "없는 조항·문단을 지어낸다"),
    ("환각 위험", "그럴듯한 허위를 확정으로 낸다"),
]
ITEM_X, ITEM_W = G.MARGIN_L, 3.15
ITEM_TOP, ITEM_H, ITEM_PITCH = 4.66, 0.50, 0.54
CONV_PT_X = 4.60
GHOST_X, GHOST_W, GHOST_H = 4.98, 1.35, 0.95
MARK_X = GHOST_X + GHOST_W + 0.18  # 6.51
PANEL_X = 8.55


def _limit_item(slide, y, label, sub):
    add_box(slide, ITEM_X, y, ITEM_W, ITEM_H, fill=C["bg_alt"], shape="round")
    add_text(
        slide,
        ITEM_X + 0.22,
        y + 0.05,
        ITEM_W - 0.34,
        0.24,
        label,
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        slide,
        ITEM_X + 0.22,
        y + 0.27,
        ITEM_W - 0.34,
        0.20,
        sub,
        S["foot"],
        F["body"],
        C["muted"],
        anchor=MSO_ANCHOR.MIDDLE,
    )


def _solution_panel(slide):
    """검은 바로 있던 해결 원리를, 빈 우하단을 채우는 단일 액센트 패널로(§5 패널 어휘·§8 바 금지).

    fan-in이 "왜 못 세우나"를 말한 뒤, 이 패널이 "그래서 이 시스템은 세운다"로 받는다. 4카드 아님."""
    px = PANEL_X
    pw = G.RIGHT_EDGE - px
    py, ph = 4.30, 2.36
    add_box(slide, px, py, pw, ph, fill=C["bg_alt"])
    add_box(slide, px, py, 0.05, ph, fill=C["accent"])  # 액센트 좌변 바(§5 컨테인먼트)
    D.icon(slide, "circle-check-big", "accent", px + 0.24, py + 0.24, 0.30)
    add_text(
        slide,
        px + 0.64,
        py + 0.22,
        pw - 0.82,
        0.32,
        "그래서 — 편향을 구조로 건다",
        S["head"],
        F["head"],
        C["accent"],
        bold=True,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        slide,
        px + 0.26,
        py + 0.72,
        pw - 0.5,
        0.9,
        [
            [
                ("확실할 때만 확정", {"bold": True, "color": C["primary"]}),
                ("하고, 애매하면 근거를 보여주며 ", {}),
                ("유보", {"bold": True, "color": C["primary"]}),
                ("한다. 판정을 프롬프트가 아니라 시스템 ", {}),
                ("구조", {"bold": True, "color": C["primary"]}),
                ("로 세운다.", {}),
            ]
        ],
        S["caption"],
        F["body"],
        D.ink_body,
        line_spacing=1.4,
    )
    # 조건 요약 칩 2 — 확정/유보의 트리거(밀도 보강)
    cy = py + ph - 0.52
    cw = (pw - 0.5) / 2 - 0.08
    for i, (em, label) in enumerate((("확정", "근거 충분"), ("유보", "근거 부족"))):
        cx = px + 0.26 + i * (cw + 0.16)
        chip = add_box(slide, cx, cy, cw, 0.36, fill=C["bg"], line=D.line_mid, line_w=0.75)
        tf = chip.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        tf.margin_left = Inches(0.14)
        p = tf.paragraphs[0]
        p._p.get_or_add_pPr().set("eaLnBrk", "0")

        for txt, color, bold in ((em + "  ", C["accent"], True), (label, C["muted"], False)):
            r = p.add_run()
            r.text = txt
            r.font.name, r.font.bold, r.font.color.rgb = F["head"], bold, color
            r.font.size = Pt(S["foot"])
            r.font.language_id = MSO_LANGUAGE_ID.KOREAN


def build(prs, c=None):
    c = {**DEFAULT, **(c or {})}
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])

    # ── 헤더: dense 압축 ──
    D.compact_header(slide, KICKER, HEADLINE)

    # ── 구획 1: 게이트 인과도 (골든 variant_k 그대로 · 평행이동만) ──
    VK._band_head(slide, ty(VK.B1_HEAD_Y), c["band1_head"])
    q = add_box(
        slide,
        G.MARGIN_L,
        ty(VK.DIA_CY) - VK.Q_H / 2,
        VK.Q_W,
        VK.Q_H,
        fill=C["bg_alt"],
        line=C["muted"],
        line_w=0.75,
        shape="round",
    )
    VK._center(q, c["question"], S["caption"], C["primary"])
    VK._arrow(slide, G.MARGIN_L + VK.Q_W, ty(VK.DIA_CY), VK.DIA_X, ty(VK.DIA_CY))
    dia = add_box(
        slide,
        VK.DIA_X,
        ty(VK.DIA_CY) - VK.DIA_H / 2,
        VK.DIA_W,
        VK.DIA_H,
        fill=C["bg"],
        line=C["primary"],
        line_w=1.25,
        shape="diamond",
    )
    VK._center(dia, c["gate"], S["caption"], C["primary"])
    VK._arrow(slide, VK.DIA_RIGHT, ty(VK.DIA_CY), VK.L2_X, ty(VK.ROW_A_CY), elbow=True)
    VK._arrow(slide, VK.DIA_RIGHT, ty(VK.DIA_CY), VK.L2_X, ty(VK.LACK_CY), elbow=True)
    VK._branch_label(slide, VK.DIA_RIGHT + 0.06, ty(VK.ROW_A_CY) - 0.30, c["branch_ok"], C["muted"])
    VK._branch_label(
        slide, VK.DIA_RIGHT + 0.06, ty(VK.LACK_CY) + 0.06, c["branch_lack"], C["primary"]
    )
    ok = add_box(
        slide,
        VK.L2_X,
        ty(VK.ROW_A_Y),
        VK.L2_W,
        VK.ROW_A_H,
        fill=C["bg_alt"],
        line=C["muted"],
        line_w=0.75,
        shape="round",
    )
    VK._center(ok, c["ok_chip"], S["body"], C["primary"])
    add_text(
        slide,
        VK.L2_RIGHT + 0.4,
        ty(VK.ROW_A_Y) - 0.02,
        G.RIGHT_EDGE - VK.L2_RIGHT - 0.4,
        VK.ROW_A_H,
        c["ok_note"],
        S["caption"],
        F["body"],
        VK.INK_NOTE,
        line_spacing=1.3,
    )
    lack = add_box(
        slide,
        VK.L2_X,
        ty(VK.LACK_CY) - VK.ROW_A_H / 2,
        VK.L2_W,
        VK.ROW_A_H,
        fill=C["bg"],
        line=C["primary"],
        line_w=1.25,
        shape="round",
    )
    VK._center(lack, c["lack_chip"], S["body"], C["primary"])
    for cy in (VK.ROW_B_CY, VK.ROW_C_CY):
        VK._arrow(slide, VK.L2_RIGHT, ty(VK.LACK_CY), VK.ROW_X, ty(cy), elbow=True)

    row_style = [
        (VK.ROW_B_Y, C["accent"], C["bg"], True),
        (VK.ROW_C_Y, C["bg_alt"], C["primary"], False),
    ]
    for (ry, err_fill, err_ink, fatal), row in zip(row_style, c["rows"]):
        act = add_box(
            slide,
            VK.ROW_X,
            ty(ry),
            VK.ACT_W,
            VK.ROW_BC_H,
            fill=C["bg_alt"],
            line=C["muted"],
            line_w=0.75,
            shape="round",
        )
        VK._center(act, row["act"], S["caption"], C["primary"])
        rcy = ty(ry) + VK.ROW_BC_H / 2
        VK._arrow(slide, VK.ROW_X + VK.ACT_W, rcy, VK.ERR_X, rcy)
        err = add_box(slide, VK.ERR_X, ty(ry), VK.ERR_W, VK.ROW_BC_H, fill=err_fill, shape="round")
        VK._center(err, row["err"], S["caption"], err_ink)
        cap_y = ty(ry) + VK.ROW_BC_H + VK.CAP_DY
        if fatal:
            VK._arrow(slide, VK.ERR_X + VK.ERR_W, rcy, VK.TERM_X, rcy)
            term = add_box(
                slide, VK.TERM_X, ty(ry), VK.TERM_W, VK.ROW_BC_H, fill=C["accent"], shape="rect"
            )
            VK._center(term, row["term"], S["caption"], C["bg"])
            add_text(
                slide,
                VK.TERM_X,
                cap_y,
                VK.TERM_W,
                VK.CAP_H,
                row["mark"],
                S["caption"],
                F["head"],
                C["primary"],
                bold=True,
                align=PP_ALIGN.CENTER,
            )
            add_text(
                slide,
                VK.ERR_X,
                cap_y,
                VK.TERM_X - VK.ERR_X - 0.1,
                VK.CAP_H,
                row["note"],
                S["caption"],
                F["body"],
                VK.INK_NOTE,
            )

    # 복귀 호 — 1종은 질문으로 되돌아온다
    back_x = VK.ERR_X + VK.ERR_W / 2
    q_cx = G.MARGIN_L + VK.Q_W / 2
    VK._line(slide, back_x, ty(VK.ROW_C_Y) + VK.ROW_BC_H, back_x, ty(VK.BACK_Y))
    VK._line(slide, back_x, ty(VK.BACK_Y), q_cx, ty(VK.BACK_Y))
    VK._arrow(slide, q_cx, ty(VK.BACK_Y), q_cx, ty(VK.DIA_CY) + VK.Q_H / 2)
    add_text(
        slide,
        VK.BACK_LABEL_X,
        ty(VK.BACK_LABEL_Y),
        VK.BACK_LABEL_W,
        VK.CAP_H,
        c["back_label"],
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
    )

    # ── 구획 2: 한계 4가지 fan-in 수렴 (라벨+설명으로 채움 · 카드 아님) ──
    VK._band_head(slide, B2_HEAD_Y, c["band2_head"])
    item_cys = []
    for i, (label, sub) in enumerate(LIMITS):
        y = ITEM_TOP + i * ITEM_PITCH
        _limit_item(slide, y, label, sub)
        item_cys.append(y + ITEM_H / 2)
    fan_cy = sum(item_cys) / len(item_cys)
    for cy in item_cys:
        VK._line(slide, ITEM_X + ITEM_W, cy, CONV_PT_X, fan_cy)
    VK._arrow(slide, CONV_PT_X, fan_cy, GHOST_X, fan_cy)
    ghost = add_box(
        slide,
        GHOST_X,
        fan_cy - GHOST_H / 2,
        GHOST_W,
        GHOST_H,
        fill=C["bg"],
        line=C["muted"],
        line_w=1.25,
        shape="diamond",
    )
    VK._dash(ghost)
    VK._center(ghost, c["gate"], S["caption"], C["muted"])
    add_text(
        slide,
        MARK_X,
        fan_cy - VK.CAP_H / 2,
        PANEL_X - MARK_X - 0.2,
        VK.CAP_H,
        c["converge_mark"],
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_text(
        slide,
        GHOST_X - 0.1,
        fan_cy + GHOST_H / 2 + 0.08,
        PANEL_X - GHOST_X - 0.1,
        0.5,
        "판정이 서지 않으니 '부족'을 인지할 수단이 없다 — 늘 위쪽 경로로 샌다.",
        S["caption"],
        F["body"],
        VK.INK_NOTE,
        line_spacing=1.2,
    )
    _solution_panel(slide)

    D.source_line(slide, SOURCE)
    return slide
