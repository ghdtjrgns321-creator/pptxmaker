"""layer_stack — 계층 3층. 위는 하나, 가운데는 동등한 N, 아래는 다시 하나로 모인다.

실물: Rehabilitative Care Alliance 「Hub and Spoke Model for Specialized Rehab」(2024-04) p3.
part-design ② 선택지 시트에서 사용자가 고른 형태(②)다.

## 무엇과 갈리나 (이 부품을 왜 따로 두나)

- **방사(`hub_spoke`)와** — 저쪽은 중심 하나에 N이 물린 **1:N 소속**이고 층이 없다.
  이쪽은 **층이 메시지**다: 위(권위·정의) · 가운데(동등한 실행 주체) · 아래(그것이 닿는 현장).
- **카탈로그 어휘 G07 계층 트리와** — 저쪽은 `1 → N → M`으로 **계속 갈라진다**(root→concepts→paras,
  `s09_dense.py:_tree`).
  이쪽은 `1 → N → 1`로 **아래에서 다시 모인다**. 갈라지는 트리는 분류를 말하고, 모이는 3층은
  "여러 주체가 결국 한 현장에 닿는다"를 말한다. 같은 "3층"이지만 주장이 반대라 한 부품에
  넣지 않았다 — 넣으면 둘 다 어중간해진다.

## 잉크

면을 채우지 않는다(§2 · `hub_spoke` 파일의 MGI 실측과 같은 규율). 실물은 파란 채움이지만 검정 채움은
2026-07-26 눈검증에서 반려됐다 — 층의 구분은 **채움이 아니라 도형 종류와 자리**가 만든다:
위는 전폭 띠, 가운데는 나란한 상자, 아래는 타원. 세 층이 같은 회색 테두리여도 층으로 읽힌다.

## 좌측 레일

실물에는 전체를 감싸는 세로 라벨(「Regional Rehab Networks」)이 있다 — 3층 전체가 **어느 범위
안의 이야기인가**를 말한다. `rail`이 없으면 그리지 않는다(글 무소유 · 없는 걸 지어 채우지 않는다).
"""

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from ..kit import add_box, add_text
from . import Box, merged, require
from .elements import chip_w, link

META = {
    "name": "layer_stack",
    "어휘": "미등재",  # 카탈로그 G07(계층 트리)과는 물성이 다르다 — 아래 설명 참조
    "물성": "층이 몇이고 각 층이 무엇인가(1→N→1 · 위는 정의 · 아래에서 다시 모인다)",
    "arity": "파생",
    # 셀 수 있는 선택 조건 — 컴포저는 재료를 세어 이것과 **비교만** 한다(산문 해석 없음).
    "accepts": {'sets': 1, 'flow': '순차', 'order': True, 'extra': False, 'n': (2, 5), 'ends': True},
    "keys": ["band", "tiers", "base"],
    "optional": ["band_note", "base_note", "rail", "note", "soft"],
}

LAYOUT = {
    "band_h": 0.46,  # 상단 띠 높이 — 글자에 묶인 치수(자리에 비례시키지 않는다)
    "tier_h": 0.62,  # 가운데 상자 높이. note가 있으면 파생해서 늘어난다
    "base_h": 0.74,  # 하단 타원 높이 — 타원은 글자가 가운데만 쓰므로 상자보다 높아야 한다
    "gap": 0.22,  # 가운데 상자 사이 간격
    "min_w": 1.15,  # 상자 최소 폭 — 이하면 글자가 세로로 쪼개진다(track이 죽인다)
    "max_w": 2.45,  # 상자 최대 폭 — **상한이 없으면 층이 아니라 표로 읽힌다**(2026-07-26 눈검증:
    #                 12.1" 자리에서 상자가 4"씩 되어 3층이 세 줄짜리 표가 됐다). 남는 가로는
    #                 3층이 설명을 두는 자리다 — 본체를 가운데로 모으고 measure()로 알린다.
    "v_gap_min": 0.30,  # 층 사이 최소 거리 — 이하면 3층이 한 덩어리로 뭉친다
    "v_gap_max": 0.62,  # 상한 — 자리가 커도 층을 더 벌리지 않는다. 벌리면 층이 따로 논다
    "base_w": 0.62,  # 하단 타원 폭 = 가운데 층 폭 × 이 비율(실물도 가운데보다 좁다)
    "base_ratio": 5.0,  # 타원 가로:세로 상한 — 넘으면 타원이 아니라 눌린 렌즈로 보인다
    "rail_w": 0.22,  # 좌측 레일 띠 폭
    "rail_gap": 0.14,  # 레일 ↔ 본체
    "lateral": True,  # 가운데 상자끼리 좌우로 잇는가(실물은 ↔ 로 이어 동등을 말한다)
    "max_n": 5,  # 6개 넘으면 층이 아니라 나열로 읽힌다
}


def _vertical(tb):
    """텍스트를 세로쓰기로 — 좌측 레일 전용. bodyPr vert 속성이 유일한 경로다."""
    from pptx.oxml.ns import qn

    bodyPr = tb.text_frame._txBody.find(qn("a:bodyPr"))
    bodyPr.set("vert", "vert270")
    return tb


def _tier_h(data, kit, L):
    """가운데 상자 높이 — note 유무에서 파생. 한 겹이면 낮고 두 겹이면 높다."""
    return L["tier_h"] + (0.22 if any(t.get("note") for t in data["tiers"]) else 0)


def _label(slide, x, y, w, h, title, note, kit, *, ink=None, note_ink=None):
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    paras = [[(title, {"size": S["body"], "bold": True, "color": ink or C["primary"]})]]
    if note:
        paras.append([(note, {"size": S["caption"], "color": note_ink or C["muted"]})])
    add_text(
        slide,
        x,
        y,
        w,
        h,
        paras,
        S["body"],
        F["head"],
        C["primary"],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
        line_spacing=1.06,
    )


def measure(data, kit, layout=None):
    """(필요 폭, 필요 높이) — 3층이 자리를 손으로 재지 않게 하는 값."""
    L = merged(LAYOUT, layout)
    n = len(data["tiers"])
    w = n * L["min_w"] + (n - 1) * L["gap"]
    if data.get("rail"):
        w += L["rail_w"] + L["rail_gap"]
    h = L["band_h"] + _tier_h(data, kit, L) + L["base_h"] + L["v_gap_min"] * 2
    return w, h


def draw(slide, box, data, kit, layout=None):
    """box 안에 3층을 그린다. 반환: {라벨: (cx, cy)} — 오딧용.

    data = {"band": str, "band_note": str,
            "tiers": [{"label": str, "note": str, "soft": bool}, ...],
            "base": str, "base_note": str, "rail": str}
    """
    from ..grid import track

    require(data, META["keys"], META["name"])
    L = merged(LAYOUT, layout)
    box = Box(*box)
    C, S = kit["rgb"], kit["sizes"]
    tiers = data["tiers"]
    n = len(tiers)
    if n > L["max_n"]:
        raise ValueError(
            f"가운데 층 {n}개는 층으로 안 읽힌다(최대 {L['max_n']}) — 나열로 보인다. "
            f"`hub_spoke` 방사나 가로 나열을 쓸 것."
        )

    rail_off = (L["rail_w"] + L["rail_gap"]) if data.get("rail") else 0.0
    # 본체 폭에 **상한**이 있다 — 상자 폭 상한에서 파생. 남는 자리는 가운데로 모아 비운다.
    bw = min(box.w - rail_off, n * L["max_w"] + (n - 1) * L["gap"])
    x0 = box.x + rail_off + (box.w - rail_off - bw) / 2
    tier_h = _tier_h(data, kit, L)

    # 층 사이 거리 — 자리에서 파생하되 상한이 있다. 남으면 벌리지 말고 덩이를 세로 중앙에.
    room = box.h - L["band_h"] - tier_h - L["base_h"]
    if room < L["v_gap_min"] * 2:
        raise ValueError(
            f'3층이 이 자리에 안 들어간다 — 층 사이에 쓸 높이 {room:.2f}" '
            f'< 최소 {L["v_gap_min"] * 2:.2f}". 자리 높이를 늘릴 것.'
        )
    vg = min(room / 2, L["v_gap_max"])
    total = L["band_h"] + vg + tier_h + vg + L["base_h"]
    y0 = box.y + (box.h - total) / 2
    band_y = y0
    tier_y = band_y + L["band_h"] + vg
    base_y = tier_y + tier_h + vg
    pos = {}

    # ① 위 — 전폭 띠. 이 층이 나머지를 정의한다.
    add_box(slide, x0, band_y, bw, L["band_h"], fill=None, line=C["primary"], line_w=1.25)
    _label(slide, x0, band_y, bw, L["band_h"], data["band"], data.get("band_note"), kit)
    pos[data["band"]] = (x0 + bw / 2, band_y + L["band_h"] / 2)

    # ② 가운데 — 나란한 N. 개수에서 폭을 파생한다(좌표 리터럴 금지 · G07이 앓던 병).
    cw = track(n, x0, x0 + bw, L["gap"], L["min_w"], what="가운데 층")
    base_cx = x0 + bw / 2
    for i, t in enumerate(tiers):
        x = x0 + i * (cw + L["gap"])
        soft = bool(t.get("soft"))
        add_box(
            slide,
            x,
            tier_y,
            cw,
            tier_h,
            fill=None,
            line=C["accent"] if soft else C["muted"],
            line_w=1.25 if soft else 0.75,
        )
        _label(
            slide,
            x,
            tier_y,
            cw,
            tier_h,
            t["label"],
            t.get("note"),
            kit,
            ink=C["accent"] if soft else C["primary"],
        )
        pos[t["label"]] = (x + cw / 2, tier_y + tier_h / 2)
        # 위 띠에서 내려오는 소속선 — 촉 없음(방향이 아니라 소속이다)
        link(
            slide,
            kit,
            x + cw / 2,
            band_y + L["band_h"],
            x + cw / 2,
            tier_y,
            head=False,
            width=0.75,
        )
        # 아래로 모이는 화살 — 여기는 촉이 있다(모인다가 주장)
        link(
            slide,
            kit,
            x + cw / 2,
            tier_y + tier_h,
            base_cx + (x + cw / 2 - base_cx) * 0.35,
            base_y,
            width=0.9,
            color="accent" if soft else "muted",
            dashed=soft,
        )
        # 좌우 잇기 — 동등을 말한다. 마지막 칸 뒤에는 없다.
        if L["lateral"] and i < n - 1:
            link(
                slide,
                kit,
                x + cw,
                tier_y + tier_h / 2,
                x + cw + L["gap"],
                tier_y + tier_h / 2,
                head=False,
                width=0.75,
            )

    # ③ 아래 — 타원 하나. 여러 주체가 결국 여기에 닿는다.
    ew = max(bw * L["base_w"], chip_w(data["base"], S["body"], 0.90))
    ew = min(ew, L["base_h"] * L["base_ratio"])  # 납작한 렌즈가 되는 것을 막는다
    add_box(
        slide,
        base_cx - ew / 2,
        base_y,
        ew,
        L["base_h"],
        fill=None,
        line=C["primary"],
        line_w=1.0,
        shape="oval",
    )
    _label(
        slide,
        base_cx - ew / 2 + 0.16,
        base_y,
        ew - 0.32,
        L["base_h"],
        data["base"],
        data.get("base_note"),
        kit,
    )
    pos[data["base"]] = (base_cx, base_y + L["base_h"] / 2)

    # ④ 좌측 레일 — 전체가 어느 범위 안인가. 없으면 안 그린다.
    if data.get("rail"):
        rx = x0 - L["rail_gap"] - L["rail_w"]  # 본체가 가운데로 모였으므로 레일도 따라간다
        add_box(slide, rx, y0, L["rail_w"], total, fill=C["primary"])
        tb = add_text(
            slide,
            rx,
            y0,
            L["rail_w"],
            total,
            data["rail"],
            S["caption"],
            kit["fonts"]["head"],
            C["bg"],
            bold=True,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        _vertical(tb)
        pos[data["rail"]] = (rx + L["rail_w"] / 2, y0 + total / 2)

    return pos


_D3 = {
    "band": "판정 엔진 계층",
    "band_note": "규칙 · 그래프 · 근거를 한 곳에서 정의한다",
    "tiers": [
        {"label": "계약 식별", "note": "조문 대조"},
        {"label": "수행의무", "note": "분리 판정"},
        {"label": "거래가격", "note": "배분"},
    ],
    "base": "현장 적용",
    "base_note": "실무 담당자가 실제로 쓰는 자리",
    "rail": "전사 회계 표준",
}
_D4 = {
    "band": "판정 엔진 계층",
    "tiers": [
        {"label": "계약 식별"},
        {"label": "수행의무"},
        {"label": "거래가격"},
        {"label": "인식 시점", "soft": True},
    ],
    "base": "현장 적용",
}
_D2 = {
    "band": "단일 진실 원천",
    "band_note": "표제어 423개",
    "tiers": [
        {"label": "자동 판정", "note": "196건"},
        {"label": "사람 검토", "note": "18건"},
    ],
    "base": "질의 응답",
}

SAMPLES = [
    ("A · 3층 · 보조줄 + 좌측 레일 (RCA 실물)", _D3, None),
    ("A · 4개 · 하나는 불확실(점선)", _D4, None),
    ("A · 2개 · 레일 없음", _D2, None),
    ("B · 좌우 잇기 없이 — 동등을 말하지 않을 때", _D3, {"lateral": False}),
]
