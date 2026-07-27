"""fan_in — 원인 N개가 한 점에서 만나 하나의 결론으로 수렴 (골든 s04 하단 구획, 2026-07-26).

출처: 골든 S4 하단의 한계 수렴 도해를 부품 계약 3조건으로 옮긴 것. 좌표·도형·색은 골든 실측
그대로여서 `compare_golden` 픽셀 동일로 무손실을 증명한다.

## 이 도해가 무엇을 말하나

서로 다른 원인 여럿이 **결국 한 곳에서 만난다**. 카드 그리드가 아니라 수렴선이 관계를 지므로,
같은 항목을 N칸 카드로 늘어놓는 것과 뜻이 다르다 — 카드는 "이런 것들이 있다"이고 이 도해는
"이것들이 전부 한 곳으로 간다"이다. 목록으로 바꾸면 그 주장이 사라진다.

## 화살촉을 갈래마다 붙이지 않는 이유

N갈래 끝에 전부 화살촉을 달면 한 점에서 뭉개져 형태가 읽히지 않는다. 갈래는 화살촉 없는
가는 선이고, 방향은 수렴점 뒤의 화살표 **하나**가 진다.

## 종착을 점선으로 둘 수 있다

`soft=True`면 종착 도형의 테두리가 점선이 된다 — 형태 언어에서 점선은 확률·불확실이다.
"원인이 모여도 결론이 서지 않는다"를 말할 때 쓴다(실선 종착과 나란히 놓으면 "같은 판정인데
하나는 서고 하나는 못 선다"가 형태로 읽힌다).
"""

from pptx.enum.text import MSO_ANCHOR

from ..kit import add_box, add_text, mix
from . import elements as E
from . import Box, merged, require

META = {
    "name": "fan_in",
    "어휘": "G02",  # archetype-catalog 등재 어휘(cause_fanin)와 대응
    "물성": "원인·항목 N개가 한 점에서 만나 하나로 수렴한다(관계가 선으로 드러난다)",
    "arity": "파생 — 항목 수가 세로 자리를 정한다. 자리를 넘으면 시끄럽게 죽는다",
    # n 범위는 **실측**이다(주력 밴드 2.16" 실측) — 지어낸 값이 아니다. 다시 재려면
    # `verify_selection.py --measure`. 실제 자리에서의 정확한 한계는 draw가 다시 본다.
    # 셀 수 있는 선택 조건 — 컴포저는 재료를 세어 이것과 **비교만** 한다(산문 해석 없음).
    "accepts": {'sets': 1, 'flow': 'N:1', 'order': False, 'extra': True, 'n': (2, 6), 'ends': False},
    "keys": ["items", "target"],
    "optional": ["mark", "note", "soft"],
}

LAYOUT = {
    # box 원점 기준 offset(inch). 골든 s04 실측값.
    "item_w": 3.15,
    "item_h": 0.50,
    "item_pitch": 0.54,
    "label_h": 0.24,  # 항목 안 라벨 줄
    "sub_h": 0.20,  # 항목 안 설명 줄
    "item_inset": 0.22,
    "conv_dx": 4.00,  # N갈래가 한 점으로 모이는 지점
    "target_dx": 4.38,
    "target_w": 1.35,
    "target_h": 0.95,
    "mark_dx": 5.91,
    "mark_h": 0.22,
    "note_gap": 0.08,  # 종착 bottom → 하단 주석 top
    "note_h": 0.50,
}


def measure(data, kit, layout=None):  # noqa: ARG001 — 계약 시그니처 고정
    """필요 폭·높이(inch). 항목 수에서 높이가 나온다."""
    L = merged(LAYOUT, layout)
    n = len(data["items"])
    stack = (n - 1) * L["item_pitch"] + L["item_h"]
    tail = stack / 2 + L["target_h"] / 2 + L["note_gap"] + L["note_h"]
    return L["mark_dx"] + 2.04, max(stack, tail)


def _item(slide, kit, x, y, w, L, label, sub):
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    add_box(slide, x, y, w, L["item_h"], fill=C["bg_alt"], shape="round")
    inset = L["item_inset"]
    add_text(
        slide,
        x + inset,
        y + 0.05,
        w - inset - 0.12,
        L["label_h"],
        label,
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        slide,
        x + inset,
        y + 0.05 + L["label_h"] - 0.02,
        w - inset - 0.12,
        L["sub_h"],
        sub,
        S["foot"],
        F["body"],
        C["muted"],
        anchor=MSO_ANCHOR.MIDDLE,
    )


def draw(slide, box, data, kit, layout=None):
    require(data, META["keys"], META["name"])
    box = Box(*box)  # 4튜플도 받는다 — 목업 러너·장이 둘 다 쓴다
    L = merged(LAYOUT, layout)
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    items = data["items"]

    need_h = (len(items) - 1) * L["item_pitch"] + L["item_h"]
    if need_h > box.h + 1e-6:
        raise ValueError(
            f'fan_in: 항목 {len(items)}개가 {need_h:.2f}" 필요한데 자리는 {box.h:.2f}"다 — '
            "칩을 줄이면 글자가 넘치므로 항목을 줄이거나 자리를 넓힌다"
        )

    x0, y0 = box.x, box.y
    # ── 항목 N개 (개수가 자리를 정한다 — 좌표 리터럴 없음) ──
    cys = []
    for i, it in enumerate(items):
        y = y0 + i * L["item_pitch"]
        _item(slide, kit, x0, y, L["item_w"], L, it["label"], it["sub"])
        cys.append(y + L["item_h"] / 2)

    # ── 한 점으로 수렴 → 종착 ──
    fan_cy = sum(cys) / len(cys)
    conv_x, tgt_x = x0 + L["conv_dx"], x0 + L["target_dx"]
    for cy in cys:
        E.hairline(slide, kit, x0 + L["item_w"], cy, conv_x, fan_cy)
    E.arrow(slide, kit, conv_x, fan_cy, tgt_x, fan_cy)
    tgt = add_box(
        slide,
        tgt_x,
        fan_cy - L["target_h"] / 2,
        L["target_w"],
        L["target_h"],
        fill=C["bg"],
        line=C["muted"] if data.get("soft") else C["primary"],
        line_w=1.25,
        shape="diamond",
    )
    if data.get("soft"):
        E.dashed_edge(tgt)
    E.center_text(
        tgt, kit, data["target"], S["caption"], C["muted"] if data.get("soft") else C["primary"]
    )

    if data.get("mark"):
        mark_x = x0 + L["mark_dx"]
        add_text(
            slide,
            mark_x,
            fan_cy - L["mark_h"] / 2,
            box.w - L["mark_dx"] - 0.2,
            L["mark_h"],
            data["mark"],
            S["caption"],
            F["head"],
            C["primary"],
            bold=True,
        )
    if data.get("note"):
        add_text(
            slide,
            tgt_x - 0.1,
            fan_cy + L["target_h"] / 2 + L["note_gap"],
            box.w - L["target_dx"] - 0.1,
            L["note_h"],
            data["note"],
            S["caption"],
            F["body"],
            mix(C["muted"], C["primary"], 0.55),
            line_spacing=1.2,
        )


SAMPLES = [
    (
        "최소 2",
        {
            "items": [
                {"label": "근거 추적 불가", "sub": "어느 문단으로 답했는지 못 밝힌다"},
                {"label": "재현성 없음", "sub": "같은 질문에 매번 다른 판정"},
            ],
            "target": "근거 충분?",
            "mark": "✕ 세울 수 없다",
            "soft": True,
        },
    ),
    (
        "골든 4",
        {
            "items": [
                {"label": "근거 추적 불가", "sub": "어느 문단으로 답했는지 못 밝힌다"},
                {"label": "재현성 없음", "sub": "같은 질문에 매번 다른 판정"},
                {"label": "판단 출처 창작", "sub": "없는 조항·문단을 지어낸다"},
                {"label": "환각 위험", "sub": "그럴듯한 허위를 확정으로 낸다"},
            ],
            "target": "근거 충분?",
            "mark": "✕ 세울 수 없다",
            "note": "판정이 서지 않으니 '부족'을 인지할 수단이 없다 — 늘 위쪽 경로로 샌다.",
            "soft": True,
        },
    ),
]
