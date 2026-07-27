"""numbered_steps — 번호 붙은 절차 N단 (골든 s11 상단, 2026-07-26).

출처: `s11_dense._inject_and_branch`의 주입 단계 흐름을 부품 계약 3조건으로 옮긴 것.
좌표·도형·색은 골든 실측 그대로여서 `compare_golden` 픽셀 동일로 무손실을 증명한다.

## 이 도해가 무엇을 말하나

**순서가 메시지**인 절차다. 단계마다 번호 배지 + 제목 + 설명 한 덩이가 가로로 놓이고 사이를
화살표가 잇는다. 박스로 감싸지 않는 이유: 상자를 두르면 "단계"가 아니라 "항목"으로 읽히고,
순서가 아니라 목록이 된다. 순서가 메시지가 아니면 이 도해가 아니라 카드 나열을 쓴다.

## 단계 수는 자유다

칼럼 폭을 단계 수에서 파생한다(`grid.track`). 최소 폭보다 좁아지면 시끄럽게 죽는다 —
전엔 폭이 리터럴(`/3`)이라 단계가 하나 늘면 마지막이 조용히 칸을 넘었다.
"""

from pptx.enum.text import MSO_ANCHOR

from ..kit import add_box, add_text, set_shape_text
from . import elements as E
from . import Box, merged, require

META = {
    "name": "numbered_steps",
    "어휘": "G10",  # archetype-catalog 등재 어휘(numbered_steps)와 대응
    "물성": "절차가 몇 단계이고 순서 자체가 메시지다(박스 없는 번호 흐름)",
    "arity": "파생 — 칼럼 폭이 단계 수에서 나온다. 최소 폭 미달이면 시끄럽게 죽는다",
    # n 범위는 **실측**이다(주력 밴드 2.16" 실측) — 지어낸 값이 아니다. 다시 재려면
    # `verify_selection.py --measure`. 실제 자리에서의 정확한 한계는 draw가 다시 본다.
    # 셀 수 있는 선택 조건 — 컴포저는 재료를 세어 이것과 **비교만** 한다(산문 해석 없음).
    "accepts": {'sets': 1, 'flow': '순차', 'order': True, 'extra': False, 'n': (2, 5), 'ends': False},
    "keys": ["steps"],
    "optional": [],
}

LAYOUT = {
    # box 원점 기준 offset(inch). 골든 s11 실측값.
    "badge_d": 0.26,  # 번호 배지 지름 — 글자에 묶인 치수(자리에 비례시키지 않는다)
    "gap": 0.30,  # 칼럼 사이(화살표가 지나는 자리)
    "col_min": 1.30,  # 이보다 좁으면 제목이 두 줄로 깨진다
    "head_dx": 0.34,  # 배지 우변 → 제목
    "head_dy": -0.02,
    "head_h": 0.30,
    "body_dy": 0.38,  # 배지 top → 설명 top
    "body_h": 0.62,
    "arrow_dy": 0.13,  # 배지 세로중심
    "arrow_inset": 0.04,  # 칼럼 끝에서 살짝 떨어뜨린다(붙으면 글자에 닿는다)
}


def measure(data, kit, layout=None):  # noqa: ARG001 — 계약 시그니처 고정
    """필요 폭·높이(inch). 폭은 단계 수 × 최소 칼럼."""
    L = merged(LAYOUT, layout)
    n = len(data["steps"])
    return n * L["col_min"] + (n - 1) * L["gap"], L["body_dy"] + L["body_h"]


def draw(slide, box, data, kit, layout=None):
    """data = {"steps": [{"head": str, "body": str}, ...]}"""
    from .. import grid as G

    require(data, META["keys"], META["name"])
    box = Box(*box)  # 4튜플도 받는다 — 목업 러너·장이 둘 다 쓴다
    L = merged(LAYOUT, layout)
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    steps, gap, d = data["steps"], L["gap"], L["badge_d"]

    cw = G.track(len(steps), box.x, box.x + box.w, gap, L["col_min"], what="단계")
    for i, st in enumerate(steps):
        cx = box.x + i * (cw + gap)
        badge = add_box(slide, cx, box.y, d, d, fill=C["primary"], shape="oval")
        set_shape_text(badge, str(i + 1), S["foot"], F["head"], C["bg"], bold=True)
        add_text(
            slide,
            cx + L["head_dx"],
            box.y + L["head_dy"],
            cw - L["head_dx"],
            L["head_h"],
            st["head"],
            S["caption"],
            F["head"],
            C["primary"],
            bold=True,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            slide,
            cx,
            box.y + L["body_dy"],
            cw,
            L["body_h"],
            st["body"],
            S["foot"],
            F["body"],
            C["muted"],
            line_spacing=1.18,
        )
        if i < len(steps) - 1:
            ay = box.y + L["arrow_dy"]
            E.arrow(
                slide, kit, cx + cw + L["arrow_inset"], ay, cx + cw + gap - L["arrow_inset"], ay
            )


SAMPLES = [
    (
        "3단",
        {
            "steps": [
                {"head": "주제 지목", "body": "35개 주제에서 최대 3개 지목"},
                {"head": "직속 개념 매칭", "body": "발동 개념과 겹치면 걸린다"},
                {"head": "걸린 것 전부 주입", "body": "1개만 고르지 않는다"},
            ]
        },
    ),
]
