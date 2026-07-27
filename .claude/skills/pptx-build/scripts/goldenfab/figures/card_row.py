"""card_row — 순서 없는 독립 항목 N개를 나란히 (골든 s08·s09·s11·s16, 2026-07-27).

출처: 골든 4장이 공유하는 카드 띠를 부품 계약으로 옮긴 것. 카드 한 장을 그리는 정본은
`dense.hero_card`이고 이 부품은 **그 카드를 몇 장 어떻게 벌려 놓나**를 맡는다.
좌표는 골든 실측 그대로여서 `compare_golden` 픽셀 동일로 무손실을 증명한다.

## 이 도해가 무엇을 말하나

관계도 순서도 없이 **나란히 놓이는** 항목 N개다. 실전 보고서에서 가장 흔한 구조이고
(게이트 목록·사고 목록·한계 목록·특장점 목록), 그래서 창고에 이게 없으면 덱이 자주 막힌다
(2026-07-27 실전 재료 시험: 막힌 4건 중 3건이 이 물성이었다).

**관계가 있으면 이 부품이 아니다.** 항목들이 한 곳으로 모이면 `fan_in`, 하나에서 갈라지면
`n_branch`·`hub_spoke`, 순서가 메시지면 `numbered_steps`다. 관계를 카드로 늘어놓으면 그
관계가 사라진다 — 그게 "관계를 목록으로 파괴한다"고 여러 번 반려된 실패다.

## 왜 카드가 이렇게 두꺼운가

카드 한 장이 배지·제목·태그·배너·아이콘·불릿·칩 일곱 겹을 갖는다. 한 겹짜리 상자를 N개
늘어놓으면 화면이 커질수록 초라해진다("그냥 도형 나열 / 고등학생 ppt" 반려). 밀도는 크기가
아니라 **겹의 수**에서 온다.

## 계층에 대한 정직한 메모

2층 부품이 3층 헬퍼(`dense.hero_card`)를 부른다 — 계층 역전이다. 카드는 원소 하나가 아니라
복합 컴포넌트이고 그 정본이 `dense`에 있기 때문이다. 카드를 두 벌로 만드는 것(골든이 두 벌이라
갈라졌던 그 병)보다 이 역전이 낫다고 판단했다. 카드 컴포넌트를 1층으로 내리는 것은 별 작업.
"""

from .. import dense as D
from . import Box, merged, require

META = {
    "name": "card_row",
    "어휘": "미등재",  # 결정표는 이 물성의 정본을 `dense.hero_card`로 적어 뒀다
    "물성": "순서 없는 독립 항목 N개를 나란히(관계 없음 · 각 항목이 여러 겹을 갖는다)",
    # n 범위는 **실측**이다(주력 밴드 2.16" 실측) — 지어낸 값이 아니다. 다시 재려면
    # `verify_selection.py --measure`. 실제 자리에서의 정확한 한계는 draw가 다시 본다.
    # 셀 수 있는 선택 조건 — 컴포저는 재료를 세어 이것과 **비교만** 한다(산문 해석 없음).
    "accepts": {
        "sets": 1,
        "flow": "없음",
        "order": False,
        "extra": True,
        "ends": False,
        "n": (2, 5),
    },
    "arity": "파생 — 카드 폭이 장수에서 나온다. 최소 폭 미달이면 시끄럽게 죽는다",
    "keys": ["cards"],
    "optional": [],
}

LAYOUT = {
    # box 원점 기준. 골든 s08·s09·s11 실측(카드 띠 y 3.62 · h 3.32 · 4장 · 간격 0.18).
    "gap": 0.18,
    "row_gap": 0.14,  # (격자) 행 사이 — 골든 s12 실측
    # 골든 실측 최솟값 — s12가 1.93" 카드를 쓴다(s08·s09·s11은 2.90"). 지어낸 값이 아니다.
    "col_min": 1.92,
    "row_min": 3.00,  # 카드 한 장의 최소 높이 — 골든 실측 3.02(s12)·3.32(s08 계열)
    "hero": 0.44,  # 히어로 아이콘 지름 — 글자에 묶인 치수(자리에 비례시키지 않는다)
    # 열 수. None이면 한 줄(골든 s08·s09·s11·s16). 숫자면 격자(골든 s12는 2열).
    # 한 줄에 다 못 넣을 만큼 항목이 많을 때 **줄을 접는다** — 폭을 줄이면 카드가 깨진다.
    "cols": None,
}


def measure(data, kit, layout=None):  # noqa: ARG001 — 계약 시그니처 고정
    """필요 폭·높이(inch). 폭은 장수 × 최소 칼럼, 높이는 카드 권장 최소."""
    L = merged(LAYOUT, layout)
    n = len(data["cards"])
    return n * L["col_min"] + (n - 1) * L["gap"], 3.20


def draw(slide, box, data, kit, layout=None):
    """data = {"cards": [{"num","title","tag","ic","banner","items","chips"}, ...]}"""
    from .. import grid as G

    require(data, META["keys"], META["name"])
    box = Box(*box)
    L = merged(LAYOUT, layout)
    cards = data["cards"]
    cols = L["cols"] or len(cards)
    rows = -(-len(cards) // cols)  # 올림
    w = G.track(cols, box.x, box.x + box.w, L["gap"], L["col_min"], what="카드")
    h = (box.h - (rows - 1) * L["row_gap"]) / rows
    if h < L["row_min"]:
        raise ValueError(
            f'card_row: {rows}줄로 접으면 카드 높이가 {h:.2f}"다 — 최소 {L["row_min"]}" 필요. '
            "일곱 겹(배지·제목·태그·배너·아이콘·불릿·칩)이 안 들어간다. 항목을 줄이거나 자리를 넓힌다."
        )
    for i, c in enumerate(cards):
        D.hero_card(
            slide,
            box.x + (i % cols) * (w + L["gap"]),
            box.y + (i // cols) * (h + L["row_gap"]),
            w,
            h,
            c["num"],
            c["title"],
            c["tag"],
            c["ic"],
            c["banner"],
            c["items"],
            c["chips"],
            hero=L["hero"],
        )
    return len(cards)


SAMPLES = [
    (
        "3장",
        {
            "cards": [
                {
                    "num": str(i),
                    "title": t,
                    "tag": g,
                    "ic": "shield-check",
                    "banner": b,
                    "items": [("검출", "결함을 실제로 잡는다"), ("오탐", "정상은 통과시킨다")],
                    "chips": [("circle-check-big", "자기 검증", "있음")],
                }
                for i, (t, g, b) in enumerate(
                    (
                        ("회귀 게이트", "compare_golden", "17/17장 · 797/797도형"),
                        ("계약 게이트", "check_contract", "검사 6종"),
                        ("밀도 오딧", "audit_deck", "generic 8항목"),
                    ),
                    1,
                )
            ]
        },
    ),
]
