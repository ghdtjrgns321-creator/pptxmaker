"""n_branch — 하나가 몇 갈래로 갈리나 (골든 s11 하단, 2026-07-26).

출처: `s11_dense._inject_and_branch`의 갈래 블록을 부품 계약 3조건으로 옮긴 것. 좌표·도형·색은
골든 실측 그대로여서 `compare_golden` 픽셀 동일로 무손실을 증명한다.

## 이 도해가 무엇을 말하나

**갈래 수 자체가 메시지**다. 판정이 결과를 가르는 인과(`gate_branch`)와 갈리는 지점: 저쪽은
"무엇이 갈랐나"가 주제라 갈래마다 결말을 따라가고, 이쪽은 "몇 갈래냐"가 주제라 갈래를
나란히 편다. 갈래마다 결말이 달라 따라가야 하면 이 부품이 아니다.

## 부채꼴 축은 출발 상자의 세로중심이다

갈래 블록을 출발 상자 중심에 맞춘다. 위에서부터 쌓으면 갈래가 늘어날수록 부채꼴이 아래로
기울어 "한 곳에서 갈린다"가 흐려진다.

## 갈래 수는 자유다

행 피치를 갈래 수에서 파생한다(`grid.pitch`, 리듬 상한 cap). 밴드에 안 들어가면 시끄럽게
죽는다 — 전엔 y가 리터럴 3개라 갈래가 하나 늘면 마지막이 조용히 사라졌다.
"""

from pptx.enum.text import MSO_ANCHOR

from ..kit import add_box, add_text, set_shape_text
from . import elements as E
from . import merged, require

META = {
    "name": "n_branch",
    "어휘": "G11",  # archetype-catalog 등재 어휘(n_branch)와 대응
    "물성": "하나가 몇 갈래로 갈리나(갈래 수 자체가 메시지 · 갈래는 나란히 편다)",
    "arity": "파생 — 갈래 수가 피치를 정한다. 밴드를 넘으면 시끄럽게 죽는다",
    "keys": ["source", "branches"],
    "optional": [],
}

LAYOUT = {
    # box 원점 기준 offset(inch). 골든 s11 실측값. box = 갈래가 쓸 수 있는 세로 밴드.
    "src_dy": 0.28,  # 밴드 top → 출발 상자 top
    "src_w": 1.35,
    "src_h": 0.44,
    "rows_gap": 0.50,  # 출발 상자 우변 → 갈래 행 좌변(화살표 자리)
    "row_h": 0.24,
    "pitch_cap": 0.30,  # 리듬 상한 — 자리가 남아도 이보다 벌리지 않는다
    "arrow_inset": 0.06,  # 행 좌변에서 살짝 떨어뜨린다
    "arrow_dy": 0.11,  # 행 top → 화살표 도착 y(행 세로중심 근사)
}


def measure(data, kit, layout=None):  # noqa: ARG001 — 계약 시그니처 고정
    """필요 폭·높이(inch)."""
    L = merged(LAYOUT, layout)
    n = len(data["branches"])
    return L["src_w"] + L["rows_gap"] + 2.0, (n - 1) * L["pitch_cap"] + L["row_h"]


def draw(slide, box, data, kit, layout=None):
    """data = {"source": str, "branches": [{"label": str, "desc": str}, ...]}"""
    from .. import grid as G

    require(data, META["keys"], META["name"])
    L = merged(LAYOUT, layout)
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    branches, row_h = data["branches"], L["row_h"]

    qy = box.y + L["src_dy"]
    qb = add_box(slide, box.x, qy, L["src_w"], L["src_h"], fill=C["primary"], shape="round")
    set_shape_text(qb, data["source"], S["foot"], F["head"], C["bg"], bold=True)

    rows_x = box.x + L["src_w"] + L["rows_gap"]
    n = len(branches)
    pitch = G.pitch(n, box.y, box.y + box.h, row_h, cap=L["pitch_cap"], what="분기 갈래")
    # 부채꼴 축 = 출발 상자 세로중심
    top = (qy + L["src_h"] / 2) - ((n - 1) * pitch + row_h) / 2
    for i, br in enumerate(branches):
        ry = top + i * pitch
        E.arrow(
            slide,
            kit,
            box.x + L["src_w"],
            qy + L["src_h"] / 2,
            rows_x - L["arrow_inset"],
            ry + L["arrow_dy"],
        )
        add_text(
            slide,
            rows_x,
            ry,
            box.x + box.w - rows_x,
            row_h,
            [[(br["label"], {"bold": True, "color": C["primary"]}), (" " + br["desc"], {})]],
            S["foot"],
            F["body"],
            C["muted"],
            anchor=MSO_ANCHOR.MIDDLE,
        )


SAMPLES = [
    (
        "3갈래",
        {
            "source": "주입된 질문",
            "branches": [
                {"label": "사실관계 부족", "desc": "되물음 — 부족한 사실을 묻는다"},
                {"label": "산술 필요", "desc": "계산을 확인한 뒤 답변"},
                {"label": "그 외", "desc": "확정 또는 조건부 답변"},
            ],
        },
    ),
]
