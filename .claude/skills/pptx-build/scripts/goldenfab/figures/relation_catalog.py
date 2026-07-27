"""relation_catalog — 관계 종류를 전수로 펼치고 각 줄에 근거 출처를 붙인다 (골든 s09, 2026-07-26).

출처: `s09_variants.build_catalog`(골든 S9 우측)를 부품 계약 3조건으로 옮긴 것. 좌표·도형·색은
골든 실측 그대로여서 `compare_golden` 픽셀 동일로 무손실을 증명한다.

## 이 도해가 무엇을 말하나

"무엇이 무엇을 잇는가"를 **종류별로 한 줄씩** 전수로 편다. 각 줄이 실제 노드 두 개와 그 사이
화살표로 그려지므로, 표로 적었을 때와 달리 **관계가 형태로** 보인다. 오른쪽 끝의 출처 열이
"이 관계를 어디서 가져왔나"를 같은 줄에서 답한다 — 관계와 근거가 떨어지면 주장만 남는다.

## 정렬은 폭이 아니라 화살표 구간으로 잡는다

칩 폭을 균일하게 두면 짧은 라벨이 빈 상자로 보인다(채움률 21~26%로 오딧 7건이 걸렸던 자리).
그래서 칩 폭은 글자에서 파생하고, 대신 **A는 우변 고정 · B는 좌변 고정**으로 둔다 — 화살표가
모든 줄에서 같은 자리·같은 길이라 열이 기둥으로 읽힌다.

## 행 수는 자유다

피치를 행 수에서 파생한다(`grid.pitch`). 자리를 넘으면 조용히 겹치지 않고 시끄럽게 죽는다.
"""

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from ..kit import add_text
from . import elements as E
from . import merged, require

META = {
    "name": "relation_catalog",
    "어휘": "G08",  # archetype-catalog 등재 어휘(relation_catalog)와 대응
    "물성": "무엇이 무엇을 잇고 그 근거가 어디서 왔나(관계 종류 전수 + 출처)",
    "arity": "파생 — 행 수가 피치를 정한다. 자리를 넘으면 시끄럽게 죽는다",
    "keys": ["rows"],
    "optional": [],
}

LAYOUT = {
    # box 원점 기준 offset(inch). 골든 s09 실측값.
    "row_h": 0.26,  # 행 노드 높이 — 계열 균일
    "kind_w": 1.05,  # 종류·수 앵커 열
    "a_right_dx": 2.77,  # A칩 **우변** — 화살표 출발점이 모든 행에서 같아야 기둥으로 읽힌다
    "arrow_w": 1.05,
    "b_max_w": 1.85,  # B칩 상한 — 넘으면 출처 열을 밀어낸다
    "src_gap": 0.30,  # B칩 상한 → 출처 열
    "inset": 0.44,  # 칩 글자 좌우 여백
    "floor": 0.95,  # 최소 칩 폭
    "label_dy": -0.20,  # 화살표 중심 → 관계 라벨 top (선 **위**에 얹는다)
    "label_h": 0.17,
    "text_dy": 0.01,  # 행 top → 좌우 텍스트 top
    "text_h": 0.24,
}


def measure(data, kit, layout=None):  # noqa: ARG001 — 계약 시그니처 고정
    """필요 폭·높이(inch). 높이는 행 수에서 나온다."""
    from .. import grid as G

    L = merged(LAYOUT, layout)
    n = len(data["rows"])
    w = L["a_right_dx"] + L["arrow_w"] + L["b_max_w"] + L["src_gap"] + 1.5
    return w, (n - 1) * G.CAP_PITCH + L["row_h"]


def draw(slide, box, data, kit, layout=None):
    """box 안에 관계 카탈로그를 그린다. 반환: 행 수(오딧용).

    data = {"rows": [{"kind","count","a","a_kind","label","b","b_kind","src"}, ...]}
    """
    from .. import grid as G

    require(data, META["keys"], META["name"])
    L = merged(LAYOUT, layout)
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    rows, h, pt = data["rows"], L["row_h"], kit["sizes"]["foot"]

    a_right = box.x + L["a_right_dx"]
    b_left = a_right + L["arrow_w"]
    src_x = b_left + L["b_max_w"] + L["src_gap"]
    src_w = box.x + box.w - src_x

    pitch = G.pitch(len(rows), box.y, box.y + box.h - h, h, what="카탈로그 행")
    for i, r in enumerate(rows):
        y = box.y + i * pitch
        cy = y + h / 2
        add_text(  # 종류 + 수 — 수치는 관계에 종속된 보조다(좌측 앵커)
            slide,
            box.x,
            y + L["text_dy"],
            L["kind_w"],
            L["text_h"],
            [
                (f"{r['kind']} ", {"color": C["primary"], "bold": True}),
                (r["count"], {"color": C["muted"]}),
            ],
            S["foot"],
            F["body"],
            C["muted"],
            anchor=MSO_ANCHOR.MIDDLE,
        )
        wa = E.chip_w(r["a"], pt, L["inset"], L["floor"])
        wb = min(E.chip_w(r["b"], pt, L["inset"], L["floor"]), L["b_max_w"])
        E.node_chip(slide, kit, a_right - wa, y, wa, h, r["a"], r["a_kind"])  # A는 우변 고정
        E.node_chip(slide, kit, b_left, y, wb, h, r["b"], r["b_kind"])  # B는 좌변 고정
        E.arrow(slide, kit, a_right, cy, b_left, cy)
        add_text(  # 관계 라벨은 화살표 **위** — 선 위에 얹으면 글자가 선에 잘린다
            slide,
            a_right,
            cy + L["label_dy"],
            L["arrow_w"],
            L["label_h"],
            r["label"],
            S["foot"] - 1,
            F["body"],
            C["muted"],
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            src_x,
            y + L["text_dy"],
            src_w,
            L["text_h"],
            r["src"],
            S["foot"] - 1,
            F["body"],
            C["muted"],
            anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.05,
        )
    return len(rows)


SAMPLES = [
    (
        "최소 2행",
        {
            "rows": [
                {
                    "kind": "계층",
                    "count": "1,204",
                    "a": "변동대가",
                    "a_kind": "concept",
                    "label": "하위",
                    "b": "문단 50",
                    "b_kind": "para",
                    "src": "기준서 소제목 계층",
                },
                {
                    "kind": "참조",
                    "count": "418",
                    "a": "문단 56",
                    "a_kind": "para",
                    "label": "참조",
                    "b": "B33",
                    "b_kind": "para",
                    "src": "본문 상호참조",
                },
            ]
        },
    ),
]
