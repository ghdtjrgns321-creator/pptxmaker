"""goldenfab.frames — 3층 배치 틀 (2026-07-27 신설).

## 왜 이 모듈이 생겼나

슬라이드 한 장은 그림 하나가 아니다 — 제목·핵심 메시지·설명 패널·수치 띠·그림의 크기와 자리가
함께 있다. 도해를 `figures/`로 꺼낸 뒤에도 **이 나머지가 장 함수에 굳어 있으면** 컴포저는 여전히
"골든 8번 장 통째로"밖에 배정할 수 없다. 실전 덱 타입 분포가 골든 `SLIDE_ORDER`와 사실상
동일했던 실측이 그 결과다(= 골든 복제).

그래서 자리를 장에서 떼어 **틀**로 만든다. 틀은 그림을 모르고 자리만 안다.

    한 장 = 배치 틀 1개 + 그림 N개 + 글

## 틀은 새로 디자인하지 않는다

전부 골든 17장에서 **실측으로 추출**했다. 승인된 배치를 꺼내는 것이므로 반려 위험이 0이고,
골든이 "조합 사례"라는 정의(PIPELINE.md)가 여기서 실제로 쓰인다.

`twin_top_cards`가 골든에서 3장(S8·S9·S11)에 쓰였다 — 카드 띠 좌표가 셋 다 같다. 이것이
이 공장의 주력 틀이다.

## 틀 고르는 법 — 재료가 정한다

틀마다 **채워야 하는 칸**이 있고 그 칸을 채울 재료가 그 장에 있어야 성립한다. `needs`가 그
선언이고, 컴포저는 재료를 세어 `needs`를 만족하는 틀만 후보로 올린다(산문 해석 없음).
그림 고를 때와 같은 방식이다 — 조건 선언 → 세서 후보 → ① 덱 내 미사용 ② 개수 적합도.

## 좌우 분할은 틀이 정하지 않는다

두 자리 틀에서 좌우 폭은 **그림이 요구한 크기**(`measure`)로 가른다. 틀이 비율을 박으면 긴
라벨을 가진 도해가 잘리고, 반대로 자리가 남는 도해는 헐거워진다. 골든에서도 S8은 좌 7.95",
S9는 좌 4.4"로 장마다 달랐다 — 그 차이는 내용에서 온 것이지 틀에서 온 것이 아니다.
"""

from . import dense as D
from . import grid as G
from .figures import Box

# ── 공통 자리 (골든 전 장 공유) ──────────────────────────────────────────────────
HEADER_H = 0.76  # compact_header가 쓰는 상단 의식
BAND_TOP = 0.80  # 본문 밴드 top (소제목이 여기서 시작)
FIG_TOP = 1.14  # 소제목 룰 밑 — 도해가 실제로 시작하는 y
CARD_Y, CARD_H = 3.62, 3.32  # 하단 카드 띠 — S8·S9·S11이 **같은 값**을 쓴다
CARD_GAP = 0.18
FULL_W = D.FULL_W  # 12.133
LEFT = G.MARGIN_L  # 0.60
RIGHT = G.RIGHT_EDGE  # 12.733


def _cards(n, y=CARD_Y, h=CARD_H, gap=CARD_GAP, floor=2.2):
    """카드 띠 — 폭은 장수에서 파생한다. 안 들어가면 `grid.track`이 시끄럽게 죽는다."""
    w = G.track(n, LEFT, RIGHT, gap, floor, what="카드")
    return [Box(LEFT + i * (w + gap), y, w, h) for i in range(n)]


def _split(band, sizes, gutter=0.30):
    """두 자리로 가른다 — 폭은 **그림이 요구한 크기**의 비율로. 틀이 비율을 박지 않는다.

    sizes = [(필요폭, ...), (필요폭, ...)] — 각 도해의 `measure()` 첫 값.
    """
    need = [max(s, 1.0) for s in sizes]
    avail = band.w - gutter
    total = sum(need)
    if total > avail:  # 둘 다 자리를 넘으면 조용히 줄이지 않고 죽는다
        raise ValueError(
            f'두 도해가 요구한 폭 {total:.2f}"가 자리 {avail:.2f}"를 넘는다 — '
            "라벨을 줄이거나 한 장으로 나눈다"
        )
    scale = avail / total
    lw = need[0] * scale
    return (
        Box(band.x, band.y, lw, band.h),
        Box(band.x + lw + gutter, band.y, avail - lw, band.h),
    )


# ── 틀 5종 (전부 골든 실측) ──────────────────────────────────────────────────────
# `needs`: 이 틀이 성립하려면 재료에 무엇이 있어야 하나. 컴포저가 세는 대상.
#   figures — 도해 자리 수 · cards — 카드 장수 범위 · panel/columns — 보조 재료
FRAMES = {
    "twin_top_cards": {
        "설명": "상단에 도해 둘을 나란히, 하단에 카드 띠. 이 공장의 주력.",
        "출처": "골든 S8(매핑+JSON) · S9(트리+카탈로그) · S11(흐름+트리)",
        "needs": {"figures": 2, "cards": (3, 5)},
        "band": Box(LEFT, FIG_TOP, FULL_W, CARD_Y - FIG_TOP - 0.32),
    },
    "wide_top_cards": {
        "설명": "상단 도해 하나가 전폭을 쓰고 하단에 카드 띠.",
        "출처": "골든 S16(경계 캔버스 + 한계 카드)",
        "needs": {"figures": 1, "cards": (2, 5)},
        "band": Box(LEFT, FIG_TOP, FULL_W, 2.20),
        "cards_y": 3.46,
        "cards_h": 3.00,
    },
    "stacked_bands": {
        "설명": "도해 둘을 위아래로. 각 밴드에 소제목이 붙는다.",
        "출처": "골든 S4(판정 갈림 + 한계 수렴)",
        "needs": {"figures": 2, "band_heads": 2},
        "bands": [Box(LEFT, 1.29, FULL_W, 2.50), Box(LEFT, 4.66, 7.95, 2.12)],
        "panel": Box(8.55, 4.30, RIGHT - 8.55, 2.36),  # 우하단 보조 패널(선택)
    },
    "figure_and_columns": {
        "설명": "도해가 전폭을 쓰고 그 아래 설명 칼럼 N개.",
        "출처": "골든 S6(실행 레인 + 핵심 기술 4칼럼)",
        "needs": {"figures": 1, "columns": (3, 4)},
        "band": Box(LEFT, 2.23, FULL_W, 1.78),
        "col_y": 4.85,
        "col_h": 1.05,
        "col_gap": 0.35,
    },
    "mirror_split": {
        "설명": "좌우 대칭 대비 — 같은 문법을 두 벌 나란히 놓아 차이만 보이게.",
        "출처": "골든 S14(A/B 시뮬레이션)",
        "needs": {"figures": 2, "titles": 2},
        "band": Box(LEFT, 1.30, FULL_W, 4.60),
        "symmetric": True,  # 좌우 폭을 measure로 가르지 않고 **똑같이** 준다
    },
}


def candidates(material):
    """재료로 성립하는 틀만 고른다 — 산문 해석 없이 `needs`와 개수를 비교만 한다.

    material = {"figures": n, "cards": n, "columns": n, "band_heads": n, "titles": n, ...}
    """
    out = []
    for name, f in FRAMES.items():
        ok = True
        for key, want in f["needs"].items():
            have = material.get(key, 0)
            if isinstance(want, tuple):
                ok &= want[0] <= have <= want[1]
            else:
                ok &= have == want
        if ok:
            out.append(name)
    return out


def slots(name, sizes=None, n_cards=4):
    """틀 이름 → 실제 자리(Box)들. sizes = 두 자리 틀에서 각 도해의 필요 폭."""
    f = FRAMES[name]
    if name == "twin_top_cards":
        left, right = _split(f["band"], sizes) if sizes else (f["band"], f["band"])
        return {"figures": [left, right], "cards": _cards(n_cards)}
    if name == "wide_top_cards":
        return {
            "figures": [f["band"]],
            "cards": _cards(n_cards, f["cards_y"], f["cards_h"]),
        }
    if name == "stacked_bands":
        return {"figures": list(f["bands"]), "panel": f["panel"]}
    if name == "figure_and_columns":
        w = G.track(n_cards, LEFT, RIGHT, f["col_gap"], 2.0, what="설명 칼럼")
        cols = [
            Box(LEFT + i * (w + f["col_gap"]), f["col_y"], w, f["col_h"]) for i in range(n_cards)
        ]
        return {"figures": [f["band"]], "columns": cols}
    if name == "mirror_split":
        b = f["band"]
        half = (b.w - 0.30) / 2
        return {
            "figures": [
                Box(b.x, b.y, half, b.h),
                Box(b.x + half + 0.30, b.y, half, b.h),
            ]
        }
    raise KeyError(f"모르는 틀: {name}")
