#!/usr/bin/env python
"""선택 규칙 회귀 검증 — 골든이 손으로 고른 그림을 규칙이 재현하나 (2026-07-27 신설).

    uv run python .claude/skills/pptx-build/scripts/verify_selection.py

## 왜 이 파일이 있나

부품 선택을 "기계가 고른다"고 **선언**하는 것과 실제로 같은 답이 나오는 것은 다르다. 이
프로젝트는 그 차이로 여러 번 데었다(결정표가 "(A) 먼저 보라"고 산문으로 적었는데 판정 기계는
(B)만 후보로 올렸다). 그래서 주장을 문장이 아니라 **실행 가능한 대조**로 둔다.

기준선(oracle)은 **골든덱이 실제로 그린 도해**다. 사람이 장마다 손으로 고른 결과이고 눈검증을
통과했다. 규칙이 재료만 세어 같은 답을 내면 그 규칙은 사람의 선택을 재현한 것이다.

검사 3종:
  ① 재현     — 골든 8장의 도해 선택을 규칙이 그대로 맞추나
  ② 분리     — 부품 N종이 축 5개에서 몇 칸에 갈리나(겹침은 개수 규칙이 가르는지까지)
  ③ 멈춤     — 창고에 없는 물성이 후보 0으로 **멈추나**(가장 가까운 것으로 때우지 않는다)

②가 중요한 이유: 부품이 늘면 새 부품이 기존 부품과 같은 칸에 들어갈 수 있다. 그때 이 검사가
"축이 부족해졌다"를 알려준다 — 축은 지어내는 게 아니라 창고를 못 가르게 됐을 때 늘린다.
"""

import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from goldenfab import select as SEL  # noqa: E402

# ── ① 기준선: 골든이 실제로 그린 도해 (장 → 재료, 기대 부품) ──────────────────────
# 재료는 그 장의 content에서 **셀 수 있는 것만** 뽑은 것이다. 부품 이름을 보고 역산하지 않았다.
GOLDEN = [
    (
        "S4 상단 판정 갈림",
        dict(items=[1, 2], links=[("q", "a"), ("q", "b")], extra=True),
        "gate_branch",
    ),
    (
        "S4 하단 한계 수렴",
        dict(items=[1, 2, 3, 4], links=[(i, "g") for i in range(4)], extra=True),
        "fan_in",
    ),
    ("S6 실행 레인", dict(items=list(range(7)), ordered=True, extra=True), "routing_lane"),
    ("S8 용어 매핑", dict(items=list(range(5)), sets=2, links=[(1, 2)] * 5), "bipartite_map"),
    (
        "S9 간선 카탈로그",
        dict(items=list(range(7)), sets=2, links=[(1, 2)] * 7, extra=True),
        "relation_catalog",
    ),
    ("S11 주입 3단", dict(items=[1, 2, 3], ordered=True), "numbered_steps"),
    (
        "S11 갈래 3",
        dict(items=[1, 2, 3], links=[("q", i) for i in range(3)], extra=True),
        "n_branch",
    ),
    ("층 구조 1→3→1", dict(items=[1, 2, 3], ordered=True, head="정의", tail="수렴"), "layer_stack"),
]

# ── ③ 창고에 없는 물성 — 멈춰야 하는 재료 ───────────────────────────────────────
# 결정표의 ⬜ 빈칸(편차·공간·정성 2열 대비)에 해당한다. 후보가 나오면 오히려 잘못이다.
NO_MATCH = [
    ("편차 — 기준선 대비 +/-", dict(items=list(range(12)))),
    ("공간 — 지리 위치", dict(items=list(range(20)))),
]

FAILS = []


def _ok(name, cond, detail):
    mark = "OK  " if cond else "✘FAIL"
    print(f"{mark} {name}\n       {detail}")
    if not cond:
        FAILS.append(name)


def test_reproduce():
    """① 골든이 손으로 고른 선택을 규칙이 재현하나 — 덱 순서대로(미사용 규칙 포함)."""
    used, wrong = [], []
    for label, data, want in GOLDEN:
        got = SEL.choose(SEL.shape_of(data), used=used)
        if got != want:
            wrong.append(f"{label}: {got} ≠ {want}")
        if got:
            used.append(got)
    _ok(
        "① 재현 — 골든 도해 선택을 재료만 세어 맞추나",
        not wrong,
        f"{len(GOLDEN)}장 전부 일치" if not wrong else f"불일치 {wrong}",
    )


def test_separation():
    """② 축 5개가 창고를 가르나 — 겹치는 칸은 개수 규칙으로 갈려야 한다."""
    cat = SEL.catalog()
    cells = defaultdict(list)
    for name, meta in cat.items():
        a = meta["accepts"]
        cells[tuple(a[k] for k in SEL.AXES)].append(name)
    dups = {k: v for k, v in cells.items() if len(v) > 1}
    # 겹치는 칸은 개수 범위가 달라야 한다(같으면 2순위로도 못 가른다)
    unresolved = [
        v for v in dups.values() if len({tuple(cat[n]["accepts"]["n"]) for n in v}) < len(v)
    ]
    _ok(
        "② 분리 — 축 5개로 부품이 칸에 갈리나 (겹침은 개수로 해소)",
        not unresolved,
        f"부품 {len(cat)}종 → {len(cells)}칸 · 겹침 {len(dups)}개"
        + (f" (개수로 해소: {[v for v in dups.values()]})" if dups else "")
        + (f" · **미해소 {unresolved}**" if unresolved else ""),
    )


def test_stop():
    """③ 창고에 없는 물성은 멈추나 — 가장 가까운 것으로 때우면 다음 덱이 상속하지 못한다."""
    leaked = [
        f"{label} → {SEL.choose(SEL.shape_of(d))}"
        for label, d in NO_MATCH
        if SEL.choose(SEL.shape_of(d)) is not None
    ]
    _ok(
        "③ 멈춤 — 없는 물성이 후보 0으로 멈추나 (억지 매칭 0)",
        not leaked,
        f"{len(NO_MATCH)}건 전부 멈춤 — 그 목록이 만들 부품 목록이다"
        if not leaked
        else f"샘 {leaked}",
    )


def main():
    print("═" * 70)
    print("선택 규칙 회귀 검증 — 기준선 = 골든이 손으로 고른 도해")
    print("═" * 70)
    test_reproduce()
    test_separation()
    test_stop()
    print("═" * 70)
    if FAILS:
        print(f"SELECTION FAIL {len(FAILS)}건 — {FAILS}")
        return 1
    print("SELECTION PASS — 규칙이 사람의 선택을 재현한다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
