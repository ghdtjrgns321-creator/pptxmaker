"""goldenfab.select — 부품·틀을 **세어서** 고르는 규칙 (2026-07-27 신설).

## 왜 이 모듈이 생겼나

부품 선택이 모델 취향으로 넘어가 있었다. 결정표가 "물성으로 판정하라"고 산문으로 적었지만
판정 질문이 배타적이지 않아("3단계로 판정하고 애매하면 사람에게 넘긴다" → 순서·갈림·위임이
모두 참) 같은 내용에 매번 다른 그림이 나왔다.

여기서는 **비교만 한다.** 부품이 `META["accepts"]`로 조건을 선언하고, 재료를 세어 얻은 값과
맞춰본다. 산문 해석이 없으므로 같은 재료에 항상 같은 후보가 나온다.

## 축 5개 — 전부 셀 수 있는 것

    sets    묶음 수 (1 | 2)              — 항목이 한 무리인가 두 무리인가
    flow    방향 (없음|1:N|N:1|N:M|순차)  — 무엇이 무엇으로 가나
    order   순서가 메시지인가 (bool)      — 번호·날짜가 있나
    extra   항목마다 부가 열이 있나 (bool) — 설명·출처·태그가 항목에 딸려 있나
    ends    N개를 감싸는 단일 머리/바닥 (bool) — 1→N→1 구조인가

**실측(2026-07-27):** 부품 9종이 이 5축에서 **8칸**에 갈린다. 남는 겹침은
`gate_branch`↔`n_branch` 하나뿐이고 그건 개수 규칙이 가른다(아래).

축을 더 늘리지 않은 이유: 늘리면 셀 수 없는 것이 섞인다. 줄이면 갈리지 않는다. 이 5개가
지금 창고를 정확히 가르는 최소 집합이다 — 부품이 늘면 그때 다시 잰다.

## 고르는 규칙 — 2단

후보가 여럿이면 아래 순서로 가른다. 사용자 판정은 여기 들어오지 않는다(자동화가 목적).

1. **덱 내 미사용** — 이 덱에서 아직 안 쓴 것을 먼저. 같은 그림 반복을 선택 시점에서 막는다
   (consistency-qa의 사후 다양성 검사를 앞으로 옮긴 것).
2. **개수 적합도** — 1순위가 동점이면(덱 초반엔 전부 미사용이라 항상 동점) 항목 수가 수용
   범위의 **중앙에 가까운** 쪽. 범위 경계에 걸려 헐거워지는 부품이 자동으로 밀린다.

## 후보가 0개면 멈춘다

가장 가까운 것을 변형해 넘기지 않는다. 그렇게 만든 도해는 다음 덱이 상속하지 못하고 한 번
쓰고 버려진다. 멈춘 자리 목록이 곧 **만들어야 할 부품 목록**이다 — 지어내지 않고 수요로 만든다.
"""

import importlib

PARTS = (
    "gate_branch",
    "fan_in",
    "n_branch",
    "hub_spoke",
    "routing_lane",
    "numbered_steps",
    "layer_stack",
    "bipartite_map",
    "relation_catalog",
)
AXES = ("sets", "flow", "order", "extra", "ends")


def catalog():
    """부품 이름 → META. 창고의 단일 출처는 모듈 자신이다(별도 목록을 손으로 만들지 않는다)."""
    out = {}
    for name in PARTS:
        m = importlib.import_module(f"goldenfab.figures.{name}")
        out[name] = m.META
    return out


def shape_of(data):
    """재료에서 5축을 **센다** — 판정이 아니다. 없는 축은 호출자가 채워 넣는다.

    data = {"items": [...], "sets": 1|2, "links": [(a, b), ...], "ordered": bool,
            "extra": bool, "head": any, "tail": any}
    """
    items = data.get("items") or []
    links = data.get("links") or []
    n = len(items)
    if not links:
        flow = "순차" if data.get("ordered") else "없음"
    else:
        lefts = {a for a, _ in links}
        rights = {b for _, b in links}
        if data.get("sets", 1) == 2:
            flow = "N:M"
        elif len(rights) == 1:
            flow = "N:1"
        elif len(lefts) == 1:
            flow = "1:N"
        else:
            flow = "N:M"
    return {
        "n": n,
        "sets": data.get("sets", 1),
        "flow": flow,
        "order": bool(data.get("ordered")),
        "extra": bool(data.get("extra")),
        "ends": bool(data.get("head")) and bool(data.get("tail")),
    }


def _fit(n, rng):
    """개수 적합도 — 수용 범위 중앙에서 얼마나 떨어졌나. 범위 밖이면 None(탈락)."""
    lo, hi = rng
    return None if not (lo <= n <= hi) else abs(n - (lo + hi) / 2)


def candidates(shape, parts=None):
    """조건이 맞는 부품만. 비교만 한다 — 산문 해석 없음."""
    out = []
    for name, meta in (parts or catalog()).items():
        acc = meta.get("accepts")
        if not acc:
            continue
        if any(acc[k] != shape[k] for k in AXES):
            continue
        if _fit(shape["n"], acc["n"]) is None:
            continue
        out.append(name)
    return out


def choose(shape, used=(), parts=None):
    """2단 규칙으로 하나를 고른다. 후보가 없으면 None — 부르는 쪽이 멈추고 보고한다."""
    parts = parts or catalog()
    cands = candidates(shape, parts)
    if not cands:
        return None
    fresh = [c for c in cands if c not in used] or cands  # ① 덱 내 미사용
    return min(fresh, key=lambda c: _fit(shape["n"], parts[c]["accepts"]["n"]))  # ② 개수 적합도


def report(shape, used=(), parts=None):
    """왜 그게 뽑혔는지 — 결정 로그. 선택이 기계였음을 사람이 확인할 수 있어야 한다."""
    parts = parts or catalog()
    cands = candidates(shape, parts)
    if not cands:
        return {"shape": shape, "candidates": [], "pick": None, "why": "후보 0 — 멈추고 보고"}
    fresh = [c for c in cands if c not in used]
    pool, why = (fresh, "미사용") if fresh else (cands, "전부 사용됨 → 개수만")
    pick = min(pool, key=lambda c: _fit(shape["n"], parts[c]["accepts"]["n"]))
    return {
        "shape": shape,
        "candidates": cands,
        "pick": pick,
        "why": f"{why} → 개수 적합도 {_fit(shape['n'], parts[pick]['accepts']['n'])}",
    }
