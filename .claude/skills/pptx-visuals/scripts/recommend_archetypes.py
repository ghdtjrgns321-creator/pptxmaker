#!/usr/bin/env python
"""데이터 형상 → 아키타입 결정적 추천 엔진 (v4.3).

"다양하게"·"알맞게"를 LLM 감각에 맡기지 않고, deck-spec 슬라이드의 실제 데이터에서
특성을 검출해 적합 L-ID를 랭킹한다(archetype-catalog.md 형상 매핑의 기계화).
deck-composer는 이 추천군을 우선 검토하고, check_candidates가 추천군 밖 후보를 경고한다.

usage: python recommend_archetypes.py <deck-spec.json> [--slide N]
출력: 슬라이드별 {no, shape, features, recommended L-ID 랭킹}
"""

import argparse
import json
import re
import sys
from pathlib import Path

FRAME = {"cover", "toc", "part", "section", "cta"}

# 시간축 카테고리 패턴 — 연도·분기·월·버전
TIME_PAT = re.compile(r"^((19|20)\d{2}|Q[1-4]|[1-9][0-2]?월|[vV]\d|\d{4}[-./]\d{1,2})")

# L-ID ↔ deck-spec 형식 (카탈로그와 동기 — 렌더 가능한 것만)
LID_KEY = {
    "L01": "chart:bar",
    "L02": "chart:hbar",
    "L03": "chart:line",
    "L04": "chart:pie",
    "L05": "chart:stacked_bar",
    "L06": "chart:panels",
    "L07": "chart:waterfall",
    "L08": "chart:heatmap",
    "L09": "chart:dumbbell",
    "L10": "chart:slope",
    "L11": "chart:funnel",
    "L12": "chart:annotated_scatter",
    "L13": "chart:histogram",
    # ── 결번 9개(L14·L15·L16·L18·L19·L20·L21·L34·L37) — 박스+화살표 계열 전멸 ──
    # L14 flow는 진행 밴드와 렌더 동일이라 폐기, 나머지 8종은 팔레트 전수 렌더 **전면 반려**로
    # 2026-07-25 폐기(사용자 확정). 그 물성의 정본은 (A) 골든 도해다(archetype-catalog crosswalk).
    # **번호는 밀지 않는다**: 결번을 메우면 문서·워크스페이스의 L-ID 참조 121건이 전부 어긋난다.
    "L17": "diagram:timeline",
    "L22": "diagram:matrix_2x2",
    "L23": "diagram:spectrum",
    "L24": "diagram:harvey_table",
    "L25": "diagram:check_matrix",
    "L26": "table",
    "L27": "table:matrix",
    "L28": "metrics",
    "L29": "two_column",
    "L30": "bullets",
    "L31": "chart:bubble",
    "L32": "diagram:venn",
    "L33": "chart:waffle",
    "L35": "diagram:icon_rows",
    "L36": "diagram:stat_split",
    "L38": "diagram:split_detail",
}

# 형상 → L-ID 랭킹 (앞이 우선 — archetype-catalog.md 매핑과 동일, 순서만 추천 강도)
# 2026-07-25: 폐기 9종을 뺐다. 그래서 "흐름전환"은 L11 퍼널 하나만 남는다 — 분기·라우팅 물성의
# 정본이 (A) 골든 도해(G01·G04·G09·G11)로 옮겨갔기 때문이고, 이 표는 (B) 축만 랭킹한다.
# 결정표(layout-matching)가 (A)를 먼저 보게 하므로 여기 빈 자리가 곧 "(A)로 가라"는 신호다.
SHAPE_RANK = {
    "시계열": ["L03", "L13", "L06"],
    "구성비": ["L07", "L02", "L05", "L33", "L04"],
    "전후": ["L09", "L10", "L29"],
    "범주비교": ["L02", "L01", "L31", "L27", "L09"],
    "교차다축": ["L24", "L08", "L25", "L27"],
    "흐름전환": ["L11"],
    "분포상관": ["L12", "L13"],
    "시간여정": ["L17", "L23"],
    "숫자": ["L36", "L28", "L07"],
    "정성": ["L35", "L38", "L22", "L23", "L24", "L25", "L32"],
}

# ── (A) 골든 문법 우선 랭킹 (2026-07-25) ─────────────────────────────────────────
# 왜 있나: 결정표(layout-matching)는 "(A) 1순위"로 재작성됐는데 **판정 기계는 (B)만 랭킹**해서
# 자동 경로를 타면 여전히 (B)가 유일 후보로 올라왔다("고를 수 있게 만들었지만 그 길로 가는지는
# 산문"). 이 표가 그 산문을 기계로 옮긴다 — 추천 출력의 **맨 앞**이 (A)가 되고,
# check_candidates가 "(A) 후보 0개 + why_not 없음"을 FAIL로 잡는다.
#
# **경계(정직):** (A)가 오르는 물성은 **구조·정성뿐**이다. `detect()`는 수치·카테고리 형상만
# 읽으므로 "판정 인과냐 집합 대응이냐" 같은 구조 물성은 기계가 못 가른다 — 그래서 형상이
# 정성·흐름전환일 때 **후보군 전체**를 올리고 그 중 무엇이냐는 사람/컴포저가 결정표 표 A의
# 판정 질문으로 고른다. 정량 물성(시계열·구성비·전후·범주비교·교차다축·분포상관·숫자·시간여정)은
# 골든 도해에 어휘가 0이므로(인벤토리 §7 실측) 여기 비어 있고 (B)가 정본이다 — 억지로 채우면
# 거짓 추천이 된다.
GRAMMAR_KEY = {  # G-ID → 장 타입 (archetype-catalog (A) 표·layout-matching 표 A와 동기)
    "G01": "golden.problem_grid",
    "G02": "golden.problem_grid",
    "G04": "golden.exec_graph",
    "G05": "golden.tech_evidence",
    "G06": "golden.tech_tree",
    "G08": "golden.tech_tree",
    "G09": "golden.tech_mechanism",
    "G10": "golden.tech_mechanism",
    "G11": "adapted.tech_mechanism",
    "G13": "golden.ab_simulation",
    "G14": "golden.ab_simulation",
    "G15": "golden.validation",
    "G17": "adapted.validation",
    "G19": "golden.boundary",
    "G21": "golden.screenshot",
}
GRAMMAR_RANK = {
    "정성": ["G01", "G05", "G06", "G08", "G19", "G17", "G02", "G14", "G21"],
    "흐름전환": ["G04", "G09", "G11", "G10", "G15", "G13"],
}


def _numbers(sl):
    """슬라이드 spec에서 수치 시리즈를 수집(chart series·values·panels, table 숫자 셀)."""
    if sl.get("panels"):
        merged = {}
        for i, pn in enumerate(sl["panels"]):
            for k, v in (pn.get("series") or {}).items():
                merged[f"{i}:{k}"] = list(v)
        if merged:
            return merged
    if sl.get("series"):
        return {k: list(v) for k, v in sl["series"].items()}
    if sl.get("values") and isinstance(sl["values"], list):
        flat = sl["values"]
        if flat and isinstance(flat[0], list):
            return {f"row{i}": r for i, r in enumerate(flat)}
        return {"values": flat}
    nums = []
    for row in sl.get("rows", []):
        cells = row if isinstance(row, list) else []
        nums += [
            float(m.group()) for c in cells for m in [re.search(r"-?\d+(\.\d+)?", str(c))] if m
        ]
    return {"cells": nums} if nums else {}


def detect(sl):
    """특성 검출 — 전부 결정적 규칙, LLM 판단 없음."""
    series = _numbers(sl)
    all_vals = [v for vs in series.values() for v in vs if isinstance(v, (int, float))]
    cats = [str(c) for c in sl.get("categories", []) or sl.get("columns", []) or []]
    f = {
        "has_numbers": len(all_vals) >= 3,
        "time_axis": sum(bool(TIME_PAT.match(c)) for c in cats) >= max(2, len(cats) - 1)
        and cats != [],
        "sum100": False,
        "pair2": False,
        "grouped": bool(sl.get("groups"))
        or (
            sl.get("rows")
            and isinstance(sl["rows"][0], list)
            and len(sl["rows"][0]) >= 3
            and len(sl.get("rows", [])) >= 3
        ),
        "stagey": sl.get("layout") == "timeline" or bool(sl.get("stages")),
        "kpi": sl.get("type") == "metrics"
        or (
            bool(sl.get("items"))
            and all("value" in i for i in sl.get("items", []) if isinstance(i, dict))
        ),
    }
    single = [s for s in series.values() if all(isinstance(v, (int, float)) for v in s)]
    if len(series) == 1 and single:
        total = sum(single[0])
        f["sum100"] = 95 <= total <= 105
        # 단조 감소 수열 = 전환·이탈 신호
        f["funnel_like"] = (
            all(a >= b for a, b in zip(single[0], single[0][1:])) and len(single[0]) >= 3
        )
    else:
        f["funnel_like"] = False
    two_series = len(series) == 2 and all(
        len(v) == len(list(series.values())[0]) for v in series.values()
    )
    two_cols = len(cats) == 2 and bool(series)
    f["pair2"] = two_series or two_cols
    return f


def classify(sl, f):
    """특성 → 형상 (우선순위 결정적)."""
    if f["kpi"]:
        return "숫자"
    if not f["has_numbers"]:
        if f["stagey"]:
            return "시간여정" if sl.get("layout") == "timeline" else "흐름전환"
        return "정성"
    if f["time_axis"]:
        return "시계열"
    if f["funnel_like"] and (f["stagey"] or sl.get("stages")):
        return "흐름전환"
    if f["pair2"]:
        return "전후"
    if f["sum100"]:
        return "구성비"
    if f["grouped"]:
        return "교차다축"
    return "범주비교"


def recommend(spec):
    """슬라이드별 추천 — **(A) 골든 문법이 먼저**, 그다음 (B) L-ID(결정표 판정 절차 §2 그대로)."""
    out = []
    for i, sl in enumerate(spec["slides"], 1):
        if sl["type"] in FRAME:
            continue
        f = detect(sl)
        shape = classify(sl, f)
        rec = [
            {"gid": gid, "key": GRAMMAR_KEY[gid], "axis": "A"}
            for gid in GRAMMAR_RANK.get(shape, [])
        ]
        rec += [{"lid": lid, "key": LID_KEY[lid], "axis": "B"} for lid in SHAPE_RANK[shape]]
        out.append(
            {
                "no": i,
                "type": sl["type"],
                "shape": shape,
                "features": {k: v for k, v in f.items() if v},
                "grammar_first": bool(GRAMMAR_RANK.get(shape)),  # (A)가 1순위인 물성인가
                "recommended": rec,
            }
        )
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--slide", type=int)
    args = ap.parse_args()
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    recs = recommend(spec)
    for r in recs:
        if args.slide and r["no"] != args.slide:
            continue
        keys = " > ".join(f"{x.get('gid') or x['lid']}({x['key']})" for x in r["recommended"])
        head = "(A)먼저 " if r["grammar_first"] else ""
        print(f"s{r['no']:02d} [{head}{r['shape']}] {sorted(r['features'])} → {keys}")


if __name__ == "__main__":
    sys.exit(main())
