#!/usr/bin/env python
"""익스히빗 후보 JSON 기계 검증기 — deck-composer가 보고 전에 스스로 돌리는 자기 심사.

검사: ① 커버 전수(모집단 = deck-spec 본문, 양방향 차집합) ② 슬라이드당 후보 ≥3안
③ 슬라이드 내 상이 유형 ≥2종(3안째는 조합 변주 허용) ④ 조합 장치 ≥1 전수
⑤ shape 필드 전수 ⑥ 디폴트 어휘(bar·flow·cards) 후보의 why_not 의무
⑦ 아키타입 히스토그램 + 최선 조합 게이트 시뮬레이션(단일 유형 30% 초과 경고)

usage: python check_candidates.py <candidates.json> <deck-spec.json>
exit 0=PASS, 1=FAIL
"""

import json
import sys
from collections import Counter
from itertools import product
from pathlib import Path

FRAME = {"cover", "toc", "part", "section", "cta"}
COMBO = ("emphasis", "annotations", "sub_table", "banner", "ref")
DEFAULTS = {"chart:bar", "diagram:flow", "diagram:cards"}
SHAPES = {
    "시계열",
    "범주비교",
    "구성비",
    "흐름전환",
    "전후",
    "교차다축",
    "분포상관",
    "정성",
    "시간여정",
    "숫자",
}

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "consistency-qa" / "scripts"))
sys.path.insert(0, str(Path(__file__).parent))
from audit_pptx import diversity_checks  # noqa: E402
from recommend_archetypes import recommend  # noqa: E402


def key(c):
    if c.get("panels"):
        return "chart:panels"
    if c.get("chart_type"):
        return f"chart:{c['chart_type']}"
    if c.get("layout") or c.get("type") == "diagram":
        return f"diagram:{c.get('layout', 'flow')}"
    return c.get("type", "?")


def main():
    cand = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    spec = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
    body = {i for i, sl in enumerate(spec["slides"], 1) if sl["type"] not in FRAME}
    fails, warns = [], []
    # 추천 엔진(데이터 형상 결정적 분석) — 추천군 밖 후보는 경고(정당화 여지는 남김)
    rec_by_no = {r["no"]: {x["key"] for x in r["recommended"]} for r in recommend(spec)}

    covered = {sl["no"] for sl in cand["slides"]}
    if covered != body:
        fails.append(
            f"커버 불일치: 미커버 {sorted(body - covered)}, 모집단 밖 {sorted(covered - body)}"
        )

    hist = Counter()
    for sl in cand["slides"]:
        no, cands = sl["no"], sl.get("visual_candidates", [])
        keys = [key(c) for c in cands]
        hist.update(set(keys))
        if len(cands) < 3 and not sl.get("fixed"):
            fails.append(f"s{no}: 후보 {len(cands)}개 < 3 (fixed 아님)")
        if len(set(keys)) != len(keys):
            # v4.3: 조합 변주 허용 조항 삭제 — 슬라이드 내 동일 유형 후보는 숫자 채우기
            fails.append(f"s{no}: 슬라이드 내 후보 유형 중복 {keys}")
        if not sl.get("shape") and not all(c.get("shape") for c in cands):
            fails.append(f"s{no}: shape 필드 없음(슬라이드 또는 후보 단위 필수)")
        for i, c in enumerate(cands):
            tag = chr(ord("A") + i)
            if not any(c.get(f) for f in COMBO) and not c.get("intro") and not c.get("commentary"):
                fails.append(f"s{no}-{tag}: 조합 장치 0개 ({key(c)})")
            if key(c) in DEFAULTS and not c.get("why_not"):
                fails.append(f"s{no}-{tag}: 디폴트 어휘 {key(c)}인데 why_not 없음")
            rec = rec_by_no.get(no)
            if rec and key(c) not in rec and not c.get("why_not"):
                warns.append(
                    f"s{no}-{tag}: 추천군 밖 {key(c)} (추천: {sorted(rec)}) — why_not 권장"
                )
        cur = sl.get("current", "")
        for i, c in enumerate(cands):
            if (
                key(c).endswith(cur)
                and cur
                and not any(c.get(f) for f in COMBO)
                and not sl.get("fixed")
            ):
                fails.append(f"s{no}-{chr(ord('A') + i)}: 맨 현행 재출품({cur})")

    # 최선 조합 게이트 시뮬레이션: 각 슬라이드 후보 중 게이트 통과 조합이 존재하는가
    order = sorted(cand["slides"], key=lambda s: s["no"])
    options = [list({key(c) for c in sl["visual_candidates"]}) for sl in order]
    sim_pass = False
    n_combo = 1
    for o in options:
        n_combo *= len(o)
    if n_combo <= 200_000:
        for pick in product(*options):
            slides = [
                {"type": "chart", "chart_type": k.split(":")[1]}
                if k.startswith("chart:")
                else {"type": "diagram", "layout": k.split(":")[1]}
                if k.startswith("diagram:")
                else {"type": k}
                for k in pick
            ]
            if all(ok for _, ok, _ in diversity_checks(slides)):
                sim_pass = True
                break
    else:
        warns.append(f"조합 {n_combo}개 — 시뮬레이션 생략")
        sim_pass = True
    if not sim_pass:
        fails.append("게이트 통과 가능한 선택 조합이 존재하지 않음 — 후보 구성 재설계 필요")

    total = sum(len(sl["visual_candidates"]) for sl in cand["slides"])
    print(f"슬라이드 {len(cand['slides'])}/{len(body)} · 후보 {total}안")
    print(
        "아키타입 히스토그램(슬라이드 보유 기준):",
        ", ".join(f"{k}×{v}" for k, v in hist.most_common()),
    )
    for wmsg in warns:
        print("[warn]", wmsg)
    for fmsg in fails:
        print("[FAIL]", fmsg)
    print("RESULT:", "PASS" if not fails else f"FAIL({len(fails)}건)")
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
