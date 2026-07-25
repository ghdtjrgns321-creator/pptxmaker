"""배선 스모크 테스트 — 파이프라인 연결이 살아있는지 기계로 확인한다.

왜 있나(2026-07-20): "연결했다"고 선언한 게 실제로는 끊겨 있는 일이 반복됐다(dense가
registry에 안 물림·게이트가 빌드에 안 걸림·다양성 게이트가 golden 타입에 no-op). 손으로 매번
추적하는 건 못 미더우니, 연결 불변식(invariant)을 여기 박아 회귀를 자동으로 잡는다.

    uv run python .claude/skills/pptx-build/scripts/test_wiring.py

각 검사 = 파이프라인의 한 연결점. 종료코드 1 = 어딘가 끊김. green = 전 체인 연결됨.
"""

import importlib
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]  # 프로젝트 루트
sys.path.insert(0, str(HERE))

RESULTS = []


def check(name, ok, detail=""):
    RESULTS.append((name, bool(ok), detail))


def _read(rel):
    p = ROOT / rel
    return p.read_text(encoding="utf-8") if p.exists() else ""


# ── ① dense 건강 + 골든 승격 상태 (dense가 실제 빌드에 나올 준비가 됐나) ──────────
# 현 상태(2026-07-20): dense는 골든 **승격 대기**. 승격(registry 재지정)은 골든 회귀 하네스가
# sparse variant 기하에 박제돼 있어 선행 작업이 필요하다 — audit.py를 dense 인지로(accent 런
# 제외·경계 6.9), compare_golden/audit_golden 장별 오라클을 dense 좌표로, s15·s16 완결. 그래서
# 이 테스트는 "승격됐나"가 아니라 **"dense가 건강하고 승격 준비됐나 + 반쯤 승격해 하네스 안
# 깨졌나"**를 본다(상시 적색 게이트는 무시당하므로). 승격이 끝나면 승격분이 registry에 뜬다.
# WIP dense — 하드코딩·content 미파라미터화·미승인. 아직 계약키(DEFAULT) 불필요, 빌드만 확인.
# 완결(파라미터화)되면 여기서 뺀다 → 그때부터 이 테스트가 DEFAULT 노출을 강제.
WIP_DENSE = (
    set()
)  # s16 hero_card 재건·§8 통과로 완결(2026-07-21). 남은 dense 미완은 exec_graph(S6) 신규 예정


def test_dense_ready():
    from goldenfab.registry import LAYOUTS

    reg = {fn.__module__.rsplit(".", 1)[-1] for fn in LAYOUTS.values()}
    dense = sorted(p.stem for p in (HERE / "goldenfab").glob("s*_dense.py"))
    unhealthy = []
    for m in dense:
        mod = importlib.import_module(f"goldenfab.{m}")
        n_default = sum(
            1 for n, v in vars(mod).items() if "DEFAULT" in n.upper() and isinstance(v, dict)
        )
        if not hasattr(mod, "build"):
            unhealthy.append(f"{m}(build 없음)")
        elif m not in WIP_DENSE and n_default != 1:  # 완결 dense = 계약키 dict 정확히 1개 필수
            unhealthy.append(f"{m}(DEFAULT={n_default})")
    ready = sorted(m for m in dense if m not in WIP_DENSE)
    promoted = sorted(m for m in dense if m in reg)
    check(
        "① dense 건강+승격준비 (완결=빌드+계약키1 · WIP=빌드 · 반쯤승격 없음)",
        not unhealthy,
        f"불건강 {unhealthy}"
        if unhealthy
        else f"승격준비 {len(ready)}장 · WIP {sorted(WIP_DENSE)} · 현 승격 "
        + (str(promoted) if promoted else "0(골든=variant, 하네스 재조정 후 승격 예정)"),
    )


# ── ② 밀도 게이트 → 빌드 (오케스트레이터 없이도 빌드가 스스로 채점하나) ──────────
def test_gate_mechanical():
    src = _read(".claude/skills/pptx-build/scripts/build_pptx.py")
    # 빌드 경로가 밀도 게이트를 호출하는지(메서드 존재 + build에서 호출)
    has_gate = "_gate_density" in src
    called = src.count("_gate_density") >= 2  # 정의 + 호출
    check(
        "② 밀도 게이트→빌드 (build_pptx가 스스로 밀도 채점)",
        has_gate and called,
        "build_pptx._gate_density 정의+호출 확인"
        if (has_gate and called)
        else "build_pptx에 빌드시 밀도 게이트 없음(consistency-qa 소프트 단계에만 의존)",
    )


# ── ③ 다양성 게이트 → golden 타입 인식 (golden 덱에 no-op 아닌가) ────────────────
def test_diversity_golden_aware():
    sys.path.insert(0, str(ROOT / ".claude/skills/consistency-qa/scripts"))
    ap = importlib.import_module("audit_pptx")
    # golden 프레임 장은 본문에서 제외돼야
    frame_excluded = all(
        ap.is_frame({"type": t})
        if hasattr(ap, "is_frame")
        else (t in getattr(ap, "FRAME_TYPES", set()))
        for t in ("golden.cover", "golden.toc", "golden.part", "golden.closing")
    )
    # golden 본문 장은 의미있는 시각 키를 받아야(그냥 유일 문자열 반환이면 no-op)
    k1 = ap.visual_key({"type": "golden.problem_grid"})
    k2 = ap.visual_key({"type": "golden.exec_graph"})
    meaningful = k1 != "golden.problem_grid" or hasattr(ap, "GOLDEN_KIND")
    check(
        "③ 다양성 게이트→golden 인식 (프레임 제외 + 본문 계수)",
        frame_excluded and meaningful,
        f"프레임제외={frame_excluded} 본문키={k1!r},{k2!r}",
    )


# ── ④ 즉흥 밋밋 카드 게이트 배선 (hero_card 미사용을 기계로 막나) ────────────────
def test_adhoc_card_gate_wired():
    from goldenfab import audit as A

    # 검사가 존재하고 dense 경로에만 배선됐나(존재≠검사 hollow 방지).
    has_fn = hasattr(A, "check_adhoc_card")
    dense_names = [r[0] for r in A.generic_checks([], "D66E3A", dense=True)]
    sparse_names = [r[0] for r in A.generic_checks([], "D66E3A", dense=False)]
    wired = "§8 즉흥 카드" in dense_names and "§8 즉흥 카드" not in sparse_names
    check(
        "④ 즉흥 카드 게이트→generic_checks(dense) (밋밋 카드 기계 차단)",
        has_fn and wired,
        f"check_adhoc_card 존재={has_fn} · dense 배선={wired}"
        if (has_fn and wired)
        else f"미배선(존재={has_fn}, dense={('§8 즉흥 카드' in dense_names)})",
    )


# ── ④.5 텍스트 겹침 게이트 배선 (배경없는 세로 파일업을 generic_checks 전 경로에서 막나) ──
def test_text_collision_gate_wired():
    """2026-07-24: 텍스트 요소끼리 겹침(카드 목업 3회 반려)이 오딧 사각지대였다 — 하강.
    check_text_overflow(제 상자)·check_picture_overlap(그림)의 형제. **공유 res**라 dense·sparse
    양 경로에 물려야 한다(존재≠검사 hollow 방지 + 한쪽만 물리면 실전/골든 한쪽이 샌다)."""
    from goldenfab import audit as A

    has_fn = hasattr(A, "check_text_collision")
    dense_names = [r[0] for r in A.generic_checks([], "D66E3A", dense=True)]
    sparse_names = [r[0] for r in A.generic_checks([], "D66E3A", dense=False)]
    wired_both = "텍스트 겹침" in dense_names and "텍스트 겹침" in sparse_names
    check(
        "④.5 텍스트 겹침 게이트→generic_checks(양 경로) (배경없는 세로 파일업 기계 차단)",
        has_fn and wired_both,
        f"check_text_collision 존재={has_fn} · dense·sparse 양쪽 배선={wired_both}"
        if (has_fn and wired_both)
        else f"미배선(존재={has_fn}, dense={'텍스트 겹침' in dense_names}, "
        f"sparse={'텍스트 겹침' in sparse_names})",
    )


# ── ⑤ 산출물→덱 에이전트 게이트 (pptx 저장 감지 → deck-smith "완료 실행" 경유 강제) ──
def test_artifact_agent_gate_wired():
    """2026-07-24 전수 감사: 게이트가 경로(goldenfab 소스)만 봐서 스크래치 pptx 작업이
    무검증 통과하던 구멍. 훅 3점(matcher·arm 감지·게이트 검사)이 전부 물려 있어야 한다.
    페어링 교체(2026-07-24): 옛 게이트는 subagent_type "문자열 존재"만 봐서 **거부된**
    deck-smith 호출(tool use rejected)도 거짓통과했다(실세션 실증). 게이트가 tool_use↔
    tool_result 페어링으로 '실제 완료'를 보는지(거부·무응답 불인정) 마커로 못 박는다 —
    이 마커(is_error·tool_result·거부 문자열)가 사라지면 naive grep으로의 회귀."""
    sj = _read(".claude/settings.json")
    arm = _read(".claude/hooks/pptx_arm.py")
    gate = _read(".claude/hooks/pptx_render_gate.sh")
    matcher_ok = "Write|Edit|Bash|PowerShell" in sj
    arm_ok = "pptx_signal" in arm and ".pptx.tsv" in arm
    # 존재 마커 + 페어링(완료 실행) 판정 마커. 후자가 naive grep과의 차이를 강제한다.
    exists_ok = "subagent_type" in gate and "deck-smith" in gate and ".pptx.tsv" in gate
    pairing_ok = (
        "tool_result" in gate and "is_error" in gate and "The user doesn't want to proceed" in gate
    )
    gate_ok = exists_ok and pairing_ok
    check(
        "⑤ 산출물→덱 에이전트 게이트 (pptx 저장 감지 → deck-smith 완료 실행 경유, 거부 불인정)",
        matcher_ok and arm_ok and gate_ok,
        f"matcher={matcher_ok} arm감지={arm_ok} 게이트검사(존재={exists_ok}·페어링={pairing_ok})",
    )


# ── ⑥ 규칙 택소노미 양방향 (design-rules 강제 태그 ↔ 검사 실체 — 유령·고아·부채 래칫) ──
PROSE_DEBT_CAP = (
    4  # [산문·하강대기: 태그 수(§3·§7·§8·E-2). 증감은 P1.5 백로그 경유 + 사용자 승인 전용.
)


def test_rules_taxonomy():
    """2026-07-25: design-rules(570줄)가 목차는 있는데 **강제 상태 축**이 없어 "산문인데
    강제라고 착각"이 반복됐다("규칙이 있어도 러너에 안 물리면 없는 규칙"). 각 § 머리의
    `> **강제:**` 태그를 기계로 양방향 검증한다:
    ①전 § 태그 존재(영구 census) ②[기계: check_x]는 generic_checks에, [기계·골든: check_x]는
    audit_golden에 실재·배선(유령 태그 차단 — 골든 전용 검사를 전 덱 규칙에 달면 유령 강제)
    ③audit.py 전 검사가 문서에 등장(고아 검사 차단 — 새 검사 추가 시 문서 태그 강제 래칫)
    ④[산문·하강대기:] 총수 ≤ CAP(산문 부채 래칫 — 넘으려면 이 상수를 같은 커밋에서 올려야
    해서 부채 증가가 항상 가시화된다). 태그 전이의 정당성은 기계가 판정 못 함(사용자 승인 몫)."""
    doc = _read(".claude/skills/pptx-build/references/design-rules.md")
    audit_src = _read(".claude/skills/pptx-build/scripts/goldenfab/audit.py")
    golden_src = _read(".claude/skills/pptx-build/scripts/audit_golden.py")
    probs = []
    # ① 전 §(##~####) 헤딩 밑 3줄 내 강제 태그 줄
    lines = doc.split("\n")
    for i, ln in enumerate(lines):
        if re.match(r"^#{2,4} ", ln):
            win = lines[i + 1 : i + 4]
            if not any(w.startswith("> **강제:**") or w.startswith("> (구조") for w in win):
                probs.append(f"태그 없음: {ln.lstrip('# ')[:24]}")
    # ② 태그의 check_*가 맞는 러너에 실재·배선
    defined = set(re.findall(r"^def (check_[a-z_]+)", audit_src, re.M))
    gen_body = (
        audit_src.split("def generic_checks", 1)[1].split("\ndef ", 1)[0]
        if "def generic_checks" in audit_src
        else ""
    )
    for m in re.finditer(r"\[(기계·골든|기계)[^\]]*\]", doc):
        cls = m.group(1)
        for tok in re.findall(r"check_[a-z_]+", m.group(0)):
            if tok not in defined:
                probs.append(f"유령 태그: {tok}(audit.py에 없음)")
            elif cls == "기계" and f"{tok}(" not in gen_body:
                probs.append(f"오배선: [기계] {tok}가 generic_checks에 없음")
            elif cls == "기계·골든" and f"{tok}(" not in golden_src:
                probs.append(f"오배선: [기계·골든] {tok}가 audit_golden에 없음")
    # ③ 고아 검사(정의됐는데 문서 미등장) — 새 검사는 반드시 규칙 태그를 동반
    orphans = sorted(c for c in defined if c not in doc)
    if orphans:
        probs.append(f"고아 검사(문서 태그 없음): {orphans}")
    # ④ 산문 부채 래칫
    n_debt = doc.count("[산문·하강대기:")
    if n_debt > PROSE_DEBT_CAP:
        probs.append(
            f"산문 부채 증가: {n_debt} > 상한 {PROSE_DEBT_CAP}(백로그 경유+사용자 승인 필요)"
        )
    check(
        "⑥ 규칙 택소노미 양방향 (전 § 태그 · 유령/고아 0 · 부채 래칫)",
        not probs,
        "; ".join(probs)
        if probs
        else f"§ 태그 전수 · 검사 {len(defined)}종 전부 문서 등장 · 부채 {n_debt}/{PROSE_DEBT_CAP}",
    )


# ── 통합 A: 밀도 지표 단일 출처 (audit.py 하나, 경쟁 임계 없음) ──────────────────
def test_density_metric_single_source():
    pf = _read(".claude/skills/pptx-build/scripts/preflight_dense.py")
    uses_audit = "A.generic_checks" in pf or "A.check_density" in pf  # 정본 재사용
    no_rival = (
        "def dense_band" not in pf
        and "def check_accent_fills" not in pf  # 자체 accent 중복 제거됨(→audit dense 인지)
        and "def check_offslide" not in pf  # 자체 경계 중복 제거됨(→audit dense 인지)
        and "660" not in pf
    )
    check(
        "통합A 밀도/검사 지표 단일 출처 (preflight가 audit.generic_checks 재사용, 중복 없음)",
        uses_audit and no_rival,
        f"audit 재사용={uses_audit} 중복없음={no_rival}",
    )


# ── 통합 B: 결정표 단일 출처 (layout-matching, 세 어휘 crosswalk) ────────────────
def test_decision_table_single_source():
    lm = _read(".claude/skills/deck-compose/references/layout-matching.md")
    vs = _read(".claude/skills/pptx-visuals/references/visual-selection.md")
    ac = _read(".claude/skills/pptx-visuals/references/archetype-catalog.md")
    lm_ok = "물성" in lm and "FT Visual Vocabulary" in lm and "검색 플라이휠" in lm
    cross = "layout-matching" in vs and "layout-matching" in ac
    check(
        "통합B 결정표 단일 출처 (layout-matching + 세 어휘 crosswalk)",
        lm_ok and cross,
        f"통합표={lm_ok} crosswalk={cross}",
    )


# ── 가드: content_contract가 전 registry 레이아웃에서 해석되나 (①이 계약 안 깼나) ─
def test_content_contract_resolves():
    from goldenfab.content_contract import required_keys
    from goldenfab.registry import LAYOUTS

    fails = []
    for name, fn in LAYOUTS.items():
        try:
            required_keys(name, fn)
        except Exception as e:  # noqa: BLE001
            fails.append(f"{name}: {type(e).__name__}")
    check(
        "가드 content_contract 전 레이아웃 해석 (① 재지정이 계약 안 깸)",
        not fails,
        f"해석 실패 {fails}" if fails else f"{len(LAYOUTS)}개 전부 해석",
    )


# ── 통합 C: 카탈로그 철칙 — 모든 L-ID는 렌더러가 실존 (hollow 카탈로그 방지) ─────
def _renderer_keys():
    """전 렌더러 키 집합을 소스에서 수집(matplotlib·pptx 무거운 import 회피)."""
    vsrc = _read(".claude/skills/pptx-visuals/scripts/visuals.py")
    msrc = _read(".claude/skills/pptx-visuals/scripts/mpl_exhibits.py")
    bsrc = _read(".claude/skills/pptx-build/scripts/build_pptx.py")
    # 차트 = CHART_TYPES 키 + MPL_TYPES + panels(chart 렌더러 분기, build_pptx.chart)
    ct = re.search(r"CHART_TYPES\s*=\s*\{(.*?)\}", vsrc, re.S)
    chart = set(re.findall(r'"([a-z0-9_]+)"\s*:', ct.group(1))) if ct else set()
    mt = re.search(r"MPL_TYPES\s*=\s*\{(.*?)\}", msrc, re.S)
    chart |= set(re.findall(r'"([a-z0-9_]+)"', mt.group(1))) if mt else set()
    chart.add("panels")
    # 다이어그램 = add_diagram의 layout 분기(== 및 in (...))
    diagram = set(re.findall(r'layout == "([a-z0-9_]+)"', vsrc))
    for grp in re.findall(r"layout in \(([^)]*)\)", vsrc):
        diagram |= set(re.findall(r'"([a-z0-9_]+)"', grp))
    # 표계열 = build_pptx.RENDERERS 키
    rb = re.search(r"RENDERERS\s*=\s*\{(.*?)\}", bsrc, re.S)
    table = set(re.findall(r'"([a-z0-9_]+)"\s*:', rb.group(1))) if rb else set()
    return chart, diagram, table


def test_catalog_renderers_exist():
    # 정본 = recommend_archetypes.LID_KEY(L01~L38 → family:kind), 카탈로그 철칙의 기계 미러.
    # 철칙(archetype-catalog:8): "모든 행은 렌더러가 실존" — 렌더러 없는 L-ID 등재 금지.
    sys.path.insert(0, str(ROOT / ".claude/skills/pptx-visuals/scripts"))
    from recommend_archetypes import LID_KEY

    chart, diagram, table = _renderer_keys()
    missing = []
    for lid, key in LID_KEY.items():
        fam, _, kind = key.partition(":")
        if fam == "chart":
            ok = kind in chart
        elif fam == "diagram":
            ok = kind in diagram
        else:  # fam ∈ {table,metrics,two_column,bullets}; "table:matrix"는 base table
            ok = fam in table
        if not ok:
            missing.append(f"{lid}={key}")
    check(
        "통합C 카탈로그 렌더러 실존 (LID_KEY 전부 렌더러 해석 — hollow 행 금지)",
        not missing,
        f"미해석(렌더러 없는 L-ID) {missing}"
        if missing
        else f"L-ID {len(LID_KEY)}종 전부 실존 (chart {len(chart)}·diagram {len(diagram)}·표 {len(table)})",
    )


# ── 통합 D: brand-kit 밖 hex 하드코딩 래칫 (색 단일 출처 — design-rules §2) ────────
def test_no_new_hardcoded_hex():
    # 규율 "신규 hex 금지 — 색은 brand-kit 단일 출처"를 래칫으로 기계화. 기준선 = 승인된 예외
    # (mpl 데이터 베이스 회색 + 흰색 그라디언트 끝점 2건). 초과 = 새 하드코딩 유입 → FAIL.
    base = {
        ".claude/skills/pptx-visuals/scripts/visuals.py": 0,
        ".claude/skills/pptx-build/scripts/build_pptx.py": 0,
        ".claude/skills/pptx-visuals/scripts/mpl_exhibits.py": 2,
    }
    pat = re.compile(r"#[0-9a-fA-F]{6}\b|0x[0-9a-fA-F]{6}\b")
    over = []
    for rel, b in base.items():
        n = len(pat.findall(_read(rel)))
        if n > b:
            over.append(f"{rel.rsplit('/', 1)[-1]}: {n} > 기준선 {b}")
    check(
        "통합D brand-kit 밖 hex 하드코딩 래칫 (초과=신규 하드코딩)",
        not over,
        f"초과 {over}"
        if over
        else "visuals 0 · build_pptx 0 · mpl 2(승인 예외) — 신규 하드코딩 없음",
    )


def main():
    for t in (
        test_dense_ready,
        test_gate_mechanical,
        test_diversity_golden_aware,
        test_adhoc_card_gate_wired,
        test_text_collision_gate_wired,
        test_artifact_agent_gate_wired,
        test_rules_taxonomy,
        test_density_metric_single_source,
        test_decision_table_single_source,
        test_content_contract_resolves,
        test_catalog_renderers_exist,
        test_no_new_hardcoded_hex,
    ):
        try:
            t()
        except Exception as e:  # noqa: BLE001
            check(t.__name__, False, f"예외 — {type(e).__name__}: {e}")

    print("═" * 70)
    fails = 0
    for name, ok, detail in RESULTS:
        mark = "OK  " if ok else "✘FAIL"
        print(f"{mark} {name}\n       {detail}")
        fails += not ok
    print("═" * 70)
    print(
        "WIRING PASS — 전 체인 연결됨" if not fails else f"WIRING FAIL {fails}건 — 끊긴 연결 있음"
    )
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
