"""dense 장 프리플라이트 게이트 — 사용자에게 보이기 **전에** 기계로 채점한다.

배경(2026-07-20): dense 장을 손으로 만들어 생짜로 사용자에게 보여주고 반려받는 루프가 장마다
반복됐다(허전·탑헤비·재탕). 그 결함은 대부분 기계로 잡히는데 안 돌렸을 뿐. 이 러너가 그 게이트다
— green 되기 전엔 사용자에게 안 보낸다. 실전 파이프라인이 novel 장을 뽑을 때 쓸 게이트와 같은
메커니즘(design-rules P1.5·P4·P5).

**정본 위임(2026-07-20 통합):** 예전엔 이 파일이 자체 밀도·accent·경계 검사를 따로 뒀는데
audit.py와 경쟁하는 중복이었다. 지금은 **`goldenfab.audit.generic_checks(dense=True)` 하나에
전부 위임**한다 — 밀도(글자수)·accent(런 제외)·경계(dense 하한)·채움률(칩 폭)·그림침범(아이콘
제외)이 전부 audit.py에 dense-인지로 있다. 이 파일은 "단일 모듈을 1장짜리 덱으로 빌드해 채점"하는
편의만 남는다. 실전 덱(여러 장)은 audit_deck.py가, 빌드 시 자동은 build_pptx._gate_density가 같은
generic_checks(dense=True)를 쓴다.

사용:
    uv run python .claude/skills/pptx-build/scripts/preflight_dense.py s14_dense [s15_dense ...]
    (인자 없으면 통과작으로 밴드 자기검증 — 통과작이 전부 green인지 확인)
"""

import importlib
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from goldenfab import audit as A  # noqa: E402
from goldenfab.kit import load_kit  # noqa: E402
from goldenfab.reference import new_presentation  # noqa: E402

# 통과한 dense 장(비스크린샷) — 인자 없을 때 자기검증 대상.
PASS_REFS = ["s08_dense", "s09_dense", "s11_dense", "s12_dense"]

# 장별 설정 — 스크린샷 예외, 의도된 반복 노드 허용(§E-2 "같은 노드 재등장이 메시지").
# dup_allow는 통과작의 실측 baseline(래칫) — 이보다 늘면 FAIL.
CFG = {
    "s09_dense": {"dup_allow": 8},  # 카탈로그 — 변동대가가 hier·cp·term 3행 등장 등(§E-2)
    "s10_dense": {"screenshot": True},
    "s11_dense": {"dup_allow": 2},  # 적용 트리 라벨 재등장
    "s12_dense": {"screenshot": True},
    "s14_dense": {"dup_allow": 3},  # 양 패널 동일 질문 + 3동일 표면유사 산탄카드 = 의도
}


def _shapes(mod_name):
    mod = importlib.import_module(f"goldenfab.{mod_name}")
    prs = new_presentation()
    mod.build(prs)
    return list(prs.slides[0].shapes)


def audit_module(mod_name, accent, band):
    cfg = CFG.get(mod_name, {})
    shapes = _shapes(mod_name)
    print(f"\n══ {mod_name}: 도형 {len(shapes)} ══")
    # 정본 audit.py의 dense-인지 전역 검사 재사용(단일 출처). accent 런 제외·경계 dense 하한·
    # 칩 폭·아이콘 제외가 전부 generic_checks(dense=True) 안에 있다.
    res = A.generic_checks(
        shapes,
        accent,
        band=band,
        dense=True,
        screenshot=cfg.get("screenshot", False),
        dup_allow=cfg.get("dup_allow", 0),
    )
    fails = A.report(res)
    return [f"{mod_name}: {x}" for x in fails]


def main():
    accent = str(load_kit()["rgb"]["accent"])
    band = A.density_band()  # 정본 밴드(글자수 하한·골든 스냅샷 파생) — 재개편 시 자동 이동
    print(f"밀도 밴드(정본 audit.py): 글자 ≥{band['chars_min']} · 도형 바닥 {band['shapes_floor']}")
    targets = sys.argv[1:] or PASS_REFS  # 인자 없으면 통과작 자기검증
    all_fails = []
    for m in targets:
        try:
            all_fails += audit_module(m, accent, band)
        except Exception as e:  # noqa: BLE001
            all_fails.append(f"{m}: 예외 — {e}")
    print("\n" + ("PREFLIGHT PASS" if not all_fails else f"PREFLIGHT FAIL {len(all_fails)}건"))
    for x in all_fails:
        print(f"  - {x}")
    return 1 if all_fails else 0


if __name__ == "__main__":
    sys.exit(main())
