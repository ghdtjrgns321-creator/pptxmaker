"""goldenfab.registry — 레이아웃 타입 15종 → 확정 빌더 매핑 (docs/user/07-factory-port.md 인벤토리와 1:1).

**15종 전부 content dict 파라미터를 받는다**(`fn(prs, c)`, 실측 2026-07-15: signature 15/15).

골든 기본값은 **회귀 게이트(compare_golden) 전용 기준점**이다 — 골든이 안 변했는지 비교하는
잣대이지 콘텐츠 원천이 아니다. 공장 문(build_pptx._render_golden)은 content가 골든 기본값을
전부 덮지 않으면 빌드를 중단한다(`content_contract.assert_content`). 골든 글은 밖으로 못 나간다.
"""

# dense 승격 완료(2026-07-26, 사용자 확정). 골든은 **운영 골든(dense) 하나**다.
#
# 그전까지 골든이 둘로 갈라져 있었다 — 운영·열람용은 `golden-deck-operating.pptx`(dense)인데
# registry·회귀 하네스는 sparse variant를 지켰다. 그래서 도해마다 sparse 1벌 + dense 1벌이
# 존재했고(가로 좌표는 동일, 세로만 다름 — s08 실측: 2.45/4.75/8.20 공통, 세로 3.88 vs 1.30),
# 장 하나 고치려면 두 벌을 고쳐야 했다. 2026-07-15에 없앤 `golden/`↔`goldenfab/` 이중화와
# **같은 병이 다른 축에서 재발**한 것이다.
#
# 승격으로 sparse 기준선은 폐기했다. 회귀 스냅샷(`golden-snapshot.json`)·장별 기하 오라클
# (`audit_golden.SPECS`)은 dense 좌표로 재수립했다.
# exec_graph(S6)만 dense 기준작이 미승인이라 sparse 렌더러를 유지한다(부채 1건, 아래 주석).
from ._variant_k import variant_k as problem_grid_sparse  # noqa: F401  (S4 sparse — 참조 보존)
from .layouts import cover, part, toc
from .s04_dense import build as problem_grid
from .s06_variants import variant_c as exec_graph  # dense 미승인 — s06_mid가 기준작 후보
from .s08_dense import build as tech_evidence
from .s09_dense import build as tech_tree
from .s10_dense import build as screenshot
from .s11_dense import build as tech_mechanism
from .s12_dense import build as tech_capture
from .s14_dense import build as ab_simulation
from .s15_dense import build as validation
from .s16_dense import build as boundary
from .s17_variants import variant_c as mirror_matrix  # 골든 덱 제외 · 실전 창고용 타입
from .s21_closing import variant_a as closing

LAYOUTS = {
    "cover": cover,
    "toc": toc,
    "part": part,
    "problem_grid": problem_grid,
    "exec_graph": exec_graph,
    "tech_evidence": tech_evidence,
    "tech_tree": tech_tree,
    "screenshot": screenshot,
    "tech_mechanism": tech_mechanism,
    "tech_capture": tech_capture,
    "ab_simulation": ab_simulation,
    "validation": validation,
    "mirror_matrix": mirror_matrix,
    "boundary": boundary,
    "closing": closing,
}
