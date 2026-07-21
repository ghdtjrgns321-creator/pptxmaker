"""goldenfab.registry — 레이아웃 타입 15종 → 확정 빌더 매핑 (docs/user/07-factory-port.md 인벤토리와 1:1).

**15종 전부 content dict 파라미터를 받는다**(`fn(prs, c)`, 실측 2026-07-15: signature 15/15).

골든 기본값은 **회귀 게이트(compare_golden) 전용 기준점**이다 — 골든이 안 변했는지 비교하는
잣대이지 콘텐츠 원천이 아니다. 공장 문(build_pptx._render_golden)은 content가 골든 기본값을
전부 덮지 않으면 빌드를 중단한다(`content_contract.assert_content`). 골든 글은 밖으로 못 나간다.
"""

from ._variant_k import variant_k as problem_grid
from .layouts import cover, part, toc
from .s06_variants import variant_c as exec_graph
from .s08_variants import variant_c as tech_evidence
from .s09_variants import variant_a as tech_tree
from .s10_screenshot import variant_a as screenshot
from .s11_variants import variant_d as tech_mechanism
from .s12_variants import variant_b as tech_capture
from .s14_variants import variant_c as ab_simulation
from .s15_variants import variant_c as validation
from .s17_variants import variant_c as mirror_matrix
from .s18_variants import variant_b as boundary
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
