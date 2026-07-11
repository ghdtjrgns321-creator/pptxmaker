"""goldenfab.registry — 레이아웃 타입 15종 → 확정 빌더 매핑 (docs/user/07-factory-port.md 인벤토리와 1:1).

슬라이스 1(cover·toc·part)은 콘텐츠 dict 파라미터, 슬라이스 2의 10종은 골든 콘텐츠 내장
(closing·screenshot은 c dict 오버라이드 지원) — 심층 파라미터화는 실전 수요 발생 시 pull.
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
