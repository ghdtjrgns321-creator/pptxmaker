"""골든덱 빌드 — `goldenfab.reference.build_reference()`를 pptx로 떨어뜨린다.

**골든은 하나다**(2026-07-26 dense 승격). 그전에는 둘이었다 — 운영·열람용
`golden-deck-operating.pptx`(dense)를 뽑는 `build_operating.py`가 따로 있었고, registry와
회귀 하네스는 sparse variant를 지켰다. 그래서 도해마다 sparse 1벌 + dense 1벌이 존재했고
(가로 좌표 동일·세로만 다름), 장 하나 고치려면 두 벌을 고쳐야 했다.
승격으로 registry가 dense를 가리키므로 `build_reference()` 하나가 곧 골든이다.

실행: uv run python golden/build_golden.py → golden/golden-deck.pptx
회귀: compare_golden(스냅샷) · audit_golden(장별 기하) · preflight_dense
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / ".claude/skills/pptx-build/scripts"))

from goldenfab.reference import build_reference  # noqa: E402


def main():
    prs = build_reference()
    out = ROOT / "golden" / "golden-deck.pptx"
    prs.save(str(out))
    print(f"saved {out} ({len(prs.slides._sldIdLst)} slides)")
    print("dense 미승인 1건(sparse 렌더러 유지): exec_graph(S6) — 기준작 후보 s06_mid")


if __name__ == "__main__":
    main()
