#!/usr/bin/env python
"""부품·틀 선택 러너 — 재료를 세어 장별로 고른다 (2026-07-27).

    uv run python .claude/skills/pptx-build/scripts/pick_parts.py <재료.json> [--json]

규칙 본체는 `goldenfab/select.py`(축 5개 · 2단 규칙), 자리는 `goldenfab/frames.py`.
회귀 검증은 `verify_selection.py` — 골든이 손으로 고른 도해를 규칙이 재현하는지 대조한다.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from goldenfab.select import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
