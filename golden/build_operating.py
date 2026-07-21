"""운영 골든덱 조립 — 프레임(reference) + dense 본문(_pilot_s06 개편분)을 SLIDE_ORDER대로 조립.

레지스트리 승격(registry→dense)은 회귀 하네스(compare_golden·audit_golden)를 dense 좌표로
재조정해야 해서 보류 중(CLAUDE.md 이력). 이 스크립트는 그와 별개로 **운영·열람용 골든덱**을
dense로 뽑는다. 백업 golden-deck.pptx(구 sparse)는 건드리지 않는다.

실행: uv run python golden/build_operating.py → golden/golden-deck-operating.pptx
미완(미래 과제): exec_graph(S6) dense 없음 → sparse fallback · S14(ab_simulation) 교체 예정 ·
                카드 팔레트 확장(큰 실루엣 아키타입 2~3종). docs/troubleshooting 참조.
"""

import importlib
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / ".claude/skills/pptx-build/scripts"))

from goldenfab.reference import (  # noqa: E402
    SLIDE_ORDER,
    cover_content,
    new_presentation,
    part_content,
    toc_content,
)
from goldenfab.registry import LAYOUTS  # noqa: E402

# 본문 키 → dense 모듈(있으면 우선). exec_graph(S6)는 아직 dense 없음 → sparse fallback.
DENSE = {
    "problem_grid": "s04_dense",
    "tech_evidence": "s08_dense",
    "tech_tree": "s09_dense",
    "screenshot": "s10_dense",
    "tech_mechanism": "s11_dense",
    "tech_capture": "s12_dense",
    "ab_simulation": "s14_dense",
    "validation": "s15_dense",
    "boundary": "s16_dense",
}


def build(prs=None):
    prs = prs or new_presentation()
    part_no = 0
    sparse = []
    for key, name in SLIDE_ORDER:
        if key == "cover":
            LAYOUTS[key](prs, cover_content())
        elif key == "toc":
            LAYOUTS[key](prs, toc_content())
        elif key == "part":
            part_no += 1
            LAYOUTS[key](prs, part_content(part_no))
        elif key == "closing":  # 프레임(간지·클로징) — 골든 기본값
            LAYOUTS[key](prs, None)
        elif key in DENSE:
            importlib.import_module(f"goldenfab.{DENSE[key]}").build(prs)
        else:  # dense 미완 본문(exec_graph) — sparse 골든 기본값으로 채운다
            LAYOUTS[key](prs, None)
            sparse.append(f"{name}({key})")
    return prs, sparse


def main():
    prs, sparse = build()
    out = ROOT / "golden" / "golden-deck-operating.pptx"
    prs.save(str(out))
    print(f"saved {out} ({len(prs.slides._sldIdLst)} slides)")
    if sparse:
        print("dense 미완(sparse fallback · 미래 과제):")
        for s in sparse:
            print(f"  - {s}")


if __name__ == "__main__":
    main()
