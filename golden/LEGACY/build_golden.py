"""골든 레퍼런스 덱 렌더러 — goldenfab으로 17장(5부)을 조립해 golden-deck.pptx로 저장한다.

실행: uv run python golden/build_golden.py → golden/golden-deck.pptx

## 이 파일이 얇아진 이유 (2026-07-15 단일화)

전에는 여기에 레이아웃 코드(표지·목차·간지)와 헬퍼(content_header·source_line·pct_bar), 콘텐츠
상수가 다 들어 있었고, **똑같은 것이 goldenfab/에도 한 벌 더** 있었다. Phase 2 "공장 역설계
이식"의 비계였는데(docs/user/07) 이식이 끝난 뒤에도 안 걷혀서, 장 하나 고치려면 두 파일을
똑같이 고쳐야 했다. 한쪽만 고치면 회귀 게이트가 깨졌다.

지금은 **goldenfab이 유일 소스**다. 레퍼런스 덱 정의(콘텐츠 상수·장 순서)는
`goldenfab/reference.py`로 이관했고, 이 파일은 그걸 pptx로 떨구는 CLI만 남았다.
회귀 게이트(`compare_golden.py`)는 코드 트리끼리 비교하는 대신 도형 서명 스냅샷
(`assets/golden-snapshot.json`)과 대조한다.

이 pptx는 **눈으로 보려고** 만드는 렌더 산출물이다(.gitignore 대상). 고쳐도 다음 빌드에
덮어써진다 — 레이아웃을 바꾸려면 `goldenfab/`의 코드를 고칠 것.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / ".claude/skills/pptx-build/scripts"))

from goldenfab.reference import build_reference  # noqa: E402


def main():
    prs = build_reference()
    out = Path(__file__).parent / "golden-deck.pptx"
    try:
        prs.save(str(out))
    except PermissionError:  # PowerPoint가 열어둔 경우 — 대체 파일로 저장하고 계속
        out = Path(__file__).parent / "golden-deck.build.pptx"
        prs.save(str(out))
    n = sum(len(s.shapes) for s in prs.slides)
    print(f"saved {out} ({len(prs.slides._sldIdLst)} slides, {n} shapes)")


if __name__ == "__main__":
    main()
