"""계약 검사 — 제목·정형 장 문구가 **재료에서 따온 것인지** 대조한다.

## 왜 (2026-08-03)

`deck-outline-grill`에 "장 제목은 재료 헤딩에서 뽑는다 · 부·표지 문구도 재료에서 따온다"를
적었지만, 이 저장소에서 산문 지시는 예외 없이 무시됐다. 같은 날 실측: 본문 12장 중 8장의
제목이 조판 단계 창작이었고, 부 이름(「왜 전수인가」류)은 5개 전부 창작이었다.

그래서 계약을 재료와 기계가 대조한다. 통과 기준은 낱말 겹침이다 — 헤딩을 그대로 베끼라는
뜻이 아니라(「3-2. 기술 설명 - 룰 기반 검증」 → 「룰 기반 검증 29종」은 정상), **재료에 없는
말을 지어내지 말라**는 뜻이다.

    uv run python scripts/check_outline.py <01.5_outline.md> <재료.md>
"""

import re
import sys
from pathlib import Path

OVERLAP_MIN = 0.5  # 제목 어절의 이 비율 이상이 어떤 헤딩에 있어야 한다


def headings(md: Path) -> list[str]:
    return [ln.lstrip("#").strip() for ln in md.read_text(encoding="utf-8").splitlines()
            if re.match(r"^#{1,4} ", ln)]


def words(s: str) -> set:
    return {w for w in re.split(r"[\s·,/()\-—「」『』]+", re.sub(r"[0-9.]+", "", s)) if len(w) >= 2}


def rows(md: Path, col: str) -> list[tuple[str, str]]:
    """`col`로 시작하는 열을 가진 **표 하나**에서 (키, 값)을 뽑는다.

    표 경계를 지킨다 — 머리글을 만날 때마다 열 위치를 다시 잡고, 표가 끝나면 버린다.
    안 그러면 다음 표의 같은 자리 셀까지 긁어 온다(2026-08-03에 실제로 그랬다).
    """
    out, idx, first = [], None, True
    for ln in md.read_text(encoding="utf-8").splitlines():
        if not ln.strip().startswith("|"):
            idx, first = None, True
            continue
        c = [x.strip() for x in ln.strip().strip("|").split("|")]
        if set("".join(c)) <= set("-: "):
            continue
        head = [i for i, v in enumerate(c) if v.startswith(col)]
        if first:
            first = False
            if head:
                idx = head[0]
                continue
        if idx is not None and idx < len(c) and c[idx] not in ("", "—"):
            out.append((c[0], c[idx]))
    return out


def main() -> int:
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    outline, material = Path(sys.argv[1]), Path(sys.argv[2])
    heads = headings(material)
    pool = [words(h) for h in heads]
    bad, seen = [], 0

    for label, col, kind in [("장 제목", "장 제목", "장"), ("부 제목", "제목", "부")]:
        for key, title in rows(outline, col):
            if title in ("표지", "목차", "간지"):
                continue
            seen += 1
            w = words(title)
            if not w:
                continue
            def hit(p):
                # 조사가 붙어도 같은 말로 본다 — 「의사결정과」/「의사결정」
                return sum(any(a.startswith(b) or b.startswith(a) for b in p) for a in w) / len(w)

            best = max((hit(p) for p in pool), default=0.0)
            if best < OVERLAP_MIN:
                bad.append(f"{label} {key} 「{title}」 — 재료 헤딩과 겹침 {best:.0%} · 지어낸 말이다")

    print(f"{outline.name} ↔ {material.name} — 제목 {seen}개 · 재료 밖 {len(bad)}건")
    for b in bad:
        print("  ★", b)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
