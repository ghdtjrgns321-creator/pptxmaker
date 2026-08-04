"""계약(`01.5_outline.md`)에서 조판이 **그대로 써야 하는 것**을 읽는다.

조판이 지어내는 것을 막는 길은 둘이다 — 게이트로 뒤에서 잡거나, 애초에 계약에서 읽어오거나.
부(部) 이름은 후자로 한다. 부 내비게이션은 본문 전 장의 상단에 5줄로 깔리므로, 손으로 옮겨
적으면 **장 수 × 부 수**만큼 틀릴 자리가 생긴다. 2026-08-03 실측: 부 이름 5개가 전부 조판
단계 창작이었고(간지 「Ⅱ 검증할 데이터」 vs 계약 「Ⅱ 합성 데이터 생성」) 아무도 못 봤다.

읽는 곳은 두 표뿐이다.

- `### 부(간지)`  → `parts()`   [(로마자, 제목, 리드), ...]
- `## 장 목록`    → `part_of()` {장 번호: 로마자} — 정형 장(표지·목차·간지)은 뺀다
"""

import re
from pathlib import Path

FORMAL = ("표지", "목차", "간지")


def tables(path):
    """마크다운 표를 (헤더, 행들)로 끊어 낸다.

    표가 끝나면 헤더도 버린다 — 한 문서에 표가 여러 개고, 열 위치가 표마다 다르다.
    (구판 파서가 이걸 안 해서 장 목록의 `부` 열을 표지 표에도 적용했다.)
    """
    head, rows = None, []
    for ln in Path(path).read_text(encoding="utf-8").splitlines():
        s = ln.strip()
        if not s.startswith("|"):
            if head:
                yield head, rows
            head, rows = None, []
            continue
        cells = [c.strip() for c in s.strip("|").split("|")]
        if set("".join(cells)) <= set("-: "):
            continue
        if head is None:
            head = cells
        else:
            rows.append(cells)
    if head:
        yield head, rows


def _col(head, *names):
    for i, c in enumerate(head):
        if c.startswith(names):
            return i
    return None


def parts(path) -> list[tuple[str, str, str]]:
    """`### 부(간지)` 표 → [(로마자, 제목, 리드), ...]. 없으면 빈 목록."""
    for head, rows in tables(path):
        if head[0] != "부":
            continue
        ti, li = _col(head, "제목"), _col(head, "리드")
        if ti is None:
            continue
        out = []
        for r in rows:
            if not r[0] or r[0] == "—" or ti >= len(r):
                continue
            out.append((r[0], r[ti], r[li] if li is not None and li < len(r) else ""))
        if out:
            return out
    return []


def part_of(path) -> dict[int, str]:
    """`## 장 목록` 표 → {장 번호: 로마자}. **정형 장은 뺀다** — 내비가 붙는 장만 남긴다."""
    for head, rows in tables(path):
        if head[0] != "#":
            continue
        pi, ti = _col(head, "부"), _col(head, "장 제목", "제목")
        if pi is None:
            continue
        out = {}
        for r in rows:
            m = re.match(r"^S?(\d+)$", r[0])
            if not m or pi >= len(r) or r[pi] in ("", "—"):
                continue
            if ti is not None and ti < len(r) and r[ti].strip().startswith(FORMAL):
                continue
            out[int(m.group(1))] = r[pi]
        if out:
            return out
    return {}
