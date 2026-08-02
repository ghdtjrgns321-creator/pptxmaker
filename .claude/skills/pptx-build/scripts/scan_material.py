#!/usr/bin/env python
"""재료 스캐너 — FINAL-REPORT(md)에서 **셀 수 있는 것만** 뽑는다 (2026-07-27 신설).

    uv run python .claude/skills/pptx-build/scripts/scan_material.py FINAL-REPORT/*.md
    uv run python .claude/skills/pptx-build/scripts/scan_material.py FINAL-REPORT/ --json > 재료.json

## 왜 이 파일이 있나

장별 형태 후보를 낼 때(pptmaker 파이프라인 ②) 재료를 세어야 형식이 정해진다. 그런데 그 "센 값"을
지금까지 **사람이 손으로 적어 넣었다**(2026-07-27 실전 시험). 손으로 적으면 두 가지가 무너진다:

1. 사람이 그림을 먼저 떠올리고 거기 맞는 숫자를 적을 수 있다 — 판정을 앞으로 숨긴 것이다.
2. 재현되지 않는다. 같은 보고서를 두 번 읽으면 두 번 다른 값이 나온다.

그래서 md 원문에서 **기계가 센다.** 이 스캐너는 그림을 모르고 부품 이름도 모른다 — 표가 몇
행인지, 목록이 몇 개인지, 화살표가 어디서 어디로 가는지만 본다.

## 무엇을 세나 (전부 md 문법)

    표          `| a | b |` 연속 줄       → 행 수 = n · 열 3개 이상이면 부가열 있음
    번호 목록   `1. ` `2. `               → 항목 수 = n · 순서 있음
    글머리 목록 `- ` `* `                 → 항목 수 = n
    아스키 도해 코드펜스 안 `→ ▼ ─▶`      → 화살표 수 = 링크 수

한 절(`##`)에 여러 구조가 있으면 **가장 큰 것**을 그 절의 골격으로 본다. 나머지는 `_others`에
남겨 사람이 확인할 수 있게 한다.

## 이 스캐너가 하지 않는 것 (정직)

- **어느 절이 슬라이드가 되는지 고르지 않는다.** 그건 ① GRILL 인터뷰와 컴포저의 일이다.
- **의미를 읽지 않는다.** "이 표가 대비인가 나열인가"는 판정이고, 여기서는 열 수만 센다.
- 그래서 출력에 `_evidence`(무엇을 보고 세었나)를 붙인다 — 사람이 뒤집어 볼 수 있어야 한다.
"""

import argparse
import json
import re
import sys
from pathlib import Path

H2 = re.compile(r"^##+\s+(.*)$")
TABLE_ROW = re.compile(r"^\s*\|(.+)\|\s*$")
TABLE_SEP = re.compile(r"^\s*\|[\s:|-]+\|\s*$")
NUM_ITEM = re.compile(r"^\s*(\d+)[.)]\s+\S")
BUL_ITEM = re.compile(r"^\s*[-*]\s+\S")
FENCE = re.compile(r"^\s*```")
ARROW = re.compile(r"[→▶]|->|─▶")
# 순서를 뜻하는 열 이름 — 표가 단계표인지 보는 유일한 단서(의미 해석이 아니라 머리글 대조)
ORDER_HEADS = ("단계", "순서", "순번", "step", "phase", "시점", "연도", "일자")


def _cells(line):
    return [c.strip() for c in TABLE_ROW.match(line).group(1).split("|")]


def _scan_table(lines, i):
    """표 한 덩이 — (헤더, 데이터행들, 다음 인덱스). 구분선(---)은 데이터가 아니다."""
    rows = []
    while i < len(lines) and TABLE_ROW.match(lines[i]):
        if not TABLE_SEP.match(lines[i]):
            rows.append(_cells(lines[i]))
        i += 1
    return (rows[0], rows[1:], i) if rows else (None, [], i)


def _scan_section(title, body):
    """한 절에서 가장 큰 구조를 골라 **센 값**을 낸다. 그림은 모른다."""
    lines = body.splitlines()
    found, i, in_fence, fence_arrows = [], 0, False, 0
    num_items, bul_items = [], []
    while i < len(lines):
        ln = lines[i]
        if FENCE.match(ln):
            in_fence = not in_fence
            i += 1
            continue
        if in_fence:
            fence_arrows += len(ARROW.findall(ln))
            i += 1
            continue
        if TABLE_ROW.match(ln):
            head, rows, i = _scan_table(lines, i)
            if rows:
                found.append(("표", head, rows))
            continue
        if NUM_ITEM.match(ln):
            num_items.append(ln.strip())
        elif BUL_ITEM.match(ln):
            bul_items.append(ln.strip())
        i += 1

    cands = []
    for kind, head, rows in found:
        cands.append({"kind": kind, "n": len(rows), "head": head, "rows": rows})
    if num_items:
        cands.append({"kind": "번호목록", "n": len(num_items), "items": num_items})
    if bul_items:
        cands.append({"kind": "목록", "n": len(bul_items), "items": bul_items})
    if not cands:
        return None
    best = max(cands, key=lambda c: c["n"])

    out = {"title": title, "_구조": best["kind"], "items": list(range(best["n"]))}
    ev = {"셌다": f"{best['kind']} {best['n']}개"}

    if best["kind"] == "표":
        head = best["head"] or []
        ncol = len(head)
        out["extra"] = ncol >= 3  # 항목 + 설명 2열을 넘으면 부가 열이 있는 것
        ev["열"] = ncol
        if any(any(k in (h or "").lower() for k in ORDER_HEADS) for h in head):
            out["ordered"] = True
            ev["순서"] = "머리글에 단계/순서 계열 낱말"
        # 셀 안 화살표 — **대부분의 행이 쌍일 때만** 구조로 본다.
        # 소수 행에만 있으면 그건 산문 표기(예: 설명 문장 안의 "A → B")이지 이 표의 골격이
        # 아니다. 전부 링크로 세면 15행 표가 1:N 방사로 잡히는 오판이 난다(2026-07-27 실측).
        rows_with = sum(1 for r in best["rows"] if ARROW.search(" ".join(r)))
        if rows_with >= max(2, int(len(best["rows"]) * 0.8)):
            out["sets"] = 2  # 행마다 좌·우가 있다 = 두 묶음의 대응
            out["links"] = [[f"a{k}", f"b{k}"] for k in range(rows_with)]
            ev["링크"] = f"{rows_with}/{len(best['rows'])}행이 쌍 — 골격으로 인정"
        elif rows_with:
            ev["링크"] = f"{rows_with}/{len(best['rows'])}행만 화살표 — 산문 표기로 보고 무시"
    elif best["kind"] == "번호목록":
        out["ordered"] = True
        out["extra"] = any("—" in t or " · " in t for t in best["items"])
        ev["순서"] = "번호 목록"
    else:
        out["extra"] = any("—" in t or " · " in t or "**" in t for t in best["items"])

    # 코드펜스 안 아스키 도해 — 표·목록이 이미 골격이면 그건 보조 그림이다.
    # 골격이 없을 때만 흐름으로 인정한다(화살표가 있으면 순서가 메시지인 경우가 대부분).
    if fence_arrows and best["kind"] not in ("표", "번호목록"):
        ev["도해 화살표"] = f"{fence_arrows}개 — 골격이 목록뿐이라 흐름으로 인정"
        out["ordered"] = True
    # ── 못 정한 축은 물어볼 목록으로 ────────────────────────────────────────────
    # 스캐너가 셀 수 없는 것이 있다(2026-07-27 손·자동 대조 실측):
    #   · 열이 여럿인 표가 **나열인가 좌우 대비인가** — 열 이름의 뜻을 읽어야 안다
    #   · 시간순인데 머리글에 단계/연도 낱말이 없는 표
    #   · 산문 문장 속 수치("코드 소비자 9종") — 표도 목록도 아니라 안 잡힌다
    # 이건 파서를 더 똑똑하게 만들 문제가 아니라 **① GRILL이 물어야 할 것**이다.
    # 사용자는 도해 용어를 몰라도 "이건 옛날 방식과 비교예요"는 말할 수 있다.
    ask = []
    if best["kind"] == "표" and len(best.get("head") or []) >= 3 and not out.get("links"):
        ask.append("이 표가 그냥 나열인가, 두 쪽을 맞대어 비교하는 것인가")
    if best["kind"] == "표" and not out.get("ordered"):
        ask.append("행에 순서(시간·단계)가 있나")
    if ask:
        out["_ask"] = ask

    out["_evidence"] = ev
    if len(cands) > 1:
        out["_others"] = [f"{c['kind']} {c['n']}" for c in cands if c is not best]
    return out


def scan(paths):
    files = []
    for p in paths:
        q = Path(p)
        files += sorted(q.glob("*.md")) if q.is_dir() else [q]
    out = []
    for f in files:
        text = f.read_text(encoding="utf-8")
        parts = re.split(r"^(##+\s+.*)$", text, flags=re.M)
        for k in range(1, len(parts), 2):
            title = H2.match(parts[k]).group(1).strip()
            sec = _scan_section(title, parts[k + 1])
            if sec and sec["items"]:
                sec["_출처"] = f"{f.name} · {title}"
                out.append(sec)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="+")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--min", type=int, default=2, help="이 개수 미만인 절은 뺀다")
    a = ap.parse_args(argv)

    secs = [s for s in scan(a.paths) if len(s["items"]) >= a.min]
    if a.json:
        for i, s in enumerate(secs, 1):
            s["no"] = i
        print(json.dumps({"slides": secs}, ensure_ascii=False, indent=2))
        return 0
    print(f"{'절':44s} {'구조':8s} {'n':>3s}  근거")
    for s in secs:
        ev = " · ".join(f"{k}={v}" for k, v in s["_evidence"].items())
        print(f"{s['_출처'][:44]:44s} {s['_구조']:8s} {len(s['items']):3d}  {ev}")
    print(f"\n{len(secs)}개 절 — 어느 것이 슬라이드가 되는지는 ①GRILL·컴포저가 정한다")
    ask = [s for s in secs if s.get("_ask")]
    if ask:
        print(f"\n■ 축을 못 정한 절 {len(ask)}개 — ①GRILL이 물어볼 목록")
        for s in ask[:8]:
            print(f"  {s['title'][:36]:36s} {' / '.join(s['_ask'])}")
        if len(ask) > 8:
            print(f"  ... 외 {len(ask) - 8}개")
    return 0


if __name__ == "__main__":
    sys.exit(main())
