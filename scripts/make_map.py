"""pptmaker 지도 생성기 — docs/MAP.md를 코드에서 뽑는다.

손으로 쓴 지도는 거짓말한다. 실측(2026-08-03): `docs/PIPELINE.md`가 스스로를
"지금 어떻게 도나의 단일 출처"라 선언한 채, 2026-07-29에 폐기된 4에이전트
파이프라인·`build_pptx.py`·`part-design`을 6일 동안 설명하고 있었다.

그래서 이 지도에는 **사람이 쓴 문장을 넣지 않는다.** 모든 줄이 파일에서 파생되고,
어긋나면 exit 1로 죽는다. 절차 원문은 SKILL.md가 단일 출처라 그대로 인용만 한다.

실행: uv run python scripts/make_map.py
"""

import ast
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def rel(p: Path) -> str:
    """저장소 기준 상대경로 — 문서에는 슬래시로 적는다(윈도 백슬래시 금지)."""
    return p.resolve().relative_to(ROOT).as_posix()


OUT = ROOT / "docs/MAP.md"

SKILL_MAIN = ROOT / ".claude/skills/pptmaker/SKILL.md"
CODE_DIRS = [ROOT / ".claude/skills/pptx-build/scripts"]
AUDIT = ROOT / ".claude/skills/pptx-build/scripts/goldenfab/audit.py"
DOC_DIRS = [ROOT / "docs", ROOT / ".claude/skills", ROOT / ".claude/agents"]
# 지도가 훑는 문서 확장자 — 코드 참조가 살아있는지 대조할 대상
DOC_GLOB = "*.md"


# ── 1. 절차 ────────────────────────────────────────────────────────────────
def pipeline_block(md: Path) -> str:
    """SKILL.md의 '## 파이프라인' 다음 코드블록을 원문 그대로 가져온다."""
    txt = md.read_text(encoding="utf-8")
    m = re.search(r"^## 파이프라인\s*\n+```\n(.*?)```", txt, re.S | re.M)
    return m.group(1).rstrip() if m else ""


def referenced_tools(block: str) -> list[tuple[str, bool, str]]:
    """절차 안에 등장하는 파일·스킬 이름이 실재하는지 확인한다."""
    names = sorted(set(re.findall(r"[\w\-]+\.(?:py|ps1)|deck-outline-grill", block)))
    out = []
    for n in names:
        hits = [
            p
            for p in ROOT.rglob(f"*{n}*")
            if "_archive" not in p.parts and ".git" not in p.parts and "__pycache__" not in p.parts
        ]
        out.append((n, bool(hits), rel(hits[0]) if hits else "—"))
    return out


# ── 2. 살아있는 코드 ────────────────────────────────────────────────────────
def code_inventory() -> list[tuple[str, int, int, str]]:
    rows = []
    for d in CODE_DIRS:
        for p in sorted(d.rglob("*.py")):
            if "__pycache__" in p.parts:
                continue
            src = p.read_text(encoding="utf-8")
            tree = ast.parse(src)
            fns = [
                n.name
                for n in tree.body
                if isinstance(n, ast.FunctionDef) and not n.name.startswith("_")
            ]
            rows.append((rel(p), len(src.splitlines()), len(fns), ", ".join(fns[:6])))
    return rows


# ── 3. 검사 배선 ────────────────────────────────────────────────────────────
def check_wiring() -> list[tuple[str, str]]:
    src = AUDIT.read_text(encoding="utf-8")
    checks = re.findall(r"^def (check_\w+)", src, re.M)
    generic = src[src.index("def generic_checks") :]
    others = ""
    for d in CODE_DIRS:
        for p in d.rglob("*.py"):
            if p != AUDIT and "__pycache__" not in p.parts:
                others += "\n" + p.read_text(encoding="utf-8")
    rows = []
    for c in checks:
        if c in generic:
            rows.append((c, "generic_checks"))
        elif re.search(rf"\bA?\.?{c}\s*\(", others):
            rows.append((c, "스크립트에서 직접"))
        else:
            rows.append((c, "★ 호출자 없음"))
    return rows


# ── 4. 문서와 죽은 참조 ─────────────────────────────────────────────────────
def live_symbols() -> str:
    out = ""
    for d in CODE_DIRS:
        for p in d.rglob("*.py"):
            if "__pycache__" not in p.parts:
                out += "\n" + p.read_text(encoding="utf-8")
    return out


# 외부·표준 라이브러리 접두사 — 이 저장소가 소유하지 않으므로 대조 대상이 아니다
EXTERNAL = ("pptx.", "inspect.", "XL_", "MSO_", "PP_", "Pt.", "Inches.", "Emu.", "np.", "plt.")


def is_history(p: Path) -> bool:
    """역사 기록은 죽은 코드를 가리켜도 정상이다 — 그때의 사실을 적은 것이므로."""
    name = p.name
    return (
        bool(re.match(r"^\d{4}-\d{2}-\d{2}-", name))
        or name.startswith("handoff-")
        or name == "CHANGELOG.md"
        or "superpowers" in p.parts
    )


def live_modules() -> set[str]:
    return {p.stem for d in CODE_DIRS for p in d.rglob("*.py") if "__pycache__" not in p.parts}


# 같은 줄에 이 낱말이 있으면 "죽었다"고 밝힌 것이다 — 읽는 쪽을 아카이브로 보내지 않으므로
# 경고가 아니다. 사망을 기록한 문장과, 살아있는 척 가리키는 문장을 가르는 유일한 신호.
TOMBSTONE = (
    "아카이브",
    "_archive",
    "옮겼",
    "삭제",
    "소멸",
    "없어진",
    "없앴",
    "물러났",
    "폐기",
    "사라졌",
    "지웠",
    "미배선",
)


def _marked_dead(txt: str, ref: str) -> bool:
    """이 참조가 등장하는 자리가 전부 사망 표기를 달고 있나.

    문장이 줄바꿈으로 갈리므로 앞뒤 한 줄까지 본다 — "생산 코드(...)는 2026-07-29에 /
    아카이브됐다"처럼 표기가 다음 줄에 있는 경우를 오탐으로 잡지 않기 위해서다.
    """
    lines = txt.splitlines()
    spots = [i for i, ln in enumerate(lines) if f"`{ref}`" in ln]
    if not spots:
        return False
    return all(any(k in " ".join(lines[max(0, i - 1) : i + 2]) for k in TOMBSTONE) for i in spots)


def dead_refs(txt: str, symbols: str, modules: set[str]) -> list[str]:
    """문서가 백틱으로 가리키는 심볼·파일 중 살아있는 코드에 없는 것.

    사망 표기(TOMBSTONE)가 붙은 참조는 뺀다 — 이력을 적은 문장까지 잡으면 경고가 소음이 되고,
    소음이 되면 아무도 안 본다. 잡아야 할 것은 **없는 것을 있는 것처럼 가리키는 문장**뿐이다.
    """
    dead = []
    for r in sorted(set(re.findall(r"`([A-Za-z_][\w./-]*(?:\.[A-Za-z_][\w]*)+)`", txt))):
        if r.startswith(EXTERNAL) or _marked_dead(txt, r):
            continue
        if r.endswith((".py", ".ps1")):  # 파일 참조 — 실재 여부로 판정
            stem = Path(r).name
            hits = [
                q for q in ROOT.rglob(stem) if "_archive" not in q.parts and ".git" not in q.parts
            ]
            if not hits:
                dead.append(r)
            continue
        if r.endswith((".md", ".yaml", ".yml", ".pptx", ".png", ".tsv", ".json")):
            continue
        head, sym = r.split(".")[0], r.split(".")[-1]
        if head in modules and sym in modules:  # goldenfab.grid 같은 모듈 참조
            continue
        if sym in modules:
            continue
        if f"def {sym}" in symbols or f'"{sym}"' in symbols or f"{sym} =" in symbols:
            continue
        dead.append(r)
    return dead


def doc_inventory(symbols: str) -> list[tuple[str, int, str, bool, list[str]]]:
    modules = live_modules()
    rows = []
    for d in DOC_DIRS:
        for p in sorted(d.rglob(DOC_GLOB)):
            if "_archive" in p.parts or p.resolve() == OUT.resolve():
                continue  # 자기 자신은 안 센다 — 경고 목록을 다시 참조로 읽는다
            txt = p.read_text(encoding="utf-8")
            hist = is_history(p)
            rows.append(
                (
                    rel(p),
                    len(txt.splitlines()),
                    git_date(p),
                    hist,
                    dead_refs(txt, symbols, modules),
                )
            )
    return rows


def git_date(p: Path) -> str:
    try:
        r = subprocess.run(
            ["git", "log", "-1", "--format=%ad", "--date=short", "--", str(p)],
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=15,
        )
        return r.stdout.strip() or "미커밋"
    except Exception:
        return "?"


# ── 조립 ───────────────────────────────────────────────────────────────────
def main() -> int:
    block = pipeline_block(SKILL_MAIN)
    tools = referenced_tools(block)
    code = code_inventory()
    checks = check_wiring()
    symbols = live_symbols()
    docs = doc_inventory(symbols)

    warn: list[str] = []
    warn += [f"절차가 부르는 `{n}`이(가) 없다" for n, ok, _ in tools if not ok]
    warn += [f"검사 `{c}`를 아무도 부르지 않는다" for c, w in checks if w.startswith("★")]
    for path, _, _, hist, dead in docs:
        if dead and not hist:  # 역사 기록은 죽은 코드를 가리켜도 정상이다
            warn.append(f"`{path}` — 없는 참조 {len(dead)}개: {', '.join(dead)}")

    L = []
    L.append("# pptmaker 지도")
    L.append("")
    L.append("> **자동 생성 — 손으로 고치지 않는다.** `uv run python scripts/make_map.py`")
    L.append("> 모든 줄이 코드·SKILL.md에서 파생된다. 어긋나면 생성기가 exit 1로 죽는다.")
    L.append("")

    L.append("## 1. 덱 하나가 만들어지는 순서")
    L.append("")
    L.append(f"원문 단일 출처: `{rel(SKILL_MAIN)}`")
    L.append("")
    L.append("```")
    L.append(block)
    L.append("```")
    L.append("")
    L.append("절차가 이름을 부르는 것들이 실재하는지:")
    L.append("")
    L.append("| 이름 | 실재 | 위치 |")
    L.append("| --- | --- | --- |")
    for n, ok, where in tools:
        L.append(f"| `{n}` | {'O' if ok else '★ 없음'} | {where} |")
    L.append("")

    L.append("## 2. 살아있는 코드")
    L.append("")
    L.append("| 파일 | 줄 | 공개 함수 | 주요 함수 |")
    L.append("| --- | ---: | ---: | --- |")
    for path, n, nf, names in code:
        L.append(f"| `{path}` | {n} | {nf} | {names} |")
    L.append(f"\n합계 **{sum(r[1] for r in code)}줄**")
    L.append("")

    L.append("## 3. 기계 검사 배선")
    L.append("")
    L.append("| 검사 | 누가 부르나 |")
    L.append("| --- | --- |")
    for c, w in checks:
        L.append(f"| `{c}` | {w} |")
    L.append("")

    L.append("## 4. 문서")
    L.append("")
    L.append(
        "살아있는 문서 = 하네스가 실제로 읽는 것. 역사 기록은 그때의 사실이라 죽은 참조가 정상이다."
    )
    L.append("")
    L.append("| 파일 | 줄 | 최종 커밋 | 종류 | 없는 참조 |")
    L.append("| --- | ---: | --- | --- | --- |")
    for path, n, date, hist, dead in sorted(docs, key=lambda r: (r[3], r[0])):
        kind = "역사" if hist else "**살아있음**"
        shown = (
            ", ".join(f"`{d}`" for d in dead)
            if (dead and not hist)
            else (f"{len(dead)}개" if dead else "—")
        )
        L.append(f"| `{path}` | {n} | {date} | {kind} | {shown} |")
    L.append("")

    L.append("## 5. 경고")
    L.append("")
    if warn:
        L += [f"- {w}" for w in warn]
    else:
        L.append("없음.")
    L.append("")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print(f"생성: {rel(OUT)} — 경고 {len(warn)}건")
    for w in warn:
        print("  ★", w)
    return 1 if warn else 0


if __name__ == "__main__":
    sys.exit(main())
