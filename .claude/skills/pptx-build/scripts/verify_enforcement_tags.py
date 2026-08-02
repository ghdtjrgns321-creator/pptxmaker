#!/usr/bin/env python
"""강제 태그 대조 — design-rules.md가 "기계가 잡는다"고 말한 것이 실제로 잡는지 확인한다.

    uv run python .claude/skills/pptx-build/scripts/verify_enforcement_tags.py

## 왜 이 파일이 있나 (2026-08-02 신설)

2026-07-29 대정리로 부품·러너·훅이 아카이브됐는데 `design-rules.md`의 `> **강제:**` 태그는
"기계가 잡는다"고 말한 채 남았다 — **7개 절, 나흘간**. 그 사이 덱을 만드는 쪽은 없는 방어를
믿었다. 게다가 사람이 눈으로 훑으면 빠뜨린다(첫 대조에서 §6 `check_air` 1건 누락, grep 재대조로
검출). deck-smith 루프 0단계가 매번 읽는 것이 이 태그라서, 태그가 거짓이면 배치 판단이 통째로
틀어진다.

## test_wiring(아카이브)의 실패를 반복하지 않는 법

옛 배선 테스트 1,048줄은 **검사 대상을 하드코딩**해서, 모듈이 아카이브되자 유령을 검사하는
코드가 됐고 통째로 버려졌다. 이 스크립트는 분모를 매번 `design-rules.md`에서 뽑는다 —
규칙이 사라지면 분모도 같이 준다. 살아있는 파일만 보므로 스스로 썩지 않는다.

## 무엇을 판정하고 무엇을 안 하나 (정직)

판정한다 — `[기계]`·`[부품]`·`[훅]`·`[기계·골든]` 태그가 가리키는 실물의 **존재와 호출**:
검사 함수는 `audit.py`에 정의가 있고 **호출자가 있나**(정의만 있고 호출자 0이면 안 도는 검사다) ·
부품 모듈·속성이 `goldenfab/`에 있나 · 훅이 settings에 물렸나.

판정하지 않는다 — **그 규칙이 옳은지**. `[눈]`·`[산문·하강대기]` 태그는 사람이 진다고 적힌
것이라 대조 대상이 아니다. 즉 이 게이트가 막는 것은 **"기계가 잡는다"는 거짓말 하나**뿐이다.

이미 `~~취소선~~`으로 죽었다고 표시된 것은 주장이 아니므로 건너뛴다.
"""

import json
import re
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
RULES = SCRIPTS.parent / "references" / "design-rules.md"
GOLDENFAB = SCRIPTS / "goldenfab"
AUDIT = GOLDENFAB / "audit.py"
SETTINGS = [
    SCRIPTS.parents[2] / "settings.json",  # .claude/settings.json (프로젝트)
    Path.home() / ".claude" / "settings.json",  # 전역
]

# 실물을 주장하는 태그 종류만 대조한다. [눈]·[산문·하강대기]·[절차]·[장부]는 사람 몫.
MACHINE_KINDS = ("기계", "부품", "훅", "기계·골든")

TAG = re.compile(r"^> \*\*강제:\*\*(.*)$")
GROUP = re.compile(r"\[([^\]:]+):([^\]]*)\]")
STRUCK = re.compile(r"~~.*?~~", re.S)
CHECK = re.compile(r"\bcheck_[a-z_]+\b")
DOTTED = re.compile(r"\b([a-z][a-z0-9_]*)\.([a-z_][a-z0-9_]*)\b")
MODULE_PY = re.compile(r"\b([a-z][a-z0-9_]*)\.py\b")
HOOK = re.compile(
    r"\b([a-z][a-z0-9_]*)\.(?:sh|py)\b|\b(pptx_render_gate|pptx_arm|completion_gate|post_write_check|guard_bash)\b"
)


def wired_hooks() -> set[str]:
    """settings.json(프로젝트·전역)에 실제로 물린 훅 이름."""
    names: set[str] = set()
    for path in SETTINGS:
        if not path.exists():
            continue
        raw = path.read_text(encoding="utf-8")
        try:
            json.loads(raw)  # 형식만 확인 — 배선 표기는 파일마다 달라 문자열로 훑는다
        except json.JSONDecodeError:
            print(f"경고: {path} 파싱 실패 — 훅 대조를 건너뛴다", file=sys.stderr)
            continue
        names |= {m.group(0).rsplit(".", 1)[0] for m in re.finditer(r"[a-z_]+\.(?:sh|py)", raw)}
    return names


def audit_symbols() -> tuple[set[str], set[str]]:
    """audit.py의 (정의된 검사, 호출자가 있는 검사)."""
    src = AUDIT.read_text(encoding="utf-8")
    defined = set(re.findall(r"^def (check_[a-z_]+)", src, re.M))
    called = {n for n in defined if len(re.findall(rf"\b{n}\(", src)) > 1}
    return defined, called


def module_has(mod: str, attr: str) -> bool:
    path = GOLDENFAB / f"{mod}.py"
    if not path.exists():
        return False
    src = path.read_text(encoding="utf-8")
    return bool(
        re.search(rf"^(?:def |class |{re.escape(attr)}\s*=)", src, re.M)
        and re.search(rf"\b{re.escape(attr)}\b", src)
    )


def main() -> int:
    defined, called = audit_symbols()
    audit_src = AUDIT.read_text(encoding="utf-8")
    hooks = wired_hooks()
    dead: list[tuple[int, str, str]] = []
    unknown: list[tuple[int, str]] = []
    claims = 0

    for lineno, line in enumerate(RULES.read_text(encoding="utf-8").splitlines(), 1):
        m = TAG.match(line)
        if not m:
            continue
        body = STRUCK.sub("", m.group(1))  # 이미 죽었다고 표시된 것은 주장이 아니다
        for kind, content in GROUP.findall(body):
            if kind.strip() not in MACHINE_KINDS:
                continue
            claims += 1
            for name in CHECK.findall(content):
                if name not in defined:
                    dead.append((lineno, name, "audit.py에 정의 없음"))
                elif name not in called:
                    dead.append((lineno, name, "정의만 있고 호출자 0 — 안 도는 검사다"))
            for mod, attr in DOTTED.findall(content):
                if attr in {"py", "md", "json", "yaml", "sh", "pptx"}:
                    continue  # `grid.py`는 모듈.속성이 아니다 — 아래 MODULE_PY가 본다
                if not (GOLDENFAB / f"{mod}.py").exists():
                    dead.append(
                        (lineno, f"{mod}.{attr}", f"goldenfab/{mod}.py 없음 — 아카이브됐다")
                    )
                elif not module_has(mod, attr):
                    dead.append((lineno, f"{mod}.{attr}", f"{mod}.py에 {attr} 없음"))
            for mod in MODULE_PY.findall(content):
                if not (GOLDENFAB / f"{mod}.py").exists():
                    dead.append((lineno, f"{mod}.py", "goldenfab에 없음"))
            if kind.strip() == "훅":
                for a, b in HOOK.findall(content):
                    name = a or b
                    if name and name not in hooks:
                        dead.append((lineno, name, "settings.json 어디에도 안 물림"))
            # 기계를 주장하는데 대조 규칙이 없는 상수는 조용히 통과시키지 않는다.
            # [부품]·[훅]까지 넓히면 KOREAN·VERIFY 같은 무해한 낱말이 섞여 소음이 된다
            # — 오탐 게이트는 무시당한다(이 저장소 실적).
            if kind.strip().startswith("기계"):
                rest = CHECK.sub("", DOTTED.sub("", MODULE_PY.sub("", content)))
                for tok in re.findall(r"\b[A-Z][A-Z_]{3,}\b", rest):
                    if tok not in audit_src:
                        unknown.append((lineno, tok))

    print(f"분모: `> **강제:**` 줄에서 실물 주장 태그 {claims}개 ({RULES.name})")
    for lineno, name, why in dead:
        print(f"  FAIL  {RULES.name}:{lineno}  {name} — {why}")
    for lineno, tok in unknown:
        print(f"  확인  {RULES.name}:{lineno}  {tok} — 대조 규칙 없음, 눈으로 본다")
    if dead:
        print(f"\n죽은 강제 {len(dead)}건. 태그를 [눈]으로 강등하거나 실물을 되살린다.")
        print("(강등할 때는 ~~취소선~~ + 소멸 근거를 같은 줄에 남긴다 — 되살아남 방지)")
        return 1
    print("죽은 강제 0건 — 태그가 실물과 맞는다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
