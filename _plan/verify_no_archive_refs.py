"""살아있는 하네스가 아카이브된 것을 가리키는지 전수 검사 — 이 저장소가 반복해 실패한 지점.

"옮겼는데 참조가 남은 것"이 사고 #10(죽은 경로 22건)의 원인이었다. 두 갈래로 본다:

    [1부] 경로 참조 — 문서·코드에 적힌 파일 경로가 실존하나
    [2부] 개념 이름 — 죽은 게이트·에이전트·파이프라인을 **지시**하고 있나

2부가 필요한 이유(2026-07-29 실증): 1부만 돌렸을 때 deck-outline-grill이 "deck-compose 진입
전에 반드시 사용"을 4곳에서 지시하고 있었는데 **전부 통과했다.** `deck-compose`는 경로가 아니라
개념 이름이라 1부의 정규식에 안 걸렸다. 다음 세션은 그 지시를 읽고 없는 파이프라인을 찾아간다.

    uv run python _plan/verify_no_archive_refs.py
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LIVE = [".claude/skills", ".claude/agents", "CLAUDE.md", "_lab"]

# 경로처럼 생긴 것만 — 일반어 오탐을 원천 차단한다
PATH_PAT = re.compile(
    r"(?:\.claude/[\w\-./]+|goldenfab/[\w\-./]+|scripts/[\w\-./]+|references/[\w\-.]+"
    r"|assets/[\w\-.]+|golden/[\w\-./가-힣]+|_lab/[\w\-.]+|docs/[\w\-.가-힣]+)"
)
EXT = {".py", ".md", ".json", ".yaml", ".ps1", ".sh", ".pptx", ".png"}

# 역사 서술 표지 — 그 대목은 "지금 없다"를 말하는 중이므로 검사 제외
HISTORY = (
    "~~",
    "아카이브",
    "_archive/",
    "폐기",
    "삭제",
    "옛 ",
    "사라졌",
    "무효",
    "되살리지",
    "소멸",
    "2026-07-29",
    "이전 하네스",
    "없어진",
    "대상이 0",
    "당시 기록",
)

# 아카이브된 개념 이름 — 경로가 아니라서 1부가 못 잡는다
DEAD_NAMES = [
    "deck-compose",
    "deck-composer",
    "pptx-builder",
    "consistency-qa",
    "deck-spec",
    "build_pptx",
    "registry",
    "content_contract",
    "test_wiring",
    "audit_deck",
    "compare_golden",
    "preflight_dense",
    "audit_golden",
    "audit_pptx",
    "layout-matching",
    "select.py",
    "frames",
    "make_mockups",
    "recommend_archetypes",
    "check_candidates",
]

# 검사했고 정당하다고 판정한 자리 — 창 필터에도 안 걸리는 역사 서술만 등재한다.
# 새로 추가할 땐 반드시 사유를 적는다(사유 없는 항목 금지 — 그건 은폐다).
WHITELIST = {
    # "왜 글자수가 주 지표인가"의 출처 주석. preflight_dense가 이 파일에 흡수된 이력이다.
    r".claude\skills\pptx-build\scripts\goldenfab\audit.py:679",
}

SKILL_BASES = []
for sk in (ROOT / ".claude/skills").iterdir():
    if (sk / "SKILL.md").exists():
        SKILL_BASES += [sk, sk / "scripts"]


def live_files(targets):
    for target in targets:
        t = ROOT / target
        files = [t] if t.is_file() else [f for f in t.rglob("*") if f.is_file()]
        for f in files:
            if "__pycache__" not in str(f):
                yield f


# ── 1부: 경로 실존 ───────────────────────────────────────────────
missing, checked = [], 0
for f in live_files(LIVE):
    if f.suffix not in {".md", ".py", ".yaml", ".json", ".ps1"}:
        continue
    try:
        text = f.read_text(encoding="utf-8")
    except (UnicodeDecodeError, OSError):
        continue
    for i, line in enumerate(text.splitlines(), 1):
        if any(h in line for h in HISTORY):
            continue
        for m in PATH_PAT.finditer(line):
            p = m.group().rstrip(".,)`'\"")
            if Path(p).suffix not in EXT:
                continue  # 디렉토리 언급은 설명이 섞여 판정 불가
            if "*" in p or "NN" in p or "모듈" in p:
                continue  # 와일드카드·자리표시자
            checked += 1
            # ROOT · 전 스킬 루트와 그 scripts/ · 파일의 부모 — 문서는 경로를 짧게 쓴다
            if not any((b / p).exists() for b in [ROOT, f.parent, *SKILL_BASES]):
                missing.append((f.relative_to(ROOT), i, p))

# ── 2부: 죽은 개념 이름 ──────────────────────────────────────────
# _lab은 덱 **내용**(슬라이드에 실릴 문구)이라 제외 — 옛 구조를 소개하는 문장이 들어 있을 수
# 있고 그건 하네스 결함이 아니라 덱 내용 문제다. 경로 검사(1부)는 그대로 받는다.
concept = []
for f in live_files([t for t in LIVE if t != "_lab"]):
    # .json은 데이터(재료 스캔 산출 샘플 등) — 지시가 아니다
    if f.suffix not in {".md", ".py", ".yaml", ".ps1"}:
        continue
    try:
        lines = f.read_text(encoding="utf-8").splitlines()
    except (UnicodeDecodeError, OSError):
        continue
    for i, line in enumerate(lines, 1):
        # 문맥 창 — 표지가 이웃 줄에 있다("이전 하네스는 …" 다음 줄의 목록 항목 등)
        window = "\n".join(lines[max(0, i - 4) : i + 3])
        if any(h in window for h in HISTORY):
            continue
        if f"{f.relative_to(ROOT)}:{i}" in WHITELIST:
            continue
        for name in DEAD_NAMES:
            # 단어 경계 — frames_min(변수명)이 frames로 잡히는 오탐 차단
            if re.search(rf"(?<![\w-]){re.escape(name)}(?![\w])", line):
                concept.append((f.relative_to(ROOT), i, name, line.strip()[:70]))

print(f"[1부] 경로 참조 {checked}건 검사 · 실존하지 않는 것 {len(missing)}건")
for f, i, p in missing:
    print(f"  {f}:{i}  ->  {p}")
print(f"[2부] 죽은 개념 이름 {len(DEAD_NAMES)}종 검사 · 살아있는 지시 {len(concept)}건")
for f, i, name, line in concept:
    print(f"  {f}:{i}  [{name}]  {line}")
sys.exit(1 if (missing or concept) else 0)
