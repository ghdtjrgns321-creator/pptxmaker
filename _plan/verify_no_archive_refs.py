"""살아있는 하네스에 적힌 파일 경로가 실존하는지 전수 검사 — 이 저장소가 반복해 실패한 지점.

"옮겼는데 참조가 남은 것"이 사고 #10(죽은 경로 22건)의 원인이었다. 산문 언급이 아니라
**경로처럼 생긴 문자열**만 뽑아 실존을 확인한다(일반어 오탐 없음).

    uv run python _plan/verify_no_archive_refs.py
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LIVE = [".claude/skills", ".claude/agents", "CLAUDE.md", "_lab"]

# 경로처럼 생긴 것: 슬래시를 포함하고 알려진 확장자로 끝나거나 알려진 디렉토리로 시작
PATH_PAT = re.compile(
    r"(?:\.claude/[\w\-./]+|goldenfab/[\w\-./]+|scripts/[\w\-./]+|references/[\w\-.]+"
    r"|assets/[\w\-.]+|golden/[\w\-./가-힣]+|_lab/[\w\-.]+|docs/[\w\-.가-힣]+)"
)
EXT = {".py", ".md", ".json", ".yaml", ".ps1", ".sh", ".pptx", ".png"}
# 역사 서술 표지 — 그 줄의 경로는 "지금 없다"를 말하는 중이므로 검사 제외
HISTORY = ("~~", "아카이브", "_archive/", "폐기", "옛 ", "사라졌", "무효", "되살리지")

SKILL_BASES = []
for sk in (ROOT / ".claude/skills").iterdir():
    if (sk / "SKILL.md").exists():
        SKILL_BASES += [sk, sk / "scripts"]

missing, checked = [], 0
for target in LIVE:
    t = ROOT / target
    files = [t] if t.is_file() else [f for f in t.rglob("*") if f.is_file()]
    for f in files:
        if f.suffix not in {".md", ".py", ".yaml", ".json", ".ps1"} or "__pycache__" in str(f):
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
                    continue  # 디렉토리 언급은 건너뜀(와일드카드·설명이 섞임)
                if "*" in p or "NN" in p or "모듈" in p:
                    continue  # 와일드카드·자리표시자
                checked += 1
                # ROOT 기준 · 파일이 속한 스킬 루트 기준 · 파일의 부모 기준 순으로 해석
                # ROOT · 전 스킬 루트와 그 scripts/ · 파일의 부모 — 문서는 경로를 짧게 쓴다
                bases = [ROOT, f.parent, *SKILL_BASES]
                if not any((b / p).exists() for b in bases):
                    missing.append((f.relative_to(ROOT), i, p))

print(f"경로 참조 {checked}건 검사 · 실존하지 않는 것 {len(missing)}건")
for f, i, p in missing:
    print(f"  {f}:{i}  ->  {p}")
sys.exit(1 if missing else 0)
