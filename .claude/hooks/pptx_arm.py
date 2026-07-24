# -*- coding: utf-8 -*-
"""PostToolUse(Write|Edit|Bash|PowerShell) — pptmaker 세션 arm 마커 기록.

무장 2종 (판정은 하지 않는다 — pptx_render_gate.sh(Stop)가 마커를 읽어 검사한다):

1) 소스 arm({sid}.tsv, 2026-07-21): dense/build/audit 소스를 Write/Edit
   → 게이트가 preflight green + 렌더 신선도를 검사.
2) 산출물 arm({sid}.pptx.tsv, 2026-07-24): Bash/PowerShell 명령이 .pptx를
   저장/빌드/복사/렌더 → 게이트가 "덱 에이전트(deck-smith 등) 경유"를 검사.
   배경(2026-07-24 전수 감사): 기존 arm이 "goldenfab 소스 편집"(작업대)만 봐서,
   스크래치에서 pptx를 만들면(카드 목업 세션 실증) 하네스 전체가 잠들었다.
   게이트를 경로가 아니라 **산출물**에 붙이는 확장.
   한계(정직): 감지는 명령 문자열 휴리스틱 — `python script.py`처럼 저장 경로가
   스크립트 안에만 있으면 그 호출 자체는 미탐(렌더·복사 단계에서 잡힌다).

fail-open: 어떤 예외도 조용히 통과한다 — 이 훅은 절대 편집/실행을 막지 않는다(기록 전용).
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from pathlib import Path

# .pptx가 명령에 등장 + 산출 동사 동반 → 산출물 arm. 렌더(render_deck)도 포함 —
# 렌더는 덱 루프의 일부이며, 눈검증 없는 빌드-only 세션은 어차피 사용자에게 안 보인다.
PPTX_VERBS = ("save", "build_pptx", "render_deck", "copy-item", "cp ", "move-item", "mv ")


def armable(p: Path) -> bool:
    """이 경로가 dense 워크플로 소스인가(소스 arm)."""
    if p.suffix != ".py":
        return False
    s = p.as_posix()
    if "goldenfab" in s and p.stem.endswith("_dense"):
        return True
    return p.name in {"dense.py", "build_pptx.py", "audit.py"} and "pptx-build" in s


def pptx_signal(cmd: str) -> str | None:
    """명령이 .pptx 산출물을 만들거나 옮기는가 — 맞으면 마커 스니펫 반환."""
    low = cmd.lower()
    if ".pptx" not in low:
        return None
    if not any(v in low for v in PPTX_VERBS):
        return None
    m = re.search(r"[^\s\"';|&()]*\.pptx", cmd, re.I)
    return (m.group(0) if m else ".pptx")[:160]


def _write_marker(armf: Path, key: str, stamp: int) -> None:
    """경로/스니펫당 1줄, 최신 스탬프로 재기록. CRLF 방지(newline=\\n) — 게이트 bash read 보호."""
    lines = []
    if armf.exists():
        for ln in armf.read_text(encoding="utf-8").splitlines():
            parts = ln.split("\t", 1)
            if len(parts) == 2 and parts[1] != key:
                lines.append(ln)
    lines.append(f"{stamp}\t{key}")
    armf.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")


def main() -> None:
    try:
        payload = json.load(sys.stdin)
    except Exception:
        return
    sid = payload.get("session_id") or "nosession"
    tool = payload.get("tool_name") or ""
    ti = payload.get("tool_input") or {}
    state = Path(os.path.expanduser("~")) / ".claude" / "state" / "pptx_arm"

    if tool in ("Bash", "PowerShell"):
        sig = pptx_signal(ti.get("command") or "")
        if not sig:
            return
        state.mkdir(parents=True, exist_ok=True)
        _write_marker(state / f"{sid}.pptx.tsv", sig, time.time_ns())
        return

    # Write/Edit(및 file_path를 가진 그 외) — 소스 arm(기존 동작)
    fp = ti.get("file_path") or ""
    if not fp:
        return
    p = Path(fp)
    if not armable(p):
        return
    try:
        mtime = p.stat().st_mtime_ns
    except OSError:
        return
    state.mkdir(parents=True, exist_ok=True)
    _write_marker(state / f"{sid}.tsv", p.as_posix(), mtime)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        pass  # fail-open — 기록 실패가 편집/실행을 막지 않는다
