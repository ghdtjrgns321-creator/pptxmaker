# -*- coding: utf-8 -*-
"""PostToolUse(Write|Edit) — pptmaker dense/build 소스 편집 시 세션 arm 마커 기록.

배경(2026-07-21): "렌더 눈검증·preflight green 없이 dense 완료 선언 금지"를 기계로 강제하려면,
먼저 "이번 세션에 dense 소스를 손댔다"를 세션 스코프로 알아야 한다. git 워킹트리는 세션 스코프가
아니라(구 세션의 미커밋 변경이 상시 적색을 만든다) 여기서 세션별 마커에 기록한다. 판정은 하지
않는다 — pptx_render_gate.sh(Stop)가 이 마커를 읽어 preflight·렌더 신선도를 검사한다.

arm 대상(pptmaker pptx 소스): goldenfab/*_dense.py · dense.py · build_pptx.py · goldenfab/audit.py.
fail-open: 어떤 예외도 조용히 통과한다 — 이 훅은 절대 편집/빌드를 막지 않는다(기록 전용).
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path


def armable(p: Path) -> bool:
    """이 경로가 dense 워크플로 소스인가."""
    if p.suffix != ".py":
        return False
    s = p.as_posix()
    if "goldenfab" in s and p.stem.endswith("_dense"):
        return True
    return p.name in {"dense.py", "build_pptx.py", "audit.py"} and "pptx-build" in s


def main() -> None:
    try:
        payload = json.load(sys.stdin)
    except Exception:
        return
    fp = (payload.get("tool_input") or {}).get("file_path") or ""
    if not fp:
        return
    p = Path(fp)
    if not armable(p):
        return
    try:
        mtime = p.stat().st_mtime_ns
    except OSError:
        return

    sid = payload.get("session_id") or "nosession"
    state = Path(os.path.expanduser("~")) / ".claude" / "state" / "pptx_arm"
    state.mkdir(parents=True, exist_ok=True)
    armf = state / f"{sid}.tsv"

    # 같은 경로 중복 제거 후 최신 mtime으로 재기록(경로당 1줄).
    key = p.as_posix()
    lines = []
    if armf.exists():
        for ln in armf.read_text(encoding="utf-8").splitlines():
            parts = ln.split("\t", 1)
            if len(parts) == 2 and parts[1] != key:
                lines.append(ln)
    lines.append(f"{mtime}\t{key}")
    # newline="\n" — Windows 기본 텍스트모드의 CRLF 변환 방지(게이트 bash read가 \r에 걸린다).
    armf.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        pass  # fail-open — 기록 실패가 편집을 막지 않는다
