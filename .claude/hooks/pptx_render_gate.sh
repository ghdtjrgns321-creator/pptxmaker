#!/bin/bash
# Stop 훅 — pptmaker 렌더 게이트 (자가무장, 2026-07-21)
#
# 강제하는 규율: "렌더 눈검증·preflight green 없이 dense 완료 선언 금지"(design-rules §8,
#   consistency-qa §3). 지금까지 prose라서 무시됐다 — 여기서 기계로 못 하게 만든다.
#
#   무장 = 이번 세션에 goldenfab dense/build/audit 소스를 편집(pptx_arm.py가 기록).
#   차단 = ① preflight red(generic_checks: 빈상자·넘침·재탕·밀도·accent)  OR
#          ② 마지막 편집 이후 렌더 PNG(render_deck.ps1의 sNN.png)가 없음/스테일.
#   통과 = 무장 없음 / 둘 다 green / 백스톱(N=3 무진행) / fail-open(오류 시).
#
# self-regulation(completion_gate와 동일 — 이 프로젝트가 "상시 적색 게이트는 무시당한다"를
#   test_wiring.py:37에 학습): 같은 SIG로 N=3회 무진행 재차단되면 로그 남기고 자동통과+음소거.
set -uo pipefail
N=3
INPUT=$(cat 2>/dev/null)

SID=""
if [[ "$INPUT" =~ \"session_id\"[[:space:]]*:[[:space:]]*\"([^\"]*)\" ]]; then
  SID="${BASH_REMATCH[1]}"
fi
[[ -n "$SID" ]] || exit 0

STATE="$HOME/.claude/state/pptx_arm"
ARM="$STATE/$SID.tsv"
COUNTER="$STATE/.count.$SID"
MUTED="$STATE/.muted.$SID"
OVLOG="$STATE/.override.log"
[[ -f "$ARM" ]] || exit 0

# Windows 경로(C:/...) → Git Bash 경로(/c/...) 변환 — stat·find가 드라이브레터를 못 읽는 것 방지.
to_unix() {
  local p="${1//\\//}"
  if [[ "$p" =~ ^([A-Za-z]):/(.*)$ ]]; then
    printf '/%s/%s' "$(printf '%s' "${BASH_REMATCH[1]}" | tr 'A-Z' 'a-z')" "${BASH_REMATCH[2]}"
  else
    printf '%s' "$p"
  fi
}

# 존재하는 armed 소스만 훑어 최신 mtime 파일·dense 스템·프로젝트 루트를 구한다.
NEWEST_SRC=""
NEWEST_T=0
DENSE_STEMS=""
PROJ=""
while IFS=$'\t' read -r _ns path; do
  path="${path%$'\r'}"  # CRLF 방어(Windows에서 arm 파일이 \r\n일 수 있음)
  [[ -n "$path" ]] || continue
  u=$(to_unix "$path")
  [[ -f "$u" ]] || continue
  cur=$(stat -c %Y "$u" 2>/dev/null || echo 0)
  if [[ "$cur" -ge "$NEWEST_T" ]]; then NEWEST_T="$cur"; NEWEST_SRC="$u"; fi
  base="${u##*/}"; stem="${base%.py}"
  [[ "$stem" == *_dense ]] && DENSE_STEMS+="$stem "
  [[ -z "$PROJ" && "$u" == */.claude/* ]] && PROJ="${u%%/.claude/*}"
done < "$ARM"

# 편집 소스가 하나도 안 남았으면(되돌림·삭제) 무장 해제하고 통과.
[[ -n "$NEWEST_SRC" && -n "$PROJ" ]] || { rm -f "$ARM" "$COUNTER" "$MUTED" 2>/dev/null; exit 0; }

# 백스톱용 서명(armed 경로·mtime 집합).
if command -v sha1sum >/dev/null 2>&1; then
  SIG=$(sort "$ARM" | sha1sum | cut -d' ' -f1)
else
  SIG=$(wc -c < "$ARM" | tr -d ' ')
fi
MSIG=""
if [[ -f "$MUTED" ]]; then
  read -r MSIG < "$MUTED" 2>/dev/null || MSIG=""
  [[ "$MSIG" != "$SIG" ]] && { rm -f "$MUTED" 2>/dev/null; MSIG=""; }
fi

emit_block() {
  python - "$1" <<'PY' 2>/dev/null || exit 0
import json, sys
print(json.dumps({"decision": "block", "reason": sys.argv[1]}))
PY
}
emit_ctx() {
  python - "$1" <<'PY' 2>/dev/null || exit 0
import json, sys
print(json.dumps({"hookSpecificOutput": {"hookEventName": "Stop", "additionalContext": sys.argv[1]}}))
PY
}

backstop_or_block() {
  local reason="$1"
  [[ "$MSIG" == "$SIG" ]] && exit 0
  local PREV_SIG="" PREV_CNT=0 CNT
  [[ -f "$COUNTER" ]] && { read -r PREV_SIG PREV_CNT < "$COUNTER" 2>/dev/null || true; }
  [[ "$PREV_CNT" =~ ^[0-9]+$ ]] || PREV_CNT=0
  if [[ "$SIG" == "$PREV_SIG" ]]; then CNT=$((PREV_CNT + 1)); else CNT=1; fi
  printf '%s %s\n' "$SIG" "$CNT" > "$COUNTER" 2>/dev/null
  if [[ "$CNT" -ge "$N" ]]; then
    { printf '[render-gate backstop] auto-pass:\n%s\n' "$reason"; } >> "$OVLOG" 2>/dev/null
    rm -f "$COUNTER" 2>/dev/null
    printf '%s\n' "$SIG" > "$MUTED" 2>/dev/null
    emit_ctx "렌더 게이트 백스톱: ${N}회 무진행 자동통과. 렌더 눈검증 미완을 사용자에게 명시 보고할 것."
    exit 0
  fi
  emit_block "$reason (무진행 ${CNT}/${N}, 도달 시 자동통과)"
  exit 0
}

FAILS=""

# --- 검사 ①: preflight green (dense 스템 지정, 없으면 통과작 자기검증) ---
TO=""
command -v timeout >/dev/null 2>&1 && TO="timeout 120"
pf_out=$(cd "$PROJ" && $TO uv run python .claude/skills/pptx-build/scripts/preflight_dense.py $DENSE_STEMS 2>&1)
pf_rc=$?
if [[ "$pf_rc" -ne 0 ]]; then
  FAILS+="✗ preflight red (rc=${pf_rc}): $(printf '%s' "$pf_out" | tail -c 300)
"
fi

# --- 검사 ②: 렌더 신선도 (편집 이후 sNN.png 존재) ---
FRESH=$(find "$PROJ" -path '*/.venv' -prune -o -path '*/.git' -prune -o \
  -name 's[0-9][0-9].png' -newer "$NEWEST_SRC" -print 2>/dev/null | head -1)
if [[ -z "$FRESH" ]]; then
  FAILS+="✗ 렌더 없음/스테일: '${NEWEST_SRC##*/}' 편집 이후 sNN.png 없음. render_deck.ps1로 렌더→눈검증.
"
fi

if [[ -n "$FAILS" ]]; then
  backstop_or_block "종료 차단(렌더 게이트). dense 소스를 고쳤는데 아래가 안 끝났다:
${FAILS}green + 렌더 후 종료하라. 의도적 중간 양보면 백스톱까지 대기(무진행 누적)."
fi

# 둘 다 green — 무장 해제.
rm -f "$ARM" "$COUNTER" "$MUTED" 2>/dev/null
exit 0
