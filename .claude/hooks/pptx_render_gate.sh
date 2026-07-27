#!/bin/bash
# Stop 훅 — pptmaker 덱 게이트 (자가무장 2026-07-21 · 산출물→에이전트 확장 2026-07-24)
#
# 무장 2종 (pptx_arm.py가 세션 스코프로 기록):
#   ① 소스 arm({sid}.tsv): goldenfab dense/build/audit 소스 편집.
#      차단 = preflight red(generic_checks) OR 렌더 스테일(편집 후 sNN.png 없음). (기존)
#   ② 산출물 arm({sid}.pptx.tsv): 이 세션의 명령이 .pptx를 저장/빌드/복사/렌더.
#      차단 = 대화 기록(transcript)에 덱 에이전트(deck-smith·pptx-builder·deck-composer·
#      consistency-qa)의 **완료 실행**이 없음 — tool_use↔tool_result 페어링으로 판정(거부·
#      무응답 호출은 불인정). "덱 작업은 에이전트 루프 경유"(CLAUDE.md 정체)를 기계로 강제.
#      메인 즉흥 덱 작업(스크래치 우회)이 하네스 전체를 잠재우던 구멍을 막는다(2026-07-24 전수
#      감사 — 입구만 기계화). 페어링 교체(2026-07-24): 거부된 호출도 tool_use JSON은 남아
#      옛 grep(문자열 존재)이 거짓통과하던 것을 실증하고 실제 완료 판정으로 전환.
#   통과 = 무장 없음 / 전부 green / 백스톱(N=3 무진행 자동통과) / fail-open(오류 시).
#
# self-regulation(completion_gate와 동일 — "상시 적색 게이트는 무시당한다", test_wiring.py:37):
#   같은 SIG로 N=3회 무진행 재차단되면 로그 남기고 자동통과+음소거.
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
PPTX="$STATE/$SID.pptx.tsv"
COUNTER="$STATE/.count.$SID"
MUTED="$STATE/.muted.$SID"
OVLOG="$STATE/.override.log"
[[ -f "$ARM" || -f "$PPTX" ]] || exit 0

# Windows 경로(C:/...) → Git Bash 경로(/c/...) 변환 — stat·find가 드라이브레터를 못 읽는 것 방지.
to_unix() {
  local p="${1//\\//}"
  if [[ "$p" =~ ^([A-Za-z]):/(.*)$ ]]; then
    printf '/%s/%s' "$(printf '%s' "${BASH_REMATCH[1]}" | tr 'A-Z' 'a-z')" "${BASH_REMATCH[2]}"
  else
    printf '%s' "$p"
  fi
}

# ── 소스 arm 스캔 — 존재하는 armed 소스의 최신 mtime·dense 스템·프로젝트 루트 ──
NEWEST_SRC=""
NEWEST_T=0
DENSE_STEMS=""
PROJ=""
if [[ -f "$ARM" ]]; then
  while IFS=$'\t' read -r _ns path; do
    path="${path%$'\r'}"  # CRLF 방어
    [[ -n "$path" ]] || continue
    u=$(to_unix "$path")
    [[ -f "$u" ]] || continue
    cur=$(stat -c %Y "$u" 2>/dev/null || echo 0)
    if [[ "$cur" -ge "$NEWEST_T" ]]; then NEWEST_T="$cur"; NEWEST_SRC="$u"; fi
    base="${u##*/}"; stem="${base%.py}"
    [[ "$stem" == *_dense ]] && DENSE_STEMS+="$stem "
    [[ -z "$PROJ" && "$u" == */.claude/* ]] && PROJ="${u%%/.claude/*}"
  done < "$ARM"
  # 편집 소스가 하나도 안 남았으면(되돌림·삭제) 소스 무장만 해제. 산출물 무장은 별개.
  [[ -n "$NEWEST_SRC" && -n "$PROJ" ]] || rm -f "$ARM" 2>/dev/null
fi
[[ -f "$ARM" || -f "$PPTX" ]] || { rm -f "$COUNTER" "$MUTED" 2>/dev/null; exit 0; }

# 백스톱용 서명 — 두 마커 합산(경로·mtime·스니펫 집합).
if command -v sha1sum >/dev/null 2>&1; then
  SIG=$(cat "$ARM" "$PPTX" 2>/dev/null | sort | sha1sum | cut -d' ' -f1)
else
  SIG=$(cat "$ARM" "$PPTX" 2>/dev/null | wc -c | tr -d ' ')
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
    { printf '[deck-gate backstop] auto-pass:\n%s\n' "$reason"; } >> "$OVLOG" 2>/dev/null
    rm -f "$COUNTER" 2>/dev/null
    printf '%s\n' "$SIG" > "$MUTED" 2>/dev/null
    emit_ctx "덱 게이트 백스톱: ${N}회 무진행 자동통과. 미충족 항목(렌더 눈검증/에이전트 경유)을 사용자에게 명시 보고할 것."
    exit 0
  fi
  emit_block "$reason (무진행 ${CNT}/${N}, 도달 시 자동통과)"
  exit 0
}

FAILS=""

# ── 검사 ①②: 소스 arm — preflight green + 렌더 신선도 (기존) ──
if [[ -f "$ARM" && -n "$NEWEST_SRC" && -n "$PROJ" ]]; then
  TO=""
  command -v timeout >/dev/null 2>&1 && TO="timeout 120"
  pf_out=$(cd "$PROJ" && $TO uv run python .claude/skills/pptx-build/scripts/preflight_dense.py $DENSE_STEMS 2>&1)
  pf_rc=$?
  if [[ "$pf_rc" -ne 0 ]]; then
    FAILS+="✗ preflight red (rc=${pf_rc}): $(printf '%s' "$pf_out" | tail -c 300)
"
  fi
  FRESH=$(find "$PROJ" -path '*/.venv' -prune -o -path '*/.git' -prune -o \
    -name 's[0-9][0-9].png' -newer "$NEWEST_SRC" -print 2>/dev/null | head -1)
  if [[ -z "$FRESH" ]]; then
    FAILS+="✗ 렌더 없음/스테일: '${NEWEST_SRC##*/}' 편집 이후 sNN.png 없음. render_deck.ps1로 렌더→눈검증.
"
  fi
fi

# ── 검사 ③: 산출물 arm — 덱 에이전트 "완료 실행" 경유 (2026-07-24 · 페어링 교체 2026-07-24) ──
# 문자열 존재(옛 grep)가 아니라 tool_use↔tool_result 페어링으로 "실제 완료"를 본다: 덱 에이전트
# subagent_type를 가진 tool_use의 tool_use_id에 대응하는 tool_result가 있고, 그것이 거부
# (is_error:true / "The user doesn't want to proceed" …)도 무응답(미완료)도 아닐 때만 경유 인정.
# 실증(2026-07-24): 거부된 deck-smith 호출도 tool_use JSON은 transcript에 남아 옛 grep이
# 거짓통과했다(사용자 tool use rejected). 페어링은 그 거부를 걸러낸다.
# ── 산출물 분류 (2026-07-26 사용자 승인) ──────────────────────────────────────────
# 게이트가 보는 신호는 ".pptx가 저장됐다" 하나뿐이라 **덱이 아닌 산출물**까지 덱 작업으로
# 오판했다. 실측된 오판 2종:
#   ① 골든 회귀 확인 — 부품화(장에 굳은 도해를 figures/로 꺼내기)는 시각을 바꾸지 않는 구조
#      이관이고, 무손실 증명이 골든덱을 다시 찍어 compare_golden 픽셀 동일로 대조하는 것이다.
#      이 재빌드가 덱 수정으로 잡혀 남은 12회차를 전부 막았다(2026-07-26 s04 회차 실측).
#   ② 부품 목업 — fig_mockup이 results/검토/부품_*.pptx를 찍는다. 덱이 아니라 부품 한 종의
#      SAMPLES 렌더인데 확장자가 같아 매 세션 종료가 막히고 백스톱에 의존했다.
# 그래서 armed 경로를 종류별로 가르고 **실전 덱이 하나라도 있을 때만** 검사 ③을 건다.
# 골든은 무조건 면제가 아니다 — compare 결과가 그 pptx보다 **새롭고 불일치 0건**일 때만이다.
# 시각을 바꾸는 골든 재작업은 compare가 빨개져 이 문을 못 지나고 deck-smith 요구로 되돌아간다.
# 훅은 <프로젝트>/.claude/hooks/ 에 있다 — 상대경로 해석의 기준점(소스 arm 유무와 무관).
SELF_ROOT=$(to_unix "$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)")
NEED_AGENT=0
if [[ -f "$PPTX" ]]; then
  while IFS=$'	' read -r _ns gpath; do
    gpath="${gpath%$''}"  # CRLF 방어 — 이스케이프로 쓴다(원문 CR 바이트는 안 먹는다)
    [[ -n "$gpath" ]] || continue
    gu=$(to_unix "$gpath")
    base="${gu##*/}"
    # arm은 명령에서 뽑은 **그대로** 적는다 — 상대경로도 절대경로도 온다. 그래서 파일명으로 가른다.
    case "$base" in
      .pptx)          # 경로를 못 뽑은 신호(파일명 없음) — 덱이 저장됐다는 증거가 아니다
        continue ;;
      부품_*.pptx)    # 부품 목업 — 덱이 아니라 부품 한 종의 SAMPLES 렌더
        continue ;;
      golden-deck.pptx)
        # 상대경로면 훅 자기 위치에서 루트를 뽑는다 — 소스 arm(PROJ)은 없을 수도 있다.
        if [[ "$gu" == /* ]]; then GROOT="${gu%/golden/golden-deck.pptx}"; else GROOT="$SELF_ROOT"; fi
        CMP="$GROOT/golden/variants/compare_full.md"
        if [[ -n "$GROOT" && -f "$CMP" && "$CMP" -nt "$GROOT/golden/golden-deck.pptx" ]]            && grep -q '불일치 0건' "$CMP" 2>/dev/null; then
          continue
        fi
        NEED_AGENT=1 ;;
      *)
        # 실전 덱 후보 — 다만 **실제로 파일이 있어야** 센다. arm은 명령 텍스트에서 뽑은
        # 토큰이라 셸 변수가 안 풀린 문자열($HOME/...)이나 조사용 가짜 경로도 들어온다
        # (2026-07-26 실측: 게이트를 조사하는 명령 자체가 게이트를 무장시켰다).
        # 진짜 덱은 디스크에 있다 — 없는 경로는 저장의 증거가 아니다.
        if [[ "$gu" == /* ]]; then ABS="$gu"; else ABS="$SELF_ROOT/$gu"; fi
        [[ -f "$ABS" ]] || continue
        NEED_AGENT=1 ;;
    esac
  done < "$PPTX"
fi

if [[ -f "$PPTX" && "$NEED_AGENT" -eq 1 ]]; then
  TR_RAW=$(printf '%s' "$INPUT" | python -c "import json,sys;print(json.load(sys.stdin).get('transcript_path','').replace(chr(92),'/'))" 2>/dev/null)
  TR=$(to_unix "$TR_RAW")
  if [[ -n "$TR" && -f "$TR" ]]; then
    VIA=$(TR="$TR" python - <<'PY' 2>/dev/null
import json, os
DECK = {"deck-smith", "pptx-builder", "deck-composer", "consistency-qa"}
uses = {}     # tool_use_id -> subagent_type (덱 에이전트 Agent 호출만)
results = {}  # tool_use_id -> (is_error, content_str)
try:
    fh = open(os.environ["TR"], encoding="utf-8")
except Exception:
    raise SystemExit  # 못 읽으면 아무것도 출력 안 함 → 아래서 fail-open
for line in fh:
    line = line.strip()
    if not line:
        continue
    try:
        obj = json.loads(line)
    except Exception:
        continue
    msg = obj.get("message")
    if not isinstance(msg, dict):
        continue
    content = msg.get("content")
    if not isinstance(content, list):
        continue
    for b in content:
        if not isinstance(b, dict):
            continue
        t = b.get("type")
        if t == "tool_use":
            st = (b.get("input") or {}).get("subagent_type")
            if st in DECK:
                uses[b.get("id")] = st
        elif t == "tool_result":
            cc = b.get("content")
            cs = cc if isinstance(cc, str) else json.dumps(cc, ensure_ascii=False)
            results[b.get("tool_use_id")] = (b.get("is_error"), cs)
completed = False
for uid in uses:
    if uid not in results:  # 무응답(미완료·pending) — 경유 불인정
        continue
    err, cs = results[uid]
    if err is True:  # tool_result 에러/거부
        continue
    if cs.startswith("The user doesn't want to proceed"):  # 거부 문자열 방어
        continue
    completed = True
    break
print("PASS" if completed else "BLOCK")
PY
)
    if [[ "$VIA" == "BLOCK" ]]; then
      FAILS+="✗ 덱 산출물(.pptx) 저장 감지 — 덱 에이전트 완료 실행 없음(호출만 있고 거부/무응답이거나 아예 미경유). 메인 즉흥 덱 작업 금지: 수정·재작업은 deck-smith, 신규 빌드는 pptmaker 파이프라인 에이전트로 **실제 완료까지** 수행하라(CLAUDE.md 정체 루프).
"
    fi
    # VIA="" (python 오류) 또는 "PASS" → 차단 안 함(fail-open / 정상 통과). BLOCK만 적재.
  fi
  # transcript를 못 읽으면 fail-open(검사 생략) — 게이트 오류가 종료를 막지 않는다.
fi

if [[ -n "$FAILS" ]]; then
  backstop_or_block "종료 차단(덱 게이트). 이 세션에 덱 작업 흔적이 있는데 아래가 안 끝났다:
${FAILS}해소(green + 렌더 / 에이전트 경유) 후 종료하라. 의도적 중간 양보면 백스톱까지 대기(무진행 누적)."
fi

# 전부 green — 무장 해제.
rm -f "$ARM" "$PPTX" "$COUNTER" "$MUTED" 2>/dev/null
exit 0
