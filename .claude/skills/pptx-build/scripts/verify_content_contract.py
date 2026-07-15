"""content_contract 검증 — 침묵 폴백(골든 글 유출)이 막혔는지 전 레이아웃 전수 확인.

케이스: ① 키 해석 · ② 완비=통과(오탐 0) · ③ 부분 주입=FAIL · ④ 빈 값(None)=FAIL
      · ⑤ 함수 내부 폴백 패턴(`c["x"] or [...]`) 잔존 스캔

④⑤는 2026-07-15 제3자 반증이 찾은 CRITICAL(`points=None`으로 골든 원문 유출, 공장 문·오염
검사 동시 실명)의 회귀 방지다. 상세: docs/superpowers/specs/2026-07-15-gate-repair-design.md 부록 C.

사용: uv run python .claude/skills/pptx-build/scripts/verify_content_contract.py  (종료코드 0 = PASS)
"""

import inspect
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from goldenfab.content_contract import assert_content, required_keys  # noqa: E402
from goldenfab.registry import LAYOUTS  # noqa: E402

fails = []

# ── 전 레이아웃(15/15) 필수 키 해석 가능한가 ──
print("=== 필수 키 해석 (분모 N=15) ===")
ok = 0
for name, fn in LAYOUTS.items():
    try:
        req = required_keys(name, fn)
        assert req, f"{name}: 키 집합이 비었다"
        print(f"  {name:16} {len(req):2}키 OK")
        ok += 1
    except Exception as e:
        print(f"  {name:16} FAIL {e}")
        fails.append(f"required_keys({name})")
print(f"해석 {ok}/{len(LAYOUTS)}")

# ── 케이스 ②: 완비 → 통과(오탐 0). 전 레이아웃 전수 ──
print("\n=== 케이스 ②: content 완비 → 통과해야 정상 (분모 N=15) ===")
ok2 = 0
for name, fn in LAYOUTS.items():
    full = {k: "X" for k in required_keys(name, fn)}
    try:
        assert_content(name, fn, full, 1)
        print(f"  {name:16} PASS")
        ok2 += 1
    except Exception as e:
        print(f"  {name:16} 오탐! {str(e)[:70]}")
        fails.append(f"오탐({name})")
print(f"완비 통과 {ok2}/{len(LAYOUTS)}")

# ── 케이스 ③: 필수 키 1개 누락 → FAIL해야 정상. 전 레이아웃 전수 ──
print("\n=== 케이스 ③: 키 1개 누락 → FAIL해야 정상 (분모 N=15) ===")
ok3 = 0
for name, fn in LAYOUTS.items():
    req = sorted(required_keys(name, fn))
    partial = {k: "X" for k in req[1:]}  # 첫 키 제외
    try:
        assert_content(name, fn, partial, 1)
        print(f"  {name:16} 미검출! 누락 '{req[0]}'이 골든 글로 메워진다")
        fails.append(f"미검출({name})")
    except ValueError as e:
        assert req[0] in str(e), f"{name}: 에러에 누락 키 미명시"
        print(f"  {name:16} FAIL 검출 OK (누락 '{req[0]}' 명시)")
        ok3 += 1
print(f"부분주입 검출 {ok3}/{len(LAYOUTS)}")

# ── 케이스 ④: 빈 값(None) 주입 → FAIL해야 정상. 전 레이아웃 전수 ──
# 회귀 방지: 2026-07-15 제3자 반증이 `points=None`으로 골든 원문 유출을 재현했다
# (키가 있으니 통과 → 함수 내부 `c["x"] or [리터럴]` 폴백이 골든 글 주입).
print("\n=== 케이스 ④: 키는 있으나 값이 None → FAIL해야 정상 (분모 N=15) ===")
ok4 = 0
for name, fn in LAYOUTS.items():
    allnone = {k: None for k in required_keys(name, fn)}
    try:
        assert_content(name, fn, allnone, 1)
        print(f"  {name:16} 미검출! 빈 값이 골든 기본값으로 메워진다")
        fails.append(f"빈값미검출({name})")
    except ValueError:
        print(f"  {name:16} FAIL 검출 OK")
        ok4 += 1
print(f"빈값 검출 {ok4}/{len(LAYOUTS)}")

# ── 케이스 ⑤: DEFAULT 밖 함수 내부 폴백 패턴 잔존 검사 (원사고 클래스) ──
print('\n=== 케이스 ⑤: 함수 내부 폴백 패턴(`c["x"] or [...]`) 잔존 ===')
leaky = []
for name, fn in LAYOUTS.items():
    src = inspect.getsource(sys.modules[fn.__module__])
    # 주석 제거 후 스캔 — 이 패턴을 **설명하는** 주석이 코드로 오인되지 않도록
    code = "\n".join(ln.split("#", 1)[0] for ln in src.splitlines())
    for m in re.finditer(r"c\[\"(\w+)\"\]\s+or\s+[\[\(\"]", code):
        leaky.append(f"{name}:{m.group(1)}")
if leaky:
    print(f"  검출 {leaky} — DEFAULT로 올리고 폴백 제거할 것")
    fails.append(f"함수내폴백({leaky})")
else:
    print("  잔존 0건 OK (기본값은 DEFAULT dict에만 → 오염 검사가 볼 수 있다)")

print("\n" + ("VERIFY PASS" if not fails else f"VERIFY FAIL: {fails}"))
sys.exit(1 if fails else 0)
