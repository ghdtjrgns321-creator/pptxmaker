"""goldenfab.audit — 슬라이드 기계 오딧. **채점자가 캔 규칙을 코드로 제련해 두는 곳.**

## 왜 이 파일이 있나 (2026-07-15)

S4 한 장에 제3자 채점 2회 = **30분**이 들었다. 그런데 두 채점자가 낸 FAIL 9건을 분류하니
**7건이 순수 산수**였다 — 대비비·채움률·높이 종류·텍스트 중복·라벨 정렬·같은 fill 충돌·accent 계수.
15분짜리 모델이 필요한 일이 아니었다. 내 오딧이 그 규칙을 **안 갖고 있어서** 못 잡았을 뿐이다.

즉 **채점자는 검사기가 아니라 광산이다.** 한 번 캔 광석을 코드로 내려두면 다시 캘 필요가 없다.
design-rules의 철학("확정할 때마다 판단을 박제")과 같되, **산문이 아니라 실행되는 코드로** 박제한다
— 산문 규칙은 잊히지만 assert는 건너뛸 수 없다(P1 주석의 자기 인정).

## 운영 규칙

1. 레이아웃을 만들면 먼저 이 오딧을 통과시킨다(2초).
2. 제3자 채점은 **1회만**. 재채점 금지 — 오늘 실증: 6/10 → 5/10으로 **안 수렴**했고 두 채점자가
   **서로 다른 항목**을 찾았다. 주관적 채점에는 고정점이 없다.
3. 채점자가 새 결함을 찾으면 **반드시 여기에 규칙으로 내린다**. 산문으로 남기면 다음 장에서 또 캔다.
4. 회를 거듭할수록 채점자가 찾을 게 줄어든다. 그게 이 파일이 하는 일이다.

## 규칙 출처

각 규칙에 발견 경위를 적는다 — 왜 있는지 모르는 규칙은 다음 사람이 지운다.
"""

import re

from pptx.enum.shapes import MSO_SHAPE

EMU = 914400
AIR_MIN = 0.12  # §6 공기 하한
ACCENT_MAX = 4  # §2 accent 상한
CONTENT_BOTTOM = 6.35  # P2③
RIGHT_EDGE = 12.733  # P2④
CONTRAST_MIN = 4.5  # WCAG AA (일반 텍스트)


# ── 픽셀·색 산수 ──────────────────────────────────────────────


def _lum(hexstr):
    """상대 휘도 (WCAG)."""

    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4

    r, g, b = (int(hexstr[i : i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(fg, bg="FFFFFF"):
    """대비비. 판정 단어가 산문보다 흐린지 재는 데 쓴다."""
    a, b = _lum(str(fg)), _lum(str(bg))
    hi, lo = max(a, b), min(a, b)
    return round((hi + 0.05) / (lo + 0.05), 2)


def shape_kind(sh):
    try:
        return sh.auto_shape_type
    except Exception:
        return None


def fill_hex(sh):
    try:
        if sh.fill.type is not None:
            return str(sh.fill.fore_color.rgb)
    except Exception:
        pass
    return None


def line_hex(sh):
    try:
        if sh.line.color.type is not None:
            return str(sh.line.color.rgb)
    except Exception:
        pass
    return None


def runs_of(sh):
    """(텍스트, hex색, pt, bold) 목록."""
    out = []
    if not sh.has_text_frame:
        return out
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            try:
                col = str(r.font.color.rgb)
            except Exception:
                col = None
            out.append((r.text, col, r.font.size.pt if r.font.size else None, r.font.bold))
    return out


def box(sh):
    return (
        sh.left / EMU,
        sh.top / EMU,
        sh.width / EMU,
        sh.height / EMU,
    )


# ── 규칙 ──────────────────────────────────────────────────────


def check_accent(shapes, accent, cap=ACCENT_MAX):
    """§2 accent 상한. **채움만 세면 안 된다 — 테두리·커넥터 선도 accent 예산이다.**

    출처: v4에서 채움·런만 세서 4로 통과시켰는데 제3자가 7군데를 셌다(선 미계수).
    """
    f = [s for s in shapes if fill_hex(s) == accent]
    ln = [s for s in shapes if line_hex(s) == accent]
    rn = [r for s in shapes for r in runs_of(s) if r[1] == accent]
    tot = len(f) + len(ln) + len(rn)
    ok = tot <= cap
    return ok, f"accent 채움 {len(f)} + 선 {len(ln)} + 런 {len(rn)} = {tot} (상한 {cap})", tot


def check_verdict_contrast(shapes, prose_pt):
    """**판정을 말하는 단어는 그것을 설명하는 산문보다 흐릴 수 없다.**

    출처: v5에서 결론 단어 '↩ 되돌아옴'(3.22) · '✕ 불가역'(3.41)이 그것을 설명하는
    산문(8.47)보다 2.5배 흐렸다. 단순 명찰은 16.16을 가져갔다. 제3자 2회 모두 지적.
    """
    bad = []
    for s in shapes:
        for text, col, pt, _bold in runs_of(s):
            if not col or not text.strip():
                continue
            if re.match(r"^\s*[✕✓↩→]", text) or "불가역" in text or "없다" in text:
                cr = contrast(col)
                if cr < CONTRAST_MIN:
                    bad.append((text.strip()[:18], col, cr))
    return not bad, f"판정 단어 대비 < {CONTRAST_MIN}: {len(bad)} {bad[:3]}", len(bad)


def check_fill_ratio(shapes, min_ratio=0.30, min_w=1.5):
    """도형이 글자에 비해 과하게 넓지 않은가(죽은 회색).

    출처: v5의 한계 칩이 폭 3.30"에 채움률 14~22%였다 — 단어와 커넥터 사이 238px 공백.
    글자폭은 pt·글자수로 근사(정밀 측정 불가) — 한국어 1자 ≈ 0.8×pt.
    """
    bad = []
    for s in shapes:
        if not s.has_text_frame or fill_hex(s) is None:
            continue
        _x, _y, w, _h = box(s)
        if w < min_w:
            continue
        rs = runs_of(s)
        if not rs:
            continue
        txt = "".join(r[0] for r in rs)
        pt = max((r[2] or 9) for r in rs)
        est = len(txt.strip()) * pt * 0.8 / 72
        ratio = est / w
        if ratio < min_ratio:
            bad.append((txt.strip()[:14], round(w, 2), f"{ratio:.0%}"))
    return not bad, f"채움률 < {min_ratio:.0%}: {len(bad)} {bad[:3]}", len(bad)


def check_duplicate_nodes(shapes):
    """**구별 장치 없는** 반복 = 재탕(P4③).

    출처: v4에서 '2종 · 허위 확정'이 두 밴드에 같은 글자·같은 fill·같은 테두리로 있었다.

    ⚠ 테두리(색·점선)가 다르면 재탕이 아니라 **대조**다 — v7의 두 ◇ '근거 충분?'은 실선(성립)
    vs 점선(성립 불가)으로 뒤가 앞을 부정한다. 제3자도 "반복이 아니라 대조"로 PASS를 줬는데
    초판 규칙이 (텍스트, 채움)만 봐서 오탐을 냈다. 구별 장치까지 서명에 넣는다.
    """
    seen = set()
    dup = []
    for s in shapes:
        f = fill_hex(s)
        if not f or not s.has_text_frame:
            continue
        t = s.text_frame.text.strip()
        if len(t) < 3:
            continue
        try:
            dash = str(s.line.dash_style)
        except Exception:
            dash = None
        sig = (t, f, line_hex(s), dash)  # 구별 장치 = 테두리 색 + 점선 여부
        if sig in seen:
            dup.append(t[:20])
        seen.add(sig)
    return not dup, f"구별 장치 없는 반복: {len(dup)} {dup[:3]}", len(dup)


def check_ink_collision(shapes, ink, allow=1):
    """같은 채움색이 서로 다른 뜻을 지지 않는가(P4⑩) — 근사: 그 색 채움 도형 수.

    출처: v4에서 '재무제표 반영'(파국)과 결론 바(우리 원칙)가 **바이트 동일**한
    primary+샤프+흰볼드였다. 훑는 사람은 검은 덩어리 둘을 같은 종류로 분류한다.
    """
    f = [s for s in shapes if fill_hex(s) == ink]
    txt = [s.text_frame.text[:16] for s in f if s.has_text_frame and s.text_frame.text.strip()]
    return len(f) <= allow, f"{ink} 채움 {len(f)} (허용 {allow}) {txt}", len(f)


def check_node_heights(shapes, top_max, left_min=1.0, h_min=0.3):
    """같은 계열(행 노드) 높이 균일(P4⑧).

    출처: v4에서 확정 61 / 2종 69 / 일반LLM 71 / band2 84px — 한 종류에 높이 4종.
    계열 제외: ◇(다른 도형) · 풀폭 바 · 뿌리 카드(흐름 시작점) · 작은 칩.
    """
    nodes = [
        s
        for s in shapes
        if fill_hex(s)
        and shape_kind(s) != MSO_SHAPE.DIAMOND
        and box(s)[2] < 11
        and box(s)[0] > left_min
        and box(s)[1] < top_max
        and box(s)[3] > h_min
    ]
    hs = sorted({round(box(s)[3], 3) for s in nodes})
    return len(hs) <= 1, f"행 노드 {len(nodes)}개 · 높이 종류 {hs}", len(hs)


def check_progress_shapes(shapes, allowed=0):
    """§5 셰브런·펜타곤은 **분기 없는 단순 진행**에만.

    출처: XOR을 셰브런 밴드로 그려 "확정한 다음 유보한다"는 거짓 인과를 만든 반려 4회.
    """
    prog = [s for s in shapes if shape_kind(s) in (MSO_SHAPE.CHEVRON, MSO_SHAPE.PENTAGON)]
    return len(prog) <= allowed, f"진행형 도형 {len(prog)} (허용 {allowed})", len(prog)


def check_picture_overlap(shapes):
    """그림이 다른 요소를 침범하지 않는가.

    출처(2026-07-15 사용자 버그 리포트): s10이 이미지 비율을 `3200/2000`(1.60)으로 **가정**하고
    폭 7.6"를 전제로 옆 칼럼을 8.6"에 놨는데, 실제 파일은 1899×926(**2.05**)이라 폭 9.74"로
    렌더돼 텍스트를 **1.74" 침범**했다. 그 `img_w`는 계산만 하고 안 쓰는 죽은 변수였다.
    → `kit.fit_picture`가 근본 수정(박스에 맞춤). 이 규칙은 재발 감시.

    ⚠ python-pptx는 접근할 때마다 **새 프록시**를 만들어 `s is p`로는 자기 자신도 못 거른다 —
    `_element` 동일성으로 비교해야 한다(이 함정에 한 번 빠져 "그림이 자기와 겹침"을 냈다).
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    pics = [s for s in shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    hits = []
    for p in pics:
        x, y, w, h = box(p)
        for s in shapes:
            if s._element is p._element:
                continue
            sx, sy, sw, sh = box(s)
            if sw > 13 and sh > 7:
                continue  # 풀블리드 배경
            if sx < x + w and x < sx + sw and sy < y + h and y < sy + sh:
                lbl = s.text_frame.text.strip()[:16] if s.has_text_frame else "(무텍스트)"
                hits.append(lbl)
    return not hits, f"PICTURE {len(pics)}개 · 침범 {len(hits)} {hits[:3]}", len(hits)


def check_bounds(shapes, bottom=CONTENT_BOTTOM, right=RIGHT_EDGE, skip_tops=()):
    """P2③ 콘텐츠 하한 · P2④ 풀폭 우변."""
    over, badr = [], []
    for s in shapes:
        x, y, w, h = box(s)
        if w > 13 and h > 7:
            continue  # 풀블리드 배경
        if any(abs(y - t) < 0.01 for t in skip_tops):
            continue
        if y + h > bottom + 0.01:
            over.append((round(y + h, 2), s.text_frame.text[:14] if s.has_text_frame else ""))
        if abs(x - 0.6) < 0.001 and w > 11 and abs(x + w - right) > 0.005:
            badr.append(round(x + w, 3))
    return (
        not (over or badr),
        f"bottom>{bottom}: {len(over)} {over[:2]} · 풀폭 right 위반: {badr}",
        len(over) + len(badr),
    )


def check_air(pairs, air_min=AIR_MIN):
    """§6 공기 — (이름, 위 요소 bottom, 아래 요소 top) 목록을 받아 간격 검사.

    전체 쌍 자동 검사는 오탐이 많다(헤더 킥커→헤드라인 0.02는 확정 설계) — 설계상
    붙으면 안 되는 지점만 명시적으로 넘긴다.
    """
    bad = [(n, round(t - b, 3)) for n, b, t in pairs if t - b < air_min]
    lines = [
        f"  {n:28} 공기 {t - b:+.3f} {'OK' if t - b >= air_min else '**FAIL**'}"
        for n, b, t in pairs
    ]
    return not bad, "\n".join(lines), len(bad)


def report(results, known=None):
    """(이름, ok, 상세, 위반수) 목록 → 출력 + 실패 목록.

    **위반 수는 규칙이 직접 반환한다** — 상세 문자열에서 정규식으로 긁으면 규칙마다 형식이
    달라 조용히 None이 되고 래칫이 죽는다(초판이 `채움 2 (허용 0)`을 못 읽어 그랬다).

    known: {규칙명: 기준선 위반 수} — 이 오딧 **이전에** 사용자 승인으로 확정된 장의 기존 빚.
      규칙을 끄지 않고 기준선으로 박는다(래칫): 기준선 이하면 통과, **넘으면 FAIL**.
      · 규칙을 빼면 hollow-PASS가 된다(존재≠검사).
      · 그대로 빨간불이면 게이트가 상시 적색이라 무시당하고, 결국 우회한다 —
        이 프로젝트가 이미 겪은 실패 경로다.
      기준선을 내리는 건 해당 장 재설계 때. 좋아지면 알려서 기준선을 조인다.
    """
    known = known or {}
    fails = []
    for name, ok, detail, n in results:
        base = known.get(name)
        if not ok and base is not None and n is not None and n <= base:
            mark = "DEBT" if n == base else "GOOD"
            note = "기지 빚(기준선)" if n == base else f"개선 {base}→{n} · 기준선을 낮출 것"
            print(f"{mark} {name:22} {detail}  ← {note}")
            continue
        print(f"{'OK  ' if ok else 'FAIL'} {name:22} {detail}")
        if not ok:
            fails.append(name)
    print("\n" + ("AUDIT PASS" if not fails else f"AUDIT FAIL: {fails}"))
    return fails
