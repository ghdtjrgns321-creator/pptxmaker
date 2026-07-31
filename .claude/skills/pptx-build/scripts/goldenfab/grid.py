"""골든 덱 GRID 상수 + 항목 수 파생 자 — 사용자 지정 (2026-07-10). 임의 좌표 금지.

슬라이드 13.333 x 7.5 inch. RIGHT_EDGE = 13.333 - 0.6 = 12.733 (12.x 착각 금지).

`pitch`(세로)·`track`(가로)는 **항목 수에서 좌표를 파생하는 단일 출처**다. 리터럴 좌표 리스트
(`step_y = [3.2, 3.98, 4.76]`)를 `zip`으로 물리면 content가 4개일 때 **4번째가 조용히 사라지고**
2개일 때 빈칸이 남는다 — 게이트는 통과한다(하한도 안 넘고 경계도 안 넘으므로). design-rules §F·
§E-2 수반 규율 "항목 수가 content로 바뀌는 모든 세로 나열"의 기계 자리가 여기다.
(2026-07-25: 같은 식이 `s08_variants._pitch`·`s08_dense._pitch`(열화 사본)·`s10.point_pitch`
3곳에 흩어져 있던 것을 여기로 합쳤다.)
"""

MARGIN_L = 0.6
RIGHT_EDGE = 12.733
COL_L_X = 0.6
COL_L_W = 3.1  # 좌 컬럼 끝 3.7
COL_R_X = 4.2
COL_R_W = 8.533  # 우 컬럼 끝 12.733

RULE_Y = 1.45  # 제목 밑 구분선
CONTENT_TOP = 1.75  # 콘텐츠 시작 — 위 여백 이보다 크면 안 됨
SUBHEAD_Y = 1.75  # 좌·우 서브헤드 top 동일
FLOW_ROW1_Y = 2.60
FLOW_H = 0.72
FLOW_ROW2_Y = 3.80
LOW_HEAD_Y = 4.95
BOX_Y = 5.30
BOX_H = 1.05  # 하단 2박스 bottom 6.35
BOX_GAP = 0.3
BOX_W = (COL_R_W - BOX_GAP) / 2  # 4.1165
BAR_Y = 6.60
BAR_H = 0.45
SOURCE_Y = 7.15

CONTENT_BOTTOM = BOX_Y + BOX_H  # 6.35 — 콘텐츠 하한
V_RULE_X = (COL_L_X + COL_L_W + COL_R_X) / 2  # 3.95 — 컬럼 사이 세로 룰


CAP_PITCH = 0.42  # 세로 피치 상한 = 골든 리듬. 자리가 남아도 이보다 벌어지지 않는다


def pitch(n, top, bottom, h, cap=CAP_PITCH, *, what="항목"):
    """세로 피치를 **항목 수에서 파생**한다 — design-rules §F(2026-07-15, S10 하한 초과에서 확정).

    리터럴 피치(`top + i*0.42`)를 쓰면 content로 항목이 늘어난 순간 조용히 하한을 넘는다.

    파생만으로는 부족하다: 항목이 아주 많으면 피치가 칩 높이보다 작아져 **칩끼리 겹친다**.
    하한(6.35)은 안 넘으므로 경계 검사는 통과한다 — 조용히 틀리는 값이다(2026-07-15 케이스
    테스트가 잡음: 개념 18개 → 피치 0.109 < 높이 0.24). 그래서 **시끄럽게 죽인다**:
    이 레이아웃이 수용할 수 있는 항목 수를 넘으면 빌드를 세운다. `add_diagram`이 모르는 layout에
    ValueError를 내는 것과 같은 원리다 — 침묵 폴백은 조용히 틀린 그림을 낳는다.

    cap 기본 0.42 = 골든 리듬 상한(여유가 있을 때 무한히 벌어지지 않게).
    """
    if n <= 1:
        return h
    p = min(cap, (bottom - top - h) / (n - 1))
    if p < h:
        room = int((bottom - top - h) / h) + 1
        raise ValueError(
            f'{what} {n}개는 이 구획({top}~{bottom}")에 안 들어간다 — 피치 {p:.3f} < 칩 높이 {h}'
            f"라 칩이 겹친다. 최대 {room}개. content를 줄이거나 레이아웃을 다시 설계할 것."
        )
    return p


def track(n, x0, x1, gap, min_w, *, what="칼럼"):
    """가로 칼럼 폭을 **개수에서 파생**한다 — `pitch`의 가로 짝(2026-07-25).

    세로와 같은 병이 가로에도 있다: `cw = (tw - 2*gap)/3`처럼 개수를 리터럴로 박으면 항목이
    4개가 된 순간 마지막 칼럼이 우변을 넘고(check_bounds가 잡으면 다행, 밴드 안이면 겹침),
    2개면 오른쪽에 죽은 여백이 남는다. 폭이 최소 판독폭 아래로 내려가면 시끄럽게 죽인다 —
    글자가 세로로 쪼개진 칼럼은 오딧을 통과하면서 사람 눈에만 깨져 보인다.
    """
    if n < 1:
        raise ValueError(f"{what} 0개 — 그릴 것이 없다. content를 확인할 것.")
    w = (x1 - x0 - (n - 1) * gap) / n
    if w < min_w:
        room = int((x1 - x0 + gap) / (min_w + gap))
        raise ValueError(
            f'{what} {n}개는 이 폭({x0}~{x1}")에 안 들어간다 — 칼럼 {w:.3f} < 최소 {min_w}. '
            f"최대 {room}개. content를 줄이거나 세로 나열로 바꿀 것."
        )
    return w
