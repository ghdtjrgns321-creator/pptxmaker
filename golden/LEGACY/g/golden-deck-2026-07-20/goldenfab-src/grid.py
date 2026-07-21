"""골든 덱 GRID 상수 — 사용자 지정 (2026-07-10). 스크립트는 이 값만 파생 사용, 임의 좌표 금지.

슬라이드 13.333 x 7.5 inch. RIGHT_EDGE = 13.333 - 0.6 = 12.733 (12.x 착각 금지).
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
