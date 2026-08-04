# ruff: noqa: E402
"""local-ai-assist 형태 갤러리 (pptmaker ②③) — 축 B 전용.

**내용은 아웃라인 계약이 이미 정했다.** 후보 4개는 전부 같은 내용이고, 다른 것은 배치·인코딩뿐.
후보는 스케치다 — 구획 비율은 실제, 라벨은 부품이 자리에 맞게 접고 "… 외 N개"로 알린다.
"""

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1] / ".claude/skills/pptx-build/scripts"))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from goldenfab import grid as G
from goldenfab.dense import compact_header, hrule, line_soft, source_line
from goldenfab.kit import add_box, new_presentation
from sketch import (
    C,
    cols,
    foot,
    head,
    rows,
    sk_area,
    sk_band,
    sk_bars,
    sk_bip,
    sk_cat,
    sk_chips,
    sk_cmp,
    sk_enc,
    sk_fan,
    sk_flow,
    sk_funnel,
    sk_gate,
    sk_gauge,
    sk_grid,
    sk_merge,
    sk_rfan,
    sk_shot,
    sk_split,
    sk_stack,
    sk_tape,
    sk_text,
    sk_trade,
)

IMG = Path(r"C:\Users\ghdtj\workspace\portfolio\local-ai-assist\images")

MX, MY = (0.60, 6.78), (1.06, 3.76)
MW, MH, LAB, PAD = 5.95, 2.58, 0.30, 0.12

# ═════════════════════════════════════════════════════ S4 · 문제 인식
AXES5 = [
    ("검증 범위", "표본 추출", "전표 전수"),
    ("분석 기준", "주관 판단 → 감사인 편차", "3모델 일관 기준"),
    ("금액 중요성", "고액 위주 → 소액 누락", "금액 규모 무관"),
    ("데이터 전처리", "수동 정렬·가공", "속성 자동 부여"),
    ("미정의 패턴", "정의된 조건만", "머신러닝이 추출"),
]
EFFECT3 = ["파생 변수 자동 생성", "분석 결과 시각화", "검토 비용 절감"]
LANE_OLD = ["표본 추출", "고액 위주 검토", "소액·미정의 누락"]
LANE_NEW = ["전량 투입", "3모델 일관 적용", "검토 목록 산출"]
SRC2 = ["금감원 감리 지적사례 230건", "회계감사기준서 240·520호"]


def s4_a(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.80, 2.35, 1.05))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "전수",
        "표본 추출의 구조",
        "표본은 못 본 데이터를 못 본 채 결론을 낸다. 다섯 축 전부에서 갈린다.",
    )
    yy = head(sl, bx, y, bw, "기존 표본 분석 ↔ AI 기반 전수 분석")
    sk_cmp(sl, bx, yy, bw, h - (yy - y), AXES5, lt="표본", rt="전수")
    yy = head(sl, cx, y, cw, "기대 효과")
    sk_flow(sl, cx, yy, cw, h - (yy - y), EFFECT3)


def s4_b(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.30, 1.15))
    yy = head(sl, ax, y, aw, "다섯 축에서 갈린다")
    sk_cmp(sl, ax, yy, aw, h - (yy - y), AXES5, lt="표본", rt="전수")
    yy = head(sl, bx, y, bw, "같은 전표, 두 경로 — 근거는 실무 문서")
    sk_split(sl, bx, yy, bw, h - (yy - y), "전표 113,465건 · 감리 230건", LANE_OLD, LANE_NEW)


def s4_c(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (1.25, 1.0))
    yy = head(sl, x, ty, w, "다섯 축 — 위가 표본, 아래가 전수")
    n = len(AXES5)
    cwid = (w - (n - 1) * 0.10) / n
    for i, (ax_, ll, rr) in enumerate(AXES5):
        cxx = x + i * (cwid + 0.10)
        add_box(sl, cxx, yy, cwid, 0.22, fill=C["primary"], line=None)
        from sketch import tx

        tx(
            sl,
            cxx + 0.03,
            yy,
            cwid - 0.06,
            0.22,
            ax_,
            6,
            C["bg"],
            bold=True,
            align=PP_ALIGN.CENTER,
            anc=MSO_ANCHOR.MIDDLE,
        )
        add_box(sl, cxx, yy + 0.26, cwid, 0.34, fill=C["bg_alt"], line=None)
        tx(
            sl,
            cxx + 0.04,
            yy + 0.26,
            cwid - 0.08,
            0.34,
            ll,
            6,
            C["muted"],
            align=PP_ALIGN.CENTER,
            anc=MSO_ANCHOR.MIDDLE,
            ls=1.15,
        )
        add_box(sl, cxx, yy + 0.64, cwid, 0.34, fill=C["bg"], line=C["primary"], line_w=0.75)
        tx(
            sl,
            cxx + 0.04,
            yy + 0.64,
            cwid - 0.08,
            0.34,
            rr,
            6,
            C["primary"],
            align=PP_ALIGN.CENTER,
            anc=MSO_ANCHOR.MIDDLE,
            ls=1.15,
        )
    (px, pw), (qx, qw) = cols(x, w, (1.15, 1.0))
    yy = head(sl, px, by, pw, "기대 효과")
    sk_flow(sl, px, yy, pw, bh - (yy - by), EFFECT3)
    yy = head(sl, qx, by, qw, "시스템은 확정하지 않는다")
    sk_cat(
        sl,
        qx,
        yy,
        qw,
        bh - (yy - by),
        [("시스템", "여기까지", "1차 선별"), ("감사인", "여기부터", "부정 확정")],
        aw_r=0.30,
    )


def s4_d(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (0.92, 1.60))
    yy = head(sl, ax, y, aw, "같은 전표를 두 방식에")
    sk_split(sl, ax, yy, aw, h - (yy - y), "전표 113,465건", LANE_OLD, LANE_NEW)
    (ty, th), (by, bh) = rows(y, h, (1.70, 0.75))
    yy = head(sl, bx, ty, bw, "다섯 축 비교")
    sk_cmp(sl, bx, yy, bw, th - (yy - ty), AXES5, lt="표본", rt="전수")
    sk_tape(
        sl,
        bx,
        by,
        bw,
        bh,
        "근거는 실무 문서",
        ["감리 230건", "기준서 240호", "기준서 520호"],
        hot=(0,),
    )


# ═════════════════════════════════════════════════════ S5 · 시스템 요약
PIPE5 = [
    "전표 원장 355,786행",
    "파생 변수 자동 부여",
    "분석 모델 3종 병렬",
    "검토 조합 구성",
    "검토 대상 전표 목록",
]
MODELS3 = [
    ("룰 기반 검증 29종", "사전 정의 조건에 해당? → 근거가 명확한 목록"),
    ("분석적 검토 5종", "그룹 분포에 통계적 이상치? → 참고 신호"),
    ("비지도 학습 VAE", "정의되지 않은 패턴? → 신종·복합 후보"),
]
IN3 = ["전수 스크리닝", "속성 자동 전처리", "1차 선별"]
OUT2 = ["증빙 확인", "부정 여부 확정"]


def s5_a(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.0, 1.55))
    yy = head(sl, ax, y, aw, "전표에서 검토 목록까지")
    sk_flow(sl, ax, yy, aw, h - (yy - y), PIPE5, dark=(4,))
    yy = head(sl, bx, y, bw, "3종 병렬 — 가지를 뻗어 하나씩")
    sk_rfan(sl, bx, yy, bw, h - (yy - y), "분석 모델\n3종 병렬", MODELS3, nw_r=0.26)


def s5_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (0.72, 1.55))
    yy = head(sl, x, ty, w, "전표에서 검토 목록까지")
    sk_band(sl, x, yy, w, 0.40, PIPE5, hot=(2, 4))
    (px, pw), (qx, qw) = cols(x, w, (1.75, 1.0))
    yy = head(sl, px, by, pw, "3종은 각각 다른 질문에 답한다")
    sk_rfan(sl, px, yy, pw, bh - (yy - by), "3종 병렬", MODELS3, nw_r=0.22)
    yy = head(sl, qx, by, qw, "어디서 멈추나")
    sk_gate(sl, qx, yy, qw, bh - (yy - by), IN3, OUT2)


def s5_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.72, 1.0, 1.30))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "3모델",
        "감사인을 대체하지 않는다",
        "1차 스크리닝까지만 한다. 세 모델이 각기 다른 질문에 답하고, 확정은 사람이 한다.",
    )
    yy = head(sl, bx, y, bw, "파이프라인")
    sk_flow(sl, bx, yy, bw, h - (yy - y), PIPE5, dark=(4,))
    (ty, th), (by, bh) = rows(y, h, (1.45, 0.90))
    yy = head(sl, cx, ty, cw, "3종 병렬")
    sk_fan(sl, cx, yy, cw, th - (yy - ty), "전표 데이터", [m[0] for m in MODELS3])
    yy = head(sl, cx, by, cw, "어디서 멈추나")
    sk_cat(
        sl,
        cx,
        yy,
        cw,
        bh - (yy - by),
        [("시스템", "여기까지", "1차 선별"), ("감사인", "여기부터", "부정 확정")],
        aw_r=0.30,
    )


def s5_d(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.50, 1.0))
    yy = head(sl, ax, y, aw, "3종 병렬 — 가지를 뻗어 하나씩")
    sk_rfan(sl, ax, yy, aw, h - (yy - y), "분석 모델\n3종 병렬", MODELS3, nw_r=0.24)
    (ty, th), (by, bh) = rows(y, h, (1.30, 0.90))
    yy = head(sl, bx, ty, bw, "어디서 멈추나")
    sk_gate(sl, bx, yy, bw, th - (yy - ty), IN3, OUT2)
    yy = head(sl, bx, by, bw, "파이프라인")
    sk_tape(
        sl, bx, yy, bw, bh - (yy - by), "", ["원장", "파생", "3모델", "조합", "목록"], hot=(2, 4)
    )


# ═════════════════════════════════════════════════════ S7 · 정상 데이터
STEPS8 = [
    "1 계정과목표",
    "2 마스터데이터",
    "3 문서흐름",
    "4 업무 이벤트",
    "5 전표 생성",
    "6 이상 주입",
    "7 잔액 추적",
    "8 CSV 출력",
]
GATE3 = ["허용 계정 통제", "계정결정표 교차검증", "대차 균형 확인"]
FLOWS7 = [
    ("결산 R2R", "6", 1.00, True),
    ("자산 A2R", "3", 0.50, False),
    ("자금 TRE", "3", 0.50, False),
    ("구매 P2P", "2", 0.33, False),
    ("판매 O2C", "2", 0.33, False),
    ("급여·관계사", "4", 0.66, False),
]
ORIGIN3 = ["EY 오픈소스 DataSynth", "K-IFRS 제조업으로 개편", "YAML 17개 환경 설정"]
OUT2S = ["정상 원장 355,786행", "66개 지표로 검증"]


def _loop(sl, x, y, w, h, steps, gate):
    """진행 + 관문 + 되돌림 호 — 통과 못하면 앞으로 돌아간다."""
    from sketch import arrow, fold, hline, line_soft, tail, tx, vline
    from sketch import box as _b

    gh, back_h = 0.26, 0.24
    show, om, bh, pi, ty = fold(steps, y, y + h - gh - back_h - 0.08, min_h=0.15, cap=0.28)
    lx = x + 0.32
    for i, lab in enumerate(show):
        by = y + i * pi
        _b(sl, lx, by, w - 0.32, bh, lab, fill=C["bg_alt"], line=None, size=6.5)
        if i < len(show) - 1:
            vline(sl, lx + (w - 0.32) / 2, by + bh, pi - bh, line_soft)
    tail(sl, lx, ty, w - 0.32, om, "단계")
    gy = y + h - gh - back_h - 0.02
    _b(
        sl,
        lx,
        gy,
        w - 0.32,
        gh,
        " · ".join(gate),
        fill=C["primary"],
        color=C["bg"],
        size=6,
        shape="round",
    )
    vline(sl, x + 0.14, y + 0.09, gy + gh / 2 - (y + 0.09), C["accent"], 1.1)
    hline(sl, x + 0.14, gy + gh / 2, 0.18, C["accent"], 1.1)
    arrow(sl, x + 0.14, y + 0.09, lx, y + 0.09, C["accent"], 1.1)
    tx(
        sl,
        x,
        gy + gh + 0.03,
        w,
        back_h,
        "미통과 전표는 즉시 폐기 후 재생성",
        6,
        C["accent"],
        bold=True,
        align=PP_ALIGN.CENTER,
    )


def s7_a(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.78, 1.20, 1.05))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "8단계",
        "실데이터가 없어 직접 만들었다",
        "EY 오픈소스를 K-IFRS 제조업으로 개편. 무결성을 통과 못한 전표는 폐기하고 다시 만든다.",
    )
    yy = head(sl, bx, y, bw, "생성 순서와 무결성 관문")
    _loop(sl, bx, yy, bw, h - (yy - y), STEPS8, GATE3)
    (ty, th), (by, bh) = rows(y, h, (1.50, 0.85))
    yy = head(sl, cx, ty, cw, "거래 시나리오 20종 · 흐름 7")
    sk_bars(sl, cx, yy, cw, th - (yy - ty), FLOWS7, lab_w=0.74, val_w=0.20)
    yy = head(sl, cx, by, cw, "산출")
    sk_flow(sl, cx, yy, cw, bh - (yy - by), OUT2S, dark=(0,))


def s7_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (0.66, 1.75))
    yy = head(sl, x, ty, w, "생성 순서 8단계")
    sk_band(sl, x, yy, w, 0.40, STEPS8, hot=(4, 6))
    (px, pw), (qx, qw) = cols(x, w, (1.0, 1.15))
    yy = head(sl, px, by, pw, "무결성은 어디서 지켜지나")
    sk_merge(sl, px, yy, pw, bh - (yy - by), GATE3, "미통과 전표는 폐기 후 재생성")
    yy = head(sl, qx, by, qw, "거래 시나리오 20종 · 흐름 7")
    sk_bars(sl, qx, yy, qw, bh - (yy - by), FLOWS7, lab_w=0.80, val_w=0.22)


def s7_c(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.0, 1.45))
    yy = head(sl, ax, y, aw, "생성 순서 8단계")
    sk_flow(sl, ax, yy, aw, h - (yy - y), STEPS8, dark=(4, 6))
    (ty, th), (by, bh) = rows(y, h, (1.30, 1.05))
    yy = head(sl, bx, ty, bw, "흐름 7종 → 시나리오 20개")
    sk_fan(
        sl,
        bx,
        yy,
        bw,
        th - (yy - ty),
        "업무 흐름 7종 · 시나리오 20개",
        ["결산 R2R (6)", "자산 A2R (3) · 자금 TRE (3)", "구매·판매·급여·관계사 (8)"],
        hot=(0,),
    )
    yy = head(sl, bx, by, bw, "생성 직후 무결성 검증")
    sk_cat(
        sl,
        bx,
        yy,
        bw,
        bh - (yy - by),
        [
            ("허용 계정 통제", "차단", "불가 계정 유입"),
            ("계정결정표 교차검증", "재대조", "차대 방향·성격"),
            ("대차 균형 확인", "폐기·재생성", "미통과 전표"),
        ],
        aw_r=0.38,
    )


def s7_d(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (1.0, 1.20, 1.0))
    yy = head(sl, ax, y, aw, "무엇에서 시작했나")
    sk_flow(sl, ax, yy, aw, h - (yy - y), ORIGIN3, dark=(1,))
    yy = head(sl, bx, y, bw, "생성 루프 — 통과 못하면 되돌아간다")
    _loop(sl, bx, yy, bw, h - (yy - y), STEPS8, GATE3)
    (ty, th), (by, bh) = rows(y, h, (1.50, 0.85))
    yy = head(sl, cx, ty, cw, "시나리오 20종")
    sk_bars(sl, cx, yy, cw, th - (yy - ty), FLOWS7, lab_w=0.74, val_w=0.20)
    yy = head(sl, cx, by, cw, "무엇이 나왔나")
    sk_flow(sl, cx, yy, cw, bh - (yy - by), OUT2S, dark=(0,))


# ═════════════════════════════════════════════════════ S8 · 이상 주입
FS_L = ["매출", "매입", "결산", "관계사", "급여"]
FS_R = ["FS01·05", "FS09·10", "FS03·04", "FS06·07", "FS02·08·12", "FS11·13", "FS14"]
FS_LINK = [(0, 0), (0, 1), (1, 2), (1, 3), (2, 4), (3, 5), (4, 6)]
MAKE4 = ["1 베이스 선택", "2 목표 속성 변이", "3 시퀀스 연결", "4 흔적 제거"]
CHAIN4 = ["매출 계상", "가공 회수", "매출채권 담보 차입", "차기 역분개"]
FS14 = [
    "FS01 가공매출",
    "FS02 진행률",
    "FS03 현금횡령",
    "FS04 횡령은폐",
    "FS05 순환거래",
    "FS06 부채누락",
    "FS07 재고과대",
    "FS08 자본화",
    "FS09 컷오프",
    "FS10 대손회피",
    "FS11 특수관계자",
    "FS12 충당금",
    "FS13 손상회피",
    "FS14 유령직원",
]


def s8_a(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.35, 1.10))
    yy = head(sl, ax, y, aw, "금감원 지적사례 → 시나리오 14종")
    sk_bip(sl, ax, yy, aw, h - (yy - y), FS_L, FS_R, FS_LINK, lw_r=0.30)
    (ty, th), (by, bh) = rows(y, h, (1.20, 1.15))
    yy = head(sl, bx, ty, bw, "이상 전표 한 건을 만드는 법")
    sk_flow(sl, bx, yy, bw, th - (yy - ty), MAKE4, dark=(2,))
    yy = head(sl, bx, by, bw, "주입 규모")
    sk_area(sl, bx, yy, bw, bh - (yy - by), "주입 후 전체 113,465건", 0.0029, "330건 · 0.29%")


def s8_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (0.66, 1.75))
    yy = head(sl, x, ty, w, "이상 전표 한 건을 만드는 법")
    sk_band(sl, x, yy, w, 0.40, MAKE4, hot=(2, 3))
    (px, pw), (qx, qw) = cols(x, w, (1.35, 1.0))
    yy = head(sl, px, by, pw, "금감원 지적사례에서 뽑은 14종")
    sk_chips(sl, px, yy, pw, bh - (yy - by), FS14, ncol=3, hot=(0,))
    yy = head(sl, qx, by, qw, "가공매출 — 실제 연계 흐름")
    sk_flow(sl, qx, yy, qw, bh - (yy - by), CHAIN4, dark=(0,))


def s8_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.78, 1.35, 1.0))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "14종",
        "실제 사례에서 나왔다",
        "금감원 지적사례를 수법 14종으로 구조화해 정상 전표에 주입했다. 단건이 아니라 연계 흐름이다.",
    )
    (ty, th), (by, bh) = rows(y, h, (1.95, 0.55))
    yy = head(sl, bx, ty, bw, "이상 데이터 시나리오 14종")
    sk_chips(sl, bx, yy, bw, th - (yy - ty), FS14, ncol=2, hot=(0,))
    sk_tape(sl, bx, by, bw, bh, "주입 규모", ["113,465건", "이상 330건", "0.29%"], hot=(2,))
    yy = head(sl, cx, y, cw, "한 건을 만드는 4단계")
    sk_flow(sl, cx, yy, cw, h - (yy - y), MAKE4, dark=(2,))


def s8_d(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.0, 1.45))
    yy = head(sl, ax, y, aw, "가공매출 — 단건이 아니다")
    sk_flow(sl, ax, yy, aw, h - (yy - y), CHAIN4, dark=(0, 3))
    (ty, th), (by, bh) = rows(y, h, (1.65, 0.80))
    yy = head(sl, bx, ty, bw, "지적사례 흐름 → 시나리오 14종")
    sk_bip(sl, bx, yy, bw, th - (yy - ty), FS_L, FS_R, FS_LINK, lw_r=0.28)
    sk_tape(
        sl,
        bx,
        by,
        bw,
        bh,
        "주입 규모 — 쉽게 맞힐 문제로 만들지 않았다",
        ["전체 113,465건", "이상 330건", "0.29%"],
        hot=(2,),
    )


# ═════════════════════════════════════════════════════ S10 · 룰 29종
RULE_SRC2 = ["금감원 감리 지적사례 230건", "회계감사기준서 240호"]
TARGET9 = [
    ("추정계정", "147", 1.00, True),
    ("관계사", "80", 0.544, True),
    ("컷오프", "23", 0.156, False),
    ("비용자산화", "22", 0.150, False),
    ("기말결산", "20", 0.136, False),
    ("매출과대", "20", 0.136, False),
    ("이상고액", "12", 0.082, False),
    ("역분개", "3", 0.020, False),
    ("가수금 체류", "2", 0.014, False),
]
METHOD10 = [
    "수기",
    "승인한도초과",
    "자기승인",
    "승인생략",
    "유령승인자",
    "직무분리 겸직",
    "소급기표",
    "희소계정쌍",
    "휴일 입력",
    "심야 입력",
]
THEME6 = [
    "시점 (6)",
    "승인/권한 (5)",
    "금액/분할 (5)",
    "계정 성격 (5)",
    "거래 성격 (4)",
    "데이터 정합성 (4)",
]


def s10_a(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.95, 1.30, 0.95))
    yy = head(sl, ax, y, aw, "두 출처")
    sk_merge(sl, ax, yy, aw, h - (yy - y), RULE_SRC2, "룰 29종")
    (ty, th), (by, bh) = rows(y, h, (1.40, 0.95))
    yy = head(sl, bx, ty, bw, "감리 지적 빈도 — 실제 건수")
    sk_bars(sl, bx, yy, bw, th - (yy - ty), TARGET9, lab_w=1.00, val_w=0.30)
    yy = head(sl, bx, by, bw, "주제별 구성")
    sk_tape(
        sl,
        bx,
        yy,
        bw,
        bh - (yy - by),
        "",
        ["시점 6", "승인 5", "금액 5", "계정 5", "거래 4", "정합 4"],
        hot=(5,),
    )
    yy = head(sl, cx, y, cw, "조작 수법 10종 — 기준서 240호")
    sk_chips(sl, cx, yy, cw, h - (yy - y), METHOD10, ncol=2)


def s10_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (0.62, 1.75))
    yy = head(sl, x, ty, w, "출처가 두 축을 만든다")
    sk_cat(
        sl,
        x,
        yy,
        w,
        th - (yy - ty),
        [("감리 지적사례 230건", "도출", "적발 대상 9종 · 기준서 240호 → 수법 10종")],
        aw_r=0.28,
    )
    (px, pw), (qx, qw) = cols(x, w, (1.25, 1.0))
    yy = head(sl, px, by, pw, "적발 대상 9종 — 감리 지적 빈도")
    sk_bars(sl, px, yy, pw, bh - (yy - by), TARGET9, lab_w=1.10, val_w=0.32)
    yy = head(sl, qx, by, qw, "룰 29종 주제별")
    sk_flow(sl, qx, yy, qw, bh - (yy - by), THEME6)


def s10_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.75, 1.25, 1.15))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "29종",
        "룰마다 출전이 있다",
        "147·80·23은 임의 가중치가 아니라 감리에서 실제로 지적된 건수다.",
    )
    yy = head(sl, bx, y, bw, "주제 6 × 룰 29")
    sk_flow(sl, bx, yy, bw, h - (yy - y), THEME6)
    (ty, th), (by, bh) = rows(y, h, (1.45, 0.90))
    yy = head(sl, cx, ty, cw, "감리 지적 빈도")
    sk_bars(sl, cx, yy, cw, th - (yy - ty), TARGET9, lab_w=1.00, val_w=0.30)
    yy = head(sl, cx, by, cw, "조작 수법 10종")
    sk_chips(sl, cx, yy, cw, bh - (yy - by), METHOD10, ncol=3)


def s10_d(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (1.85, 0.60))
    (px, pw), (qx, qw) = cols(x, w, (1, 1))
    yy = head(sl, px, ty, pw, "감리 사례 230건 → 적발 대상 9종")
    sk_bars(sl, px, yy, pw, th - (yy - ty), TARGET9, lab_w=1.10, val_w=0.32)
    yy = head(sl, qx, ty, qw, "기준서 240호 → 조작 수법 10종")
    sk_chips(sl, qx, yy, qw, th - (yy - ty), METHOD10, ncol=2)
    sk_tape(
        sl,
        x,
        by,
        w,
        bh,
        "데이터 정합성 4종은 조합에서 분리 — 데이터 수정 신호",
        ["차대균형", "필수필드", "무효계정", "기간불일치"],
    )


# ═════════════════════════════════════════════════════ S11 · 검토 조합
COMBO_ROWS = [
    "추정계정",
    "관계사",
    "컷오프",
    "비용자산화",
    "기말결산",
    "매출과대",
    "이상고액",
    "역분개",
    "가수금",
]
COMBO_MARK = {
    (0, 0),
    (0, 2),
    (0, 5),
    (1, 0),
    (1, 3),
    (3, 0),
    (3, 6),
    (4, 1),
    (4, 2),
    (6, 4),
    (7, 8),
    (8, 9),
}
COMBO_COLT = "수기 · 한도초과 · 자기승인 · 승인생략 · 유령 · 겸직 · 소급 · 희소쌍 · 휴일 · 심야"
PRESET5 = [
    ("추정·관계사 조작", "근거", "감리 최다 147·80건"),
    ("수익인식", "근거", "기준서 240 문단 26"),
    ("역분개/은폐", "근거", "기말 되돌림·경과계정"),
    ("비용자산화", "근거", "감리 확정 22건"),
    ("결산 손상·충당금", "근거", "대손충당금 미인식 5건"),
]
FUNNEL_V = [
    ("전체 전표", "113,465건", 1.00),
    ("대상 — 비용자산화", "해당 전표만", 0.68),
    ("AND 조작 수법", "통제 위반", 0.40),
    ("검토 대상", "234건 · 10.7%", 0.24),
]
FUNNEL_H = [
    ("전체 전표", "113,465건", 1.00),
    ("대상", "비용자산화", 0.68),
    ("AND 수법", "통제 위반", 0.40),
    ("검토 대상", "234건 · 10.7%", 0.24),
]
DENSITY5 = [
    ("비용자산화", "10.7%", 1.00, True),
    ("역분개/은폐", "7.0%", 0.65, False),
    ("수익인식", "4.9%", 0.46, False),
    ("추정·관계사", "2.3%", 0.21, False),
    ("결산 손상·충당금", "0.5%", 0.05, False),
]


def s11_a(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.30, 1.15))
    yy = head(sl, ax, y, aw, "대상 9 × 수법 10 — 90칸을 감사인이 켠다")
    sk_grid(sl, ax, yy, aw, h - (yy - y), COMBO_ROWS, 10, COMBO_MARK, coltitle=COMBO_COLT)
    (ty, th), (by, bh) = rows(y, h, (1.15, 1.20))
    yy = head(sl, bx, ty, bw, "조합이 목록을 좁힌다")
    sk_funnel(sl, bx, yy, bw, th - (yy - ty), FUNNEL_H, horiz=True)
    yy = head(sl, bx, by, bw, "프리셋 5종 — 자주 쓰는 묶음일 뿐")
    sk_chips(sl, bx, yy, bw, bh - (yy - by), [p[0] for p in PRESET5], ncol=2)


def s11_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (1.72, 0.38))
    yy = head(sl, x, ty, w, "대상 9종 × 수법 10종 = 90칸, 감사인이 직접 켠다")
    sk_grid(
        sl, x, yy, w, th - (yy - ty), COMBO_ROWS, 10, COMBO_MARK, coltitle=COMBO_COLT, rlab_w=1.05
    )
    sk_tape(
        sl,
        x,
        by,
        w,
        bh,
        "프리셋 5종 — 근거는 전부 감리 실측",
        ["추정·관계사 147·80", "수익인식 240-26", "역분개/은폐", "비용자산화 22", "결산 손상 5"],
        hot=(0,),
    )


def s11_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (1.15, 0.85, 1.10))
    yy = head(sl, ax, y, aw, "대상 9 × 수법 10")
    sk_grid(
        sl, ax, yy, aw, h - (yy - y), COMBO_ROWS, 10, COMBO_MARK, coltitle=COMBO_COLT, rlab_w=0.92
    )
    yy = head(sl, bx, y, bw, "좁혀지는 목록")
    sk_funnel(sl, bx, yy, bw, h - (yy - y), FUNNEL_V)
    yy = head(sl, cx, y, cw, "실제 화면")
    sk_shot(
        sl,
        cx,
        yy,
        cw,
        h - (yy - y),
        IMG / "screenshot-combo-direct.png",
        "직접 조합 — 감사인이 대상·수법을 교차 선택",
    )


def s11_d(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.72, 1.05, 1.20))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "9×10",
        "권한이 감사인에게 있다",
        "시스템은 대상과 수법을 제시만 한다. 무엇을 교차할지, 얼마나 볼지는 사람이 정한다.",
    )
    yy = head(sl, bx, y, bw, "조합 = 대상 AND 수법")
    sk_bip(
        sl,
        bx,
        yy,
        bw,
        h - (yy - y),
        COMBO_ROWS,
        METHOD10,
        [(0, 0), (0, 2), (1, 0), (3, 5), (4, 6)],
        lw_r=0.42,
    )
    (ty, th), (by, bh) = rows(y, h, (1.45, 0.95))
    yy = head(sl, cx, ty, cw, "프리셋별 검토 대상 밀도")
    sk_bars(sl, cx, yy, cw, th - (yy - ty), DENSITY5, lab_w=1.15, val_w=0.34)
    yy = head(sl, cx, by, cw, "프리셋은 사전 묶음일 뿐")
    foot(
        sl,
        cx,
        yy,
        cw,
        bh - (yy - by),
        "감리 사례 스크립트 전수 집계로 도출했다. 감사인은 언제든 90칸에서 직접 교차한다.",
    )


# ═════════════════════════════════════════════════════ S12 · 분석적 검토
IDX5 = [
    ("벤포드 법칙", "계정·월", "첫 자릿수 이탈"),
    ("라운드 넘버 밀집도", "계정·작성자", "절사 비율 초과"),
    ("신규/희소 거래처", "거래처", "첫등장·희소·휴면"),
    ("계정별 활동성 변동", "전기대비", "금액·건수 변동"),
    ("월 집중도 변화", "전기대비", "월 비중 이동"),
]
UNIT_CMP = [
    ("보는 단위", "전표 1건씩", "계정·거래처·월로 묶음"),
    ("판정", "각 건은 조건 미해당", "분포가 이론값에서 이탈"),
]


def s12_a(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.0, 1.45))
    yy = head(sl, ax, y, aw, "같은 데이터, 다른 단위")
    sk_split(
        sl,
        ax,
        yy,
        aw,
        h - (yy - y),
        "같은 전표 데이터",
        ["전표 단위로 본다", "각 건은 조건 미해당", "이상 없음"],
        ["계정·거래처·월로 묶는다", "분포가 이론값 이탈", "검토 신호"],
    )
    (ty, th), (by, bh) = rows(y, h, (1.45, 0.95))
    yy = head(sl, bx, ty, bw, "지표 5종 — 무엇을 어떤 단위로")
    sk_cat(sl, bx, yy, bw, th - (yy - ty), IDX5, aw_r=0.36)
    yy = head(sl, bx, by, bw, "판정이 아니라 참고 신호")
    foot(
        sl,
        bx,
        yy,
        bw,
        bh - (yy - by),
        "회계감사기준 520호 분석적 절차 준용. 모집단 분포 원형을 제시하고, 임계 없이 비교 "
        "가능한 전 계정을 노출한다 — 위험 확정은 하지 않는다.",
    )


def s12_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (0.86, 1.50))
    yy = head(sl, x, ty, w, "지표 5종과 각 분석 단위")
    from sketch import tx

    n = len(IDX5)
    cwid = (w - (n - 1) * 0.10) / n
    for i, (nm, unit, crit) in enumerate(IDX5):
        cxx = x + i * (cwid + 0.10)
        add_box(sl, cxx, yy, cwid, 0.24, fill=C["primary"], line=None)
        tx(
            sl,
            cxx + 0.03,
            yy,
            cwid - 0.06,
            0.24,
            nm,
            6,
            C["bg"],
            bold=True,
            align=PP_ALIGN.CENTER,
            anc=MSO_ANCHOR.MIDDLE,
        )
        tx(
            sl,
            cxx + 0.03,
            yy + 0.27,
            cwid - 0.06,
            0.15,
            unit,
            6,
            C["accent"],
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        tx(
            sl,
            cxx + 0.03,
            yy + 0.43,
            cwid - 0.06,
            0.28,
            crit,
            6,
            C["muted"],
            align=PP_ALIGN.CENTER,
            ls=1.2,
        )
    (px, pw), (qx, qw) = cols(x, w, (1.0, 1.30))
    yy = head(sl, px, by, pw, "전표 단위로는 안 보인다")
    sk_cmp(sl, px, yy, pw, bh - (yy - by), UNIT_CMP, lt="전표 단위", rt="그룹 단위", axw_r=0.28)
    yy = head(sl, qx, by, qw, "실제 화면")
    sk_shot(
        sl,
        qx,
        yy,
        qw,
        bh - (yy - by),
        IMG / "screenshot-benford.png",
        "벤포드 첫째 자릿수 — 전체 적합, 계정별 이탈 상위 별도",
    )


def s12_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.72, 1.35, 1.05))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "5지표",
        "묶어야 드러난다",
        "룰은 전표 단위 판정이라 계정·거래처·기간으로 묶여야 보이는 이상치를 못 본다.",
    )
    yy = head(sl, bx, y, bw, "실제 화면")
    sk_shot(
        sl,
        bx,
        yy,
        bw,
        h - (yy - y),
        IMG / "screenshot-benford.png",
        "벤포드 — 전체 119,170건 적합(평균 차이 0.0020), 계정별 이탈 상위 9개 별도",
    )
    yy = head(sl, cx, y, cw, "지표 5종")
    sk_flow(sl, cx, yy, cw, h - (yy - y), [i[0] for i in IDX5])


def s12_d(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.30, 1.15))
    yy = head(sl, ax, y, aw, "지표 5종 → 분석 단위 3계열")
    sk_bip(
        sl,
        ax,
        yy,
        aw,
        h - (yy - y),
        [i[0] for i in IDX5],
        ["계정 · 월", "거래처", "전기 대비"],
        [(0, 0), (1, 0), (2, 1), (3, 2), (4, 2)],
        lw_r=0.44,
    )
    (ty, th), (by, bh) = rows(y, h, (1.45, 0.75))
    yy = head(sl, bx, ty, bw, "실제 화면")
    sk_shot(
        sl,
        bx,
        yy,
        bw,
        th - (yy - ty),
        IMG / "screenshot-round-density.png",
        "라운드 넘버 밀집 — 점선은 해당 데이터에서 산출한 기준선",
    )
    sk_tape(
        sl,
        bx,
        by,
        bw,
        bh,
        "확정이 아니라 참고 신호 (기준 520호 준용)",
        ["전표 단위", "그룹 집계", "분포 이탈", "검토 신호"],
        hot=(3,),
    )


# ═════════════════════════════════════════════════════ S13 · VAE 왜
UNSUP3 = ["정답 데이터 부재", "미정의 패턴 대응", "데이터 요건 충족"]
VAE4 = [
    ("근거 제시", "복구 오차를 항목별로 분해"),
    ("혼합형 처리", "금액과 계정과목을 한 척도로"),
    ("단순 암기 방지", "확률적 제약 → 공통 구조만"),
    ("판정 원리", "압축 후 복구, 오차가 점수"),
]
VPIPE6 = ["입력 변수 선정", "데이터 분할", "전처리", "학습", "이상치 점수화", "검증 2단계"]


def s13_a(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (1.05, 1.20))
    yy = head(sl, x, ty, w, "판정 원리 — 압축하고 복구해서 그 차이를 본다")
    sk_enc(sl, x + w * 0.14, yy, w * 0.72, th - (yy - ty))
    (px, pw), (qx, qw), (rx, rw) = cols(x, w, (1, 1, 1.15))
    yy = head(sl, px, by, pw, "왜 비지도인가")
    sk_flow(sl, px, yy, pw, bh - (yy - by), UNSUP3)
    yy = head(sl, qx, by, qw, "왜 VAE인가")
    sk_chips(sl, qx, yy, qw, bh - (yy - by), [v[0] for v in VAE4], ncol=2, hot=(0,))
    yy = head(sl, rx, by, rw, "학습부터 점수까지 6단계")
    sk_tape(
        sl,
        rx,
        yy,
        rw,
        bh - (yy - by),
        "",
        ["선정", "분할", "전처리", "학습", "점수", "검증"],
        hot=(3,),
    )


def s13_b(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.0, 1.45))
    yy = head(sl, ax, y, aw, "정답 라벨이 없다는 사실에서")
    sk_split(
        sl,
        ax,
        yy,
        aw,
        h - (yy - y),
        "부정 여부 정답 라벨 없음",
        ["지도 학습", "라벨의 수법만 학습", "신종·변형 놓침"],
        ["비지도 학습", "정상 분포만 학습", "이탈 정도로 후보"],
    )
    (ty, th), (by, bh) = rows(y, h, (1.10, 1.20))
    yy = head(sl, bx, ty, bw, "판정 원리")
    sk_enc(sl, bx, yy, bw, th - (yy - ty))
    yy = head(sl, bx, by, bw, "왜 VAE였나")
    sk_cat(sl, bx, yy, bw, bh - (yy - by), [(v[0], "→", v[1]) for v in VAE4], aw_r=0.30)


def s13_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.72, 1.20, 1.20))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "복구\n오차",
        "VAE는 근거를 낸다",
        "복구 오차가 항목별로 갈라져 나오므로 왜 이상인지 말할 수 있다. 별도 설명 도구가 필요 없다.",
    )
    yy = head(sl, bx, y, bw, "압축 → 복구 → 차이")
    sk_enc(sl, bx, yy, bw, h - (yy - y))
    yy = head(sl, cx, y, cw, "이 데이터가 요구한 것 → VAE의 답")
    sk_fan(sl, cx, yy, cw, h - (yy - y), "이 데이터가 요구한 것", [v[0] for v in VAE4], hot=(0,))


def s13_d(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (1.0, 1.20, 1.0))
    yy = head(sl, ax, y, aw, "학습부터 점수까지 6단계")
    sk_flow(sl, ax, yy, aw, h - (yy - y), VPIPE6, dark=(3,))
    yy = head(sl, bx, y, bw, "판정 원리")
    sk_enc(sl, bx, yy, bw, h - (yy - y))
    (ty, th), (by, bh) = rows(y, h, (1.15, 1.15))
    yy = head(sl, cx, ty, cw, "왜 비지도인가")
    sk_chips(sl, cx, yy, cw, th - (yy - ty), UNSUP3, ncol=1)
    yy = head(sl, cx, by, cw, "왜 VAE인가")
    sk_chips(sl, cx, yy, cw, bh - (yy - by), [v[0] for v in VAE4], ncol=2, hot=(0,))


# ═════════════════════════════════════════════════════ S14 · VAE 입력
DROP4 = ["합성 생성 흔적 14", "전표 고유값 7", "정답 라벨 4", "결측 과다 3 · 판정 결과 2"]
DERIV4 = [
    ("시계열 주기성 (4)", "→", "시각 거리 왜곡 방지"),
    ("업무 리드 타임 (3)", "→", "소급·지연 식별"),
    ("외부 통제 환경 (2)", "→", "승인 한도 조인"),
    ("금액 속성 구조화 (2)", "→", "소실 패턴 보존"),
]
PREP_L = ["계정·전표유형", "작성자·거래처", "금액 7종", "여부 컬럼", "환율·리드타임"]
PREP_R = ["원-핫 인코딩", "빈도 인코딩", "부호 보존 로그", "이진화 0/1", "적응형 스케일링"]
PREP_LINK = [(0, 0), (1, 1), (2, 2), (3, 3), (4, 4), (1, 0), (2, 4)]
GROUP39 = [
    ("금액 (7)", "차·대변·원화환산·송장·세액", False),
    ("계정 (6)", "계정과목·성격·보조계정", False),
    ("기간과 소속 (5)", "회계연도·기간·회사·통화", False),
    ("거래처와 부서 (4)", "거래처·상대 유형·부서", False),
    ("거래 성격 (4)", "전표유형·프로세스·입력경로", True),
    ("사람·적요·상태·증빙·세무 (13)", "작성자·승인자·적요·증빙", False),
]


def _sieve(sl, x, y, w, h, total, drops, keep):
    """체 — 걸러내는 동작이 형태로 보인다."""
    from sketch import arrow, fold, line_mid, line_soft, tail, tx, vline
    from sketch import box as _b

    th_, kh = 0.28, 0.30
    _b(sl, x, y, w, th_, total, fill=C["primary"], color=C["bg"], size=7.5)
    top = y + th_ + 0.07
    show, om, dh, pi, ty = fold(drops, top, y + h - kh - 0.16, min_h=0.16, cap=0.28)
    for i, t in enumerate(show):
        yy = top + i * pi
        tx(
            sl,
            x,
            yy,
            0.18,
            dh,
            "✕",
            7,
            C["muted"],
            bold=True,
            align=PP_ALIGN.CENTER,
            anc=MSO_ANCHOR.MIDDLE,
        )
        tx(sl, x + 0.22, yy, w - 0.22, dh, t, 6, C["muted"], anc=MSO_ANCHOR.MIDDLE)
    vline(sl, x + 0.09, y + th_, top + (len(show) - 1) * pi + dh / 2 - (y + th_), line_soft)
    tail(sl, x, ty, w, om)
    ky = y + h - kh
    arrow(sl, x + w / 2, ky - 0.14, x + w / 2, ky - 0.02, line_mid, 1.0)
    _b(sl, x, ky, w, kh, keep, fill=C["bg"], line=C["accent"], color=C["accent"], size=7.5, lw=1.3)


def s14_a(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.0, 1.45))
    yy = head(sl, ax, y, aw, "73열에서 무엇을 걸러냈나")
    _sieve(sl, ax, yy, aw, h - (yy - y), "원시 데이터 73개 열", DROP4, "최종 입력 50개")
    (ty, th), (by, bh) = rows(y, h, (1.25, 1.15))
    yy = head(sl, bx, ty, bw, "파생 변수 11개 — 무엇을 만들려고")
    sk_cat(sl, bx, yy, bw, th - (yy - ty), DERIV4, aw_r=0.40)
    yy = head(sl, bx, by, bw, "형태가 다른 데이터를 한 척도로")
    sk_bip(sl, bx, yy, bw, bh - (yy - by), PREP_L, PREP_R, PREP_LINK, lw_r=0.42)


def s14_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (0.72, 1.60))
    yy = head(sl, x, ty, w, "73 → 배제 30 → 39 + 파생 11 = 50")
    sk_funnel(
        sl,
        x,
        yy,
        w,
        th - (yy - ty),
        [
            ("원시 데이터", "73개 열", 1.0),
            ("학습 배제", "30개 열", 0.6),
            ("원본 채택", "39개 열", 0.5),
            ("최종 입력", "50개", 0.4),
        ],
        horiz=True,
    )
    (px, pw), (qx, qw) = cols(x, w, (1.15, 1.0))
    yy = head(sl, px, by, pw, "학습에 넣은 39개 묶음")
    sk_stack(sl, px, yy, pw, bh - (yy - by), GROUP39)
    yy = head(sl, qx, by, qw, "전처리 — 성격에서 기법으로")
    sk_bip(sl, qx, yy, qw, bh - (yy - by), PREP_L, PREP_R, PREP_LINK, lw_r=0.44)


def s14_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.75, 1.15, 1.25))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "73→50",
        "생성 흔적을 배제했다",
        "합성 흔적 14열과 정답 라벨 4열을 빼야 모델이 회계적 이상만 본다.",
    )
    yy = head(sl, bx, y, bw, "무엇을 왜 뺐나")
    sk_merge(sl, bx, yy, bw, h - (yy - y), DROP4[:3], "학습 배제 30열")
    (ty, th), (by, bh) = rows(y, h, (1.25, 1.15))
    yy = head(sl, cx, ty, cw, "파생 변수 11개의 목적")
    sk_cat(sl, cx, yy, cw, th - (yy - ty), DERIV4, aw_r=0.42)
    yy = head(sl, cx, by, cw, "학습 변수 39개 묶음")
    sk_stack(sl, cx, yy, cw, bh - (yy - by), GROUP39)


def s14_d(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (0.50, 1.80))
    sk_tape(
        sl,
        x,
        ty,
        w,
        th,
        "원시 73열 → 배제 30 → 원본 39 + 파생 11 = 최종 입력 50개",
        ["73 원시", "−30 배제", "39 원본", "+11 파생", "50 입력"],
        hot=(4,),
    )
    (px, pw), (qx, qw), (rx, rw) = cols(x, w, (1.05, 1.0, 1.05))
    yy = head(sl, px, by, pw, "학습 변수 39개 묶음")
    sk_stack(sl, px, yy, pw, bh - (yy - by), GROUP39)
    yy = head(sl, qx, by, qw, "파생 11개")
    sk_flow(sl, qx, yy, qw, bh - (yy - by), [d[0] for d in DERIV4])
    yy = head(sl, rx, by, rw, "전처리 5기법")
    sk_bip(sl, rx, yy, rw, bh - (yy - by), PREP_L, PREP_R, PREP_LINK, lw_r=0.44)


# ═════════════════════════════════════════════════════ S16 · 검증 ① 룰
LIFT8 = [
    ("L4-04 희소계정쌍", "810", 1.00, True),
    ("L2-05 역분개", "108", 0.62, True),
    ("L1-05 자기승인", "59", 0.52, False),
    ("L2-04 비용자산화", "40", 0.47, False),
    ("L3-09 가수금 체류", "22", 0.40, False),
    ("L3-02 수기", "2.3", 0.21, False),
    ("L1-07 승인생략", "0.64", 0.08, True),
    ("L1-04 외 5종", "0", 0.01, True),
]
RECOVER_V = [
    ("주입한 이상 전표", "330건", 1.00),
    ("최대 회수", "99.4~100%", 0.90),
    ("lift ≤ 1.0 배제", "변별력 상실", 0.62),
    ("실질 회수", "296~300건 · 90%", 0.44),
]
RECOVER_H = [
    ("주입", "330건", 1.00),
    ("최대 회수", "99.4%↑", 0.90),
    ("lift ≤ 1.0", "배제", 0.62),
    ("실질 회수", "90%", 0.44),
]
BASE6 = [
    ("L1-07 승인생략", "22.73%", "[15, 30] 통과"),
    ("L3-04 기말결산", "14.14%", "[10, 20] 통과"),
    ("L3-02 수기", "8.16%", "[5, 12] 통과"),
    ("L3-06 심야 입력", "3.25%", "[1.2, 8] 통과"),
    ("L1-06 직무분리", "3.05%", "[0.5, 8] 통과"),
    ("L3-10 추정계정", "0.42%", "[0.07, 0.7] 통과"),
]
TEST2 = ["정상 베이스라인 28 / 28 통과 · 실패 0", "단일 룰 주입 29 / 29 발현"]


def s16_a(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.35, 1.10))
    yy = head(sl, ax, y, aw, "lift — 이상 발현율 ÷ 정상 발현율")
    sk_bars(sl, ax, yy, aw, h - (yy - y), LIFT8, lab_w=1.15, val_w=0.34)
    (ty, th), (by, bh) = rows(y, h, (1.15, 1.20))
    yy = head(sl, bx, ty, bw, "회수했다는 말의 뜻")
    sk_funnel(sl, bx, yy, bw, th - (yy - ty), RECOVER_H, horiz=True)
    yy = head(sl, bx, by, bw, "두 시험")
    sk_flow(sl, bx, yy, bw, bh - (yy - by), TEST2, dark=(0, 1))


def s16_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (0.72, 1.65))
    yy = head(sl, x, ty, w, "회수했다는 말의 뜻 — 최대에서 실질로")
    sk_funnel(sl, x, yy, w, th - (yy - ty), RECOVER_H, horiz=True)
    yy = head(sl, x, by, w, "룰별 lift — 다 잡는 룰이 아니라 변별하는 룰")
    sk_bars(sl, x, yy, w, bh - (yy - by), LIFT8, lab_w=1.25, val_w=0.36)


def s16_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.72, 1.25, 1.20))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "810→0",
        "회수율보다 lift",
        "다 잡는 룰은 변별하지 않는다. 정상에서 더 자주 뜨는 룰은 실질 회수에서 뺐다.",
    )
    yy = head(sl, bx, y, bw, "룰별 lift")
    sk_bars(sl, bx, yy, bw, h - (yy - y), LIFT8, lab_w=1.15, val_w=0.34)
    (ty, th), (by, bh) = rows(y, h, (1.55, 0.80))
    yy = head(sl, cx, ty, cw, "정상 데이터 — 실측 발현율 : 사전 선언 임계")
    sk_cat(sl, cx, yy, cw, th - (yy - ty), BASE6, aw_r=0.40)
    sk_tape(sl, cx, by, cw, bh, "두 시험", ["베이스라인 28/28", "단일 룰 29/29"], hot=(0, 1))


def s16_d(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (1.0, 1.20, 1.0))
    yy = head(sl, ax, y, aw, "룰이 다 도는가")
    sk_merge(sl, ax, yy, aw, h - (yy - y), TEST2, "룰 29종 전부 실행·발현 확인")
    yy = head(sl, bx, y, bw, "룰별 lift")
    sk_bars(sl, bx, yy, bw, h - (yy - y), LIFT8, lab_w=1.15, val_w=0.34)
    yy = head(sl, cx, y, cw, "최대 회수 → 실질 회수")
    sk_funnel(sl, cx, yy, cw, h - (yy - y), RECOVER_V)


# ═════════════════════════════════════════════════════ S17 · 검증 ②
TRADE5 = [
    ("비용자산화", 0.04, 1.00, "234건 · 10.7%"),
    ("역분개/은폐", 0.16, 0.66, "1,558건 · 7.0%"),
    ("수익인식", 0.10, 0.47, "266건 · 4.9%"),
    ("추정·관계사", 0.26, 0.21, "2,040건 · 2.3%"),
    ("결산 손상·충당금", 0.92, 0.05, "50,200건 · 0.5%"),
]
CONTRIB6 = [
    ("인도일-증빙일 간격", "39.0", 1.00, True),
    ("대변 금액", "28.1", 0.72, False),
    ("차변 금액", "14.6", 0.37, False),
    ("전표유형 = KR", "11.1", 0.28, False),
    ("업무프로세스 = 구매", "8.5", 0.22, False),
    ("인도일 간격 결측", "7.9", 0.20, True),
]
NARROW_WIDE = [
    ("검토 대상", "234건", "약 50,200건"),
    ("밀도", "10.7%", "0.5%"),
    ("적중", "25건", "229~241건"),
]
PRESET_HIT = [
    ("비용자산화", "적중 25", "234건 · 10.7%"),
    ("역분개/은폐", "108~111", "1,558건 · 7.0%"),
    ("수익인식", "12~14", "266건 · 4.9%"),
    ("추정·관계사", "46~47", "2,040건 · 2.3%"),
    ("결산 손상·충당금", "229~241", "50,200건 · 0.5%"),
]


def s17_a(sl, x, y, w, h):
    (ax, aw), (bx, bw) = cols(x, w, (1.20, 1.25))
    yy = head(sl, ax, y, aw, "조합이 검토량과 밀도를 맞바꾼다")
    sk_trade(sl, ax, yy, aw, h - (yy - y), TRADE5, xlab="검토 대상 건수 →", ylab="밀 도")
    (ty, th), (by, bh) = rows(y, h, (1.0, 1.35))
    yy = head(sl, bx, ty, bw, "VAE 판별력 (가공매출 계열 한정)")
    sk_gauge(sl, bx, yy, bw, th - (yy - ty), 0.8845, 0.5, "AUROC 0.8845")
    yy = head(sl, bx, by, bw, "무엇을 보고 판정했나 — 항목별 기여 배율")
    sk_bars(sl, bx, yy, bw, bh - (yy - by), CONTRIB6, lab_w=1.15, val_w=0.32)


def s17_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (1.20, 1.15))
    yy = head(sl, x, ty, w, "프리셋별 검토 대상 밀도 — 감사인이 고르는 지점")
    sk_bars(
        sl,
        x,
        yy,
        w,
        th - (yy - ty),
        [
            ("비용자산화 · 234건", "10.7%", 1.00, True),
            ("역분개/은폐 · 1,558건", "7.0%", 0.65, False),
            ("수익인식 · 266건", "4.9%", 0.46, False),
            ("추정·관계사 · 2,040건", "2.3%", 0.21, False),
            ("결산 손상·충당금 · 50,200건", "0.5%", 0.05, False),
        ],
        lab_w=1.85,
        val_w=0.40,
    )
    (px, pw), (qx, qw) = cols(x, w, (1.0, 1.30))
    yy = head(sl, px, by, pw, "좁게 볼지 넓게 볼지")
    sk_cmp(
        sl,
        px,
        yy,
        pw,
        bh - (yy - by),
        NARROW_WIDE,
        lt="좁게 — 비용자산화",
        rt="넓게 — 결산 손상",
        axw_r=0.30,
    )
    yy = head(sl, qx, by, qw, "VAE 판별력과 그 근거 (가공매출 한정)")
    (g1, gw1), (g2, gw2) = cols(qx, qw, (1.0, 1.15))
    sk_gauge(sl, g1, yy, gw1, bh - (yy - by), 0.8845, 0.5, "0.8845")
    sk_bars(sl, g2, yy, gw2, bh - (yy - by), CONTRIB6[:4], lab_w=1.05, val_w=0.30)


def s17_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.72, 1.30, 1.10))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "0.2~44%",
        "검토량을 감사인이 조절한다",
        "조합에 따라 234건에서 50,200건까지 갈린다. 시스템은 그 지점을 정하지 않는다.",
    )
    yy = head(sl, bx, y, bw, "정밀도와 분량의 맞바꿈")
    sk_trade(sl, bx, yy, bw, h - (yy - y), TRADE5, xlab="검토 대상 건수 →", ylab="밀 도")
    (ty, th), (by, bh) = rows(y, h, (1.0, 1.35))
    yy = head(sl, cx, ty, cw, "VAE 판별력 (가공매출 한정)")
    sk_gauge(sl, cx, yy, cw, th - (yy - ty), 0.8845, 0.5, "AUROC 0.8845")
    yy = head(sl, cx, by, cw, "항목별 기여 배율")
    sk_bars(sl, cx, yy, cw, bh - (yy - by), CONTRIB6, lab_w=1.10, val_w=0.32)


def s17_d(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (1.0, 1.15, 1.05))
    yy = head(sl, ax, y, aw, "좁게 / 넓게")
    sk_split(
        sl,
        ax,
        yy,
        aw,
        h - (yy - y),
        "전표 약 11만 건",
        ["비용자산화 조합", "검토 234건", "적중 25 · 10.7%"],
        ["결산 손상·충당금", "검토 50,200건", "적중 229~241 · 0.5%"],
        ma="좁게",
        mb="넓게",
    )
    yy = head(sl, bx, y, bw, "프리셋 5종 — 적중 : 검토 대상")
    sk_cat(sl, bx, yy, bw, h - (yy - y), PRESET_HIT, aw_r=0.36)
    (ty, th), (by, bh) = rows(y, h, (1.0, 1.35))
    yy = head(sl, cx, ty, cw, "VAE 판별력 (가공매출 한정)")
    sk_gauge(sl, cx, yy, cw, th - (yy - ty), 0.8845, 0.5, "AUROC 0.8845")
    yy = head(sl, cx, by, cw, "기여 배율")
    sk_bars(sl, cx, yy, cw, bh - (yy - by), CONTRIB6, lab_w=1.10, val_w=0.32)


# ═════════════════════════════════════════════════════ S19 · 마무리
EXCL4 = [
    ("관계 분석 · 순환거래", "사유", "단일 법인 전표로는 불가"),
    ("다중 모델 앙상블", "사유", "주입 규칙만 외우는 순환 학습"),
    ("LLM 서술 생성", "사유", "외부 API에 전표 전송 금지"),
    ("자동 위험 등급", "사유", "가중치 배점이 임의적"),
]
GRADE3 = [
    ("산출 방식", "등급 자동 부여", "시스템 등급화 폐지"),
    ("산출 근거", "주관적 가중치 배점", "감리 건수·기준서 원문"),
    ("판단 주체", "시스템이 위험 선언", "감사인이 최종 판단"),
]
LIMIT_A = ["합성 데이터 기반 검증", "정답 데이터 자체 주입", "판정 기준 내부 산출"]
LIMIT_B = ["단일 법인 단일 소스", "치밀한 조작 검증 불가", "VAE 탐지 편향성"]
STACK9 = [
    ("합성데이터", "Rust · 15 크레이트", False),
    ("언어/패키지", "Python 3.11+ · uv", False),
    ("수집/전처리", "pandas · openpyxl", False),
    ("데이터 검증", "Pandera 게이트", False),
    ("통계", "scipy.stats · numpy", False),
    ("비지도 학습", "PyTorch VAE", True),
    ("저장", "DuckDB 파일 격리", False),
    ("설정", "pydantic + YAML", False),
    ("화면", "Streamlit + Plotly", False),
]


def s19_a(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (1.15, 1.05, 1.05))
    (ty, th), (by, bh) = rows(y, h, (1.45, 0.90))
    yy = head(sl, ax, ty, aw, "무엇을 왜 뺐나")
    sk_cat(sl, ax, yy, aw, th - (yy - ty), EXCL4, aw_r=0.40)
    sk_tape(
        sl,
        ax,
        by,
        aw,
        bh,
        "구현 못한 게 아니라 근거가 서지 않아 안 하기로 한 것",
        ["데이터 한계", "순환 학습", "보안 제약"],
    )
    yy = head(sl, bx, y, bw, "위험 등급 폐지 — 판단 주체가 옮겨간다")
    sk_cmp(sl, bx, yy, bw, h - (yy - y), GRADE3, lt="변경 전", rt="변경 후", axw_r=0.30)
    yy = head(sl, cx, y, cw, "한계 6가지")
    sk_split(
        sl, cx, yy, cw, h - (yy - y), "한계 6가지", LIMIT_A, LIMIT_B, ma="데이터로", mb="원리적"
    )


def s19_b(sl, x, y, w, h):
    (ty, th), (by, bh) = rows(y, h, (1.0, 1.0))
    (px, pw), (qx, qw) = cols(x, w, (1.15, 1.0))
    yy = head(sl, px, ty, pw, "무엇을 왜 뺐나")
    sk_cat(sl, px, yy, pw, th - (yy - ty), EXCL4, aw_r=0.40)
    yy = head(sl, qx, ty, qw, "위험 등급 폐지")
    sk_cmp(sl, qx, yy, qw, th - (yy - ty), GRADE3, lt="변경 전", rt="변경 후", axw_r=0.30)
    (rx, rw), (sx2, sw) = cols(x, w, (1.0, 1.15))
    yy = head(sl, rx, by, rw, "한계 6 — 회색이 데이터로 해소 / 테두리가 원리적")
    sk_chips(
        sl,
        rx,
        yy,
        rw,
        bh - (yy - by),
        [LIMIT_A[0], LIMIT_B[0], LIMIT_A[1], LIMIT_B[1], LIMIT_A[2], LIMIT_B[2]],
        ncol=2,
        hot=(1, 3, 5),
    )
    yy = head(sl, sx2, by, sw, "기술 스택 9계층 — 전 구간 로컬")
    sk_chips(
        sl,
        sx2,
        yy,
        sw,
        bh - (yy - by),
        ["Rust", "Python", "pandas", "Pandera", "scipy", "PyTorch", "DuckDB", "YAML", "Streamlit"],
        ncol=3,
        hot=(5,),
    )


def s19_c(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (0.72, 1.30, 1.05))
    sk_text(
        sl,
        ax,
        y,
        aw,
        h,
        "폐지",
        "빼는 것도 근거로 뺀다",
        "구현하지 못한 것이 아니라 근거가 서지 않아 하지 않기로 한 것이다.",
    )
    (ty, th), (by, bh) = rows(y, h, (1.10, 1.25))
    yy = head(sl, bx, ty, bw, "제외 대상과 사유")
    sk_cat(sl, bx, yy, bw, th - (yy - ty), EXCL4, aw_r=0.40)
    yy = head(sl, bx, by, bw, "한계 6가지 — 두 계열")
    sk_cmp(
        sl,
        bx,
        yy,
        bw,
        bh - (yy - by),
        [
            ("검증", LIMIT_A[0], LIMIT_B[0]),
            ("정답", LIMIT_A[1], LIMIT_B[1]),
            ("기준", LIMIT_A[2], LIMIT_B[2]),
        ],
        lt="데이터로 해소",
        rt="원리적",
        axw_r=0.22,
    )
    yy = head(sl, cx, y, cw, "기술 스택 9계층")
    sk_stack(sl, cx, yy, cw, h - (yy - y), STACK9)


def s19_d(sl, x, y, w, h):
    (ax, aw), (bx, bw), (cx, cw) = cols(x, w, (1.05, 1.10, 1.05))
    yy = head(sl, ax, y, aw, "네 건이 한 결론으로")
    sk_merge(sl, ax, yy, aw, h - (yy - y), [e[0] for e in EXCL4], "근거가 서지 않으면 넣지 않는다")
    (ty, th), (by, bh) = rows(y, h, (1.15, 1.20))
    yy = head(sl, bx, ty, bw, "위험 등급 폐지 전/후")
    sk_cmp(sl, bx, yy, bw, th - (yy - ty), GRADE3, lt="변경 전", rt="변경 후", axw_r=0.30)
    yy = head(sl, bx, by, bw, "한계 6가지 — 두 계열")
    sk_cmp(
        sl,
        bx,
        yy,
        bw,
        bh - (yy - by),
        [
            ("검증", LIMIT_A[0], LIMIT_B[0]),
            ("정답", LIMIT_A[1], LIMIT_B[1]),
            ("기준", LIMIT_A[2], LIMIT_B[2]),
        ],
        lt="데이터로",
        rt="원리적",
        axw_r=0.22,
    )
    yy = head(sl, cx, y, cw, "기술 스택 9계층")
    sk_stack(sl, cx, yy, cw, h - (yy - y), STACK9)


# ═════════════════════════════════════════════════════ 슬라이드 목록
SLIDES = [
    dict(
        no="S4",
        part="Ⅰ",
        title="문제 인식 — 표본으로는 무엇이 안 되는가",
        focus="초점(채록): 표본 구조적 한계 = 감사인 미대체 = 3모델 분담 = 실무 근거 — 네 개 동일 비중",
        fixed="담기는 것(고정): README §2 비교 5축 · §3 기대효과 3 · 감사인 경계 · 근거 실무 문서",
        c=[
            ("좌앵커 + 중앙 5축 대결 + 우 기대효과", s4_a),
            ("좌 5축 대결 + 우 두 경로 분기 (둘 다 전체 높이)", s4_b),
            ("상단 5축 가로 3단(축·표본·전수) + 하단 기대효과 / 경계", s4_c),
            ("좌 두 경로 분기 + 우 5축 대결 + 하단 근거 띠", s4_d),
        ],
    ),
    dict(
        no="S5",
        part="Ⅰ",
        title="시스템 한 장 요약 — 무엇이 들어가 무엇이 나오는가",
        focus="초점(채록): 감사인을 대체하지 않는다 · 세 모델이 각기 다른 질문에 답한다",
        fixed="담기는 것(고정): 파이프라인 5단 · 3모델 질문·산출물 · 감사인 경계 · 원장 규모",
        c=[
            ("좌 세로 파이프라인 + 우 3종에서 오른쪽 가지 뻗어 설명", s5_a),
            ("상단 가로 파이프라인 밴드 + 하단 3종 가지 + 경계", s5_b),
            ("좌앵커 + 파이프라인 + 우 3종 팬아웃 + 경계 카탈로그", s5_c),
            ("좌 3종 가지 크게 + 우상 경계 + 우하 파이프라인 띠 (좌우 반전)", s5_d),
        ],
    ),
    dict(
        no="S7",
        part="Ⅱ",
        title="정상 데이터 생성 — 없는 전표를 어떻게 만들었나",
        focus="초점(채록): 실데이터 부재로 직접 생성 = 회계 무결성 유지 (동급 2위)",
        fixed="담기는 것(고정): 생성 8단계 · 무결성 관문 3 + 폐기·재생성 · 시나리오 20종/흐름 7 · EY→K-IFRS · 355,786행",
        c=[
            ("좌앵커 + 중앙 8단계 루프(관문·되돌림) + 우 20종/산출", s7_a),
            ("상단 8단계 가로 밴드 + 하단 관문 팬인 / 20종 막대", s7_b),
            ("좌 8단계 세로 + 우상 흐름→20종 팬 + 우하 관문 카탈로그", s7_c),
            ("좌 시작(EY) + 중앙 생성 루프 + 우 20종/산출", s7_d),
        ],
    ),
    dict(
        no="S8",
        part="Ⅱ",
        title="이상 데이터 주입 — 14종은 어디서 왔나",
        focus="초점(채록): 이상 14종이 실제 사례에서 나왔다 (최우선) · 희소성 0.29%는 후순위",
        fixed="담기는 것(고정): 감리 사례→FS01~14 흐름별 · 생성 4단계 · 가공매출 연계 체인 · 330건/113,465건",
        c=[
            ("좌 흐름→14종 매핑 + 우상 4단계 + 우하 0.29% 면적", s8_a),
            ("상단 4단계 가로 밴드 + 하단 14종 칩 격자 / 가공매출 체인", s8_b),
            ("좌앵커 + 중앙 14종 칩 + 하단 희소성 띠 + 우 4단계", s8_c),
            ("좌 가공매출 체인 크게 + 우 흐름→14종 매핑 + 하단 띠", s8_d),
        ],
    ),
    dict(
        no="S10",
        part="Ⅲ",
        title="룰 29종 — 감으로 만들지 않았다",
        focus="초점(채록): 룰 하나하나에 출전이 있다",
        fixed="담기는 것(고정): 감리 230건→대상 9종(빈도) · 기준서 240호→수법 10종 · 주제 6 · 정합성 4 분리",
        c=[
            ("좌 두 출처 팬인 + 중앙 빈도 막대 + 우 수법 10칩", s10_a),
            ("상단 출처→축 카탈로그 + 하단 빈도 막대 / 주제 6", s10_b),
            ("좌앵커 + 중앙 주제 6 세로 + 우 빈도 막대 / 수법 칩", s10_c),
            ("좌 빈도 막대 + 우 수법 10칩 + 하단 정합성 4 분리 띠", s10_d),
        ],
    ),
    dict(
        no="S11",
        part="Ⅲ",
        title="검토 조합 — 무엇을 볼지는 감사인이 정한다",
        focus="초점(채록): 조합 권한이 감사인에게 있다",
        fixed="담기는 것(고정): 대상 9 × 수법 10 = 90칸 · AND 논리 · 프리셋 5 + 근거 · 실제 화면",
        c=[
            ("좌 9×10 그리드 + 우상 AND 깔때기 + 우하 프리셋 칩", s11_a),
            ("상단 9×10 그리드 크게 + 하단 프리셋 띠", s11_b),
            ("좌 그리드 + 중앙 깔때기 + 우 실물 화면", s11_c),
            ("좌앵커 + 중앙 대상─AND─수법 매핑 + 우 프리셋 밀도 막대", s11_d),
        ],
    ),
    dict(
        no="S12",
        part="Ⅲ",
        title="분석적 검토 — 전표 하나로는 안 보이는 것",
        focus="초점: 전표 단위로는 안 보이고 그룹(계정·거래처·월)으로 표면화된다 · 확정 아닌 참고 신호",
        fixed="담기는 것(고정): 전표 vs 그룹 대비 · 지표 5종(분석 단위·판정 대상) · 520호 준용 · 실물 화면",
        c=[
            ("좌 전표/그룹 분기 + 우 지표 5 카탈로그 + 참고신호", s12_a),
            ("상단 지표 5 가로 3단 + 하단 전표/그룹 대결 + 화면", s12_b),
            ("좌앵커 + 중앙 실물 화면 크게 + 우 지표 5 세로", s12_c),
            ("좌 지표→분석단위 매핑 + 우상 화면 + 우하 신호 띠", s12_d),
        ],
    ),
    dict(
        no="S13",
        part="Ⅲ",
        title="VAE — 왜 이 모델이어야 했나",
        focus="초점(채록): VAE는 근거를 낸다",
        fixed="담기는 것(고정): 압축→복구→오차 · 비지도 채택 3사유 · VAE 채택 4사유 · 파이프라인 6단계",
        c=[
            ("상단 압축·복구 도해 풀폭 + 하단 3칼럼(비지도/VAE/6단계)", s13_a),
            ("좌 지도·비지도 분기 + 우상 압축·복구 + 우하 4사유", s13_b),
            ("좌앵커 + 중앙 압축·복구 크게 + 우 요구조건→답 팬", s13_c),
            ("좌 6단계 세로 + 중앙 압축·복구 + 우 채택사유 칩", s13_d),
        ],
    ),
    dict(
        no="S14",
        part="Ⅲ",
        title="VAE 입력 설계 — 73열 중 50개만 넣었다",
        focus="초점: 생성 흔적을 배제했다 — S13 '근거를 낸다'의 뒷받침",
        fixed="담기는 것(고정): 73→배제 30(사유)→39+파생 11=50 · 파생 4종 목적 · 전처리 5기법 · 39개 묶음",
        c=[
            ("좌 73→50 체 + 우상 파생 카탈로그 + 우하 전처리 매핑", s14_a),
            ("상단 73→50 가로 깔때기 + 하단 39묶음 스택 / 전처리 매핑", s14_b),
            ("좌앵커 + 중앙 배제 팬인 + 우 파생 / 39묶음", s14_c),
            ("상단 73→50 띠 + 하단 39묶음 / 파생 11 / 전처리 3칼럼", s14_d),
        ],
    ),
    dict(
        no="S16",
        part="Ⅳ",
        title="검증 ① 룰 — 다 잡는 룰과 변별하는 룰",
        focus="초점(채록): 회수율보다 lift (2순위)",
        fixed="담기는 것(고정): lift 810~0 · 최대/실질 회수 · 베이스라인 28/0 실측:임계 · 단일룰 29/29",
        c=[
            ("좌 lift 막대 + 우상 회수 깔때기 + 우하 두 시험", s16_a),
            ("상단 회수 깔때기 가로 + 하단 lift 막대 풀폭", s16_b),
            ("좌앵커 + 중앙 lift 막대 + 우 베이스라인 카탈로그", s16_c),
            ("좌 두 시험 팬인 + 중앙 lift 막대 + 우 회수 깔때기", s16_d),
        ],
    ),
    dict(
        no="S17",
        part="Ⅳ",
        title="검증 ② 검토량은 감사인이 조절한다",
        focus="초점(채록): 검토량을 감사인이 조절한다 (최우선)",
        fixed="담기는 것(고정): 프리셋 5 적중·검토대상·밀도 · AUROC 0.8845(무작위 0.5) · 기여도 39배 · 가공매출 한정",
        c=[
            ("좌 트레이드오프 곡선 + 우상 게이지 + 우하 기여도 막대", s17_a),
            ("상단 프리셋 밀도 막대 + 하단 좁게/넓게 대결 + VAE 2단", s17_b),
            ("좌앵커 + 중앙 트레이드오프 크게 + 우 게이지/기여도", s17_c),
            ("좌 좁게/넓게 분기 + 중앙 프리셋 카탈로그 + 우 게이지/기여도", s17_d),
        ],
    ),
    dict(
        no="S19",
        part="Ⅴ",
        title="빼는 것도 근거로 뺀다",
        focus="초점(채록): 빼는 것도 근거로 뺀다",
        fixed="담기는 것(고정): 제외 4건:사유 · 등급 폐지 전/후 3축 · 한계 6(2계열) · 스택 9계층",
        c=[
            ("좌 제외 카탈로그 + 중앙 전/후 대결 + 우 한계 2계열", s19_a),
            ("2×2 구획 — 제외 / 전·후 / 한계 칩 / 스택 칩", s19_b),
            ("좌앵커 + 중앙 제외·한계 2단 + 우 스택 9계층", s19_c),
            ("좌 제외 4 팬인 + 중앙 전·후 / 한계 대결 + 우 스택 9계층", s19_d),
        ],
    ),
]


def build():
    from sketch import tx

    prs = new_presentation()
    blank = prs.slide_layouts[6]
    for sd in SLIDES:
        sl = prs.slides.add_slide(blank)
        compact_header(sl, f"{sd['no']} · PART {sd['part']} · 형태 후보 4 (내용 동일)", sd["title"])
        tx(
            sl,
            G.MARGIN_L,
            0.72,
            G.RIGHT_EDGE - G.MARGIN_L,
            0.16,
            sd["focus"],
            7.5,
            C["accent"],
            bold=True,
        )
        tx(sl, G.MARGIN_L, 0.88, G.RIGHT_EDGE - G.MARGIN_L, 0.16, sd["fixed"], 7, C["muted"])
        for i, (form, draw) in enumerate(sd["c"]):
            mx, my = MX[i % 2], MY[i // 2]
            add_box(sl, mx, my, MW, MH, fill=None, line=line_soft, line_w=0.75)
            add_box(sl, mx, my, 0.22, 0.22, fill=C["accent"], line=None)
            tx(
                sl,
                mx,
                my,
                0.22,
                0.22,
                str(i + 1),
                9,
                C["bg"],
                bold=True,
                align=PP_ALIGN.CENTER,
                anc=MSO_ANCHOR.MIDDLE,
            )
            tx(
                sl,
                mx + 0.30,
                my,
                MW - 0.36,
                0.24,
                form,
                8,
                C["primary"],
                bold=True,
                anc=MSO_ANCHOR.MIDDLE,
            )
            hrule(sl, mx + PAD, my + LAB - 0.03, MW - 2 * PAD, color=line_soft, weight=0.75)
            draw(sl, mx + PAD, my + LAB + 0.10, MW - 2 * PAD, MH - LAB - 0.22)
        source_line(
            sl,
            "후보 4개는 전부 같은 내용이다 — 다른 것은 배치·인코딩뿐 · 라벨은 "
            "자리에 맞게 접혔다(… 외 N개) · 번호로 회신 (예: S4=1 · S7=3 · S13=전탈락)",
        )
    out = HERE / "gallery.pptx"
    prs.save(str(out))
    print(f"저장: {out} — {len(SLIDES)}장")


if __name__ == "__main__":
    build()
