"""S12 구조화 출력 — 고밀도 개편본 (s06 공식, 2026-07-20). 골든 편입 대상.

원본 s12_variants.variant_b(28pt 헤더 + 3서사 + 실물 답변 스크린샷(대형) + 스키마 카드 +
검증 flow + 검은 바)를 승리 공식으로 재배치:
- 압축 헤더(dense.compact_header)
- 상단 시그니처 밴드: 출력 스키마 ClarifyOutput 카드(주인공, 좌) + 형식 검증 flow(우)
- 하단: s06 표준 hero_card 4장(자유형 폐기 / 같은 칸 출력 / 인용 강제 / 화면 미도달)
- 실물 답변 스크린샷(대형 텍스트 블록)은 밀도 포맷 부적합 → 핵심을 카드에 흡수
- 검은 전폭 바 폐지 → 결론은 카드에 흡수 (하단 한줄평 바 금지, design-rules §8)

콘텐츠는 s12_variants.VARIANT_B_DEFAULTS 원문 그대로.
"""

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR

from . import dense as D
from . import grid as G
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, fit_picture, load_kit

K = load_kit()

KICKER = "3. 기술 설명 — TECH 04 · Generate"
SOURCE = "출처: app/agents.py (ClarifyOutput · output_validator) · 5_INTERFACE.md · 7_JOURNEY.md"

NARRATIVES = [
    (
        "왜 필요한가",
        "같은 질문에 매번 다른 형식으로 답하면 정량 평가가 불가능하다. "
        "자유형 텍스트를 폐기하고 답변의 칸을 고정해야 한다.",
    ),
    (
        "파이프라인에서의 역할",
        "Generate의 출력 계약. 모든 답변이 결론·인용·후속질문의 같은 칸으로 "
        "나오므로 화면 배치와 채점이 기계적으로 가능해진다.",
    ),
    (
        "어떻게 만들었나",
        "답변 형식을 코드(스키마)로 선언하고, 검증기가 빈 인용·빈 분기를 "
        "거부해 자동 재시도 — 형식 미달 답변은 화면에 도달하지 못한다.",
    ),
]

# content_contract 계약 키 — golden.tech_capture의 필수 키 출처.
# 2026-07-26 `s12_variants`(sparse 시안)에서 이관. 골든 전용 글이다.
IMG_CROP = str(
    Path(__file__).resolve().parents[5] / "golden" / "ref" / "s12_answer_crop.png"
)  # 508×382 발췌 — 비율은 fit_picture가 파일에서 읽는다(주석은 참고용, 계산에 안 씀)

VARIANT_B_DEFAULTS = {
    "kicker": KICKER,
    "headline": "구조화 출력 — 답변의 형식을 코드가 강제한다",
    "narratives": NARRATIVES,
    "mid_subhead": "실물 답변 화면 — 발췌",
    "caption": "전체 답변 중 [확인 질문]·[조건부 결론 Case 1·2] 발췌 — 본인 vs 대리인 질의",
    "schema_subhead": "출력 스키마 — ClarifyOutput",
    "glosses": [
        ("selected_branches", "고른 결론 분기"),
        ("answer", "답변 본문 (마크다운)"),
        ("cited_paragraphs", "인용 문단 — 비면 거부"),
        ("cited_cases", "인용 질의회신·감리 ID"),
        ("cited_ie", "인용 적용사례 번호"),
        ("follow_up_questions", "확인 질문 3개"),
        ("is_conclusion", "결론 포함 — 항상 True"),
    ],
    "validate_subhead": "형식 검증 — 거부와 자동 재시도",
    "flow_generate": "답변 생성",
    "diamond": "스키마\n검사",
    "flow_display": "화면 표시",
    "reject_note": "거부 — 자동 재시도: 인용 0개 · 분기 0개 · 결론 없는 단정",
    "bar": "형식을 지키지 못한 답변은 화면에 도달하지 못한다 — 스키마 선언 + 검증기 자동 재시도",
    "source": SOURCE,
}
C, S, F = K["rgb"], K["sizes"], K["fonts"]

# 카드 주제 = 구조화 출력의 고유 성질 (제네릭 왜/역할/구축 대신)
CARDS = [
    (
        "1",
        "자유형 폐기",
        "정량 평가 불가",
        "file-text",
        "답변의 칸을 고정한다",
        [
            ("같은 질문 매번 다른 형식", " 정량 평가 불가"),
            ("자유형 텍스트 폐기", " 답변 칸을 고정"),
            ("형식 통제", " 화면·채점의 전제"),
        ],
        [
            ("file-text", "자유형", " 폐기"),
            ("circle-check-big", "칸", " 고정"),
        ],
    ),
    (
        "2",
        "같은 칸 출력",
        "Generate 출력 계약",
        "lock",
        "결론·인용·후속질문 한 칸",
        [
            ("결론·인용·후속질문", " 같은 칸으로 출력"),
            ("모든 답변 동형", " 화면 배치 기계화"),
            ("정형 출력", " 채점 자동 가능"),
        ],
        [
            ("lock", "출력", " 계약"),
            ("circle-check-big", "채점", " 자동"),
        ],
    ),
    (
        "3",
        "인용 강제",
        "빈 인용·빈 분기 거부",
        "shield-check",
        "cited 비면 검증기가 거부",
        [
            ("cited_paragraphs 비면", " 검증기가 거부"),
            ("빈 분기·결론 없는 단정", " 함께 거부"),
            ("인용 없는 주장", " 화면 밖으로"),
        ],
        [
            ("circle-x", "빈 인용", " 거부"),
            ("triangle-alert", "단정", " 차단"),
        ],
    ),
    (
        "4",
        "화면 미도달",
        "미달은 자동 재시도",
        "circle-check-big",
        "통과분만 화면에 도달",
        [
            ("형식 미달 답변", " 화면 도달 못 함"),
            ("result_validator", " 자동 재시도"),
            ("통과분만", " 사용자에게 노출"),
        ],
        [
            ("gauge", "재시도", " 자동"),
            ("circle-x", "미달", " 0"),
        ],
    ),
]


def _subhead(slide, x, y, w, text):
    add_text(slide, x, y, w, 0.28, text, S["caption"], F["head"], C["primary"], bold=True)
    D.hrule(slide, x, y + 0.30, w, color=D.line_mid, weight=1.0)


def _schema_card(slide, c, x0, x1, y_sub):
    """출력 스키마 ClarifyOutput — 다크 코드 카드. 7 필드 + 짧은 설명(glosses). y_sub=소제목 y."""
    tw = x1 - x0
    _subhead(slide, x0, y_sub, tw, c["schema_subhead"])
    cy = y_sub + 0.44
    card = add_box(slide, x0, cy, tw, 2.02, fill=C["primary"], shape="round")
    try:
        card.adjustments[0] = 0.035
    except (IndexError, ValueError):
        pass
    for i, (field, gloss) in enumerate(c["glosses"]):
        add_text(
            slide,
            x0 + 0.28,
            cy + 0.16 + i * 0.258,
            tw - 0.5,
            0.24,
            [
                [
                    (field, {"bold": True, "color": C["accent"], "font": "Consolas"}),
                    ("   " + gloss, {"color": C["bg_alt"]}),
                ]
            ],
            S["foot"],
            F["body"],
            C["bg_alt"],
            anchor=MSO_ANCHOR.MIDDLE,
        )


def build(prs, c=None):
    c = {**VARIANT_B_DEFAULTS, **(c or {})}
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    D.compact_header(slide, c["kicker"], c["headline"])

    # ══ 2존: 좌 = 실물 답변(크게) + 스키마 / 우 = hero_card 2×2 ══════════════
    LX2 = 5.7  # 좌존 우변
    # 좌상: 실물 답변 스크린샷 (읽히게 — 폭 최대)
    _subhead(slide, G.MARGIN_L, 0.80, LX2 - G.MARGIN_L, c["mid_subhead"])
    fit_picture(
        slide, IMG_CROP, G.MARGIN_L, 1.14, LX2 - G.MARGIN_L, 3.30, align="left", line=C["muted"]
    )
    # 좌하: 출력 스키마 카드
    _schema_card(slide, c, G.MARGIN_L, LX2, 4.62)
    # 존 구분선
    add_box(slide, LX2 + 0.12, 0.84, 0.012, 6.2, fill=C["bg_alt"])

    # 우: hero_card 2×2 (검증 내용은 카드3·4에 흡수 — 별도 flow 없음)
    RX = LX2 + 0.34
    gx, gy = 0.18, 0.14
    cw = (G.RIGHT_EDGE - RX - gx) / 2
    ch = 3.02
    for i, (num, title, tag, ic, banner, items, chips) in enumerate(CARDS):
        cx = RX + (i % 2) * (cw + gx)
        cyr = 0.82 + (i // 2) * (ch + gy)
        D.hero_card(slide, cx, cyr, cw, ch, num, title, tag, ic, banner, items, chips, hero=0.40)
    D.source_line(slide, c["source"])
    return slide
