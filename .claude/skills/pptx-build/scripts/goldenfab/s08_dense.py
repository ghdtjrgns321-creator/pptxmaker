"""S8 용어사전 — 고밀도 개편본 (s06 디자인 언어, 2026-07-20). 골든 편입 대상.

원본 s08_variants.variant_c(큰 28pt 헤더 + 3서사 텍스트 컬럼 + 이분매핑 + JSON 카드 +
검은 재서술 바)를 s06 공식으로 재배치:
- 압축 헤더(dense.compact_header)
- 상단 시그니처 밴드: 이분매핑(주인공, 좌) + 실물 JSON 카드(우)
- 하단: **s06 표준 hero_card 3장**(왜/역할/구축) — s06과 동일 컴포넌트, 카드 품질 일치
- 검은 전폭 바 폐지 → 결론은 카드3 불릿에 흡수 (하단 한줄평 바 금지, design-rules §8)

콘텐츠는 s08_variants.DEFAULT_C(원문 그대로) + 서사 불릿은 그 문장을 명사구로 압축.
매핑 지오메트리(MAP_*)만 이 파일에서 재정의 — 로직은 원본과 동일(설계 주석은 s08_variants 참조).
"""

from pptx.enum.text import PP_ALIGN

from . import dense as D
from . import grid as G
from .figures import Box
from .figures import bipartite_map as BIMAP
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit

K = load_kit()

# content_contract 계약 키 — registry가 이 장을 golden.tech_evidence로 부를 때 필수 키 출처.
# 2026-07-26 `s08_variants`(sparse 시안)에서 이관. 골든 전용 글이므로 실전 덱은 content로
# 전량 덮어야 하고, 안 덮으면 공장 문이 빌드를 중단한다.
DEFAULT_C = {
    "headline": "용어사전 — 실무 언어를 기준서 개념에 잇는 진입 색인",
    "kicker": "3. 기술 설명 — TECH 01 · Analyze",
    # 상단 3칼럼 서사 (head, body) — 번호 접두는 코드에서 위치로 생성
    "narratives": [
        (
            "왜 필요한가",
            "실무는 '리베이트'라 말하고 기준서는 '고객에게 지급할 대가'라 쓴다. 이 언어 간극을 잇지 않으면 검색이 시작조차 되지 않는다.",
        ),
        (
            "파이프라인에서의 역할",
            "Analyze 단계의 진입 관문. 질문에 등재 용어가 문자 그대로 나타나면 연결 개념으로 그래프에 진입한다 — 유사도 점수 없이, 사람이 승인한 경로로만.",
        ),
        (
            "어떻게 만들었나",
            "사람 1차 자료 3종(질의 매핑 288 · 사례 제목 123 · 부록A 정의 9)에서 AI가 초안을 내고 사람이 전수 검수해 423개 등재 — 신규 창작 0건.",
        ),
    ],
    # 소제목은 **짧게** — 15pt로 31자를 쓰면 5.2"라 범례(4.75 시작)를 덮어 글자가 뭉갠다
    # (2026-07-15 렌더 실측). 발췌 건수·등급 분해는 아래 캡션이 진다.
    "map_head": "실무가 쓰는 말 → 기준서가 쓰는 말",
    # 실물 엔트리 (data/ontology/aliases.json 실측 — 축약만, 창작 0).
    # (용어, 등급, [연결 개념]) — **우열 개념은 코드가 이 목록에서 파생**한다(등장순 dedup).
    # 그래서 N:1 수렴(`고객에게 지급할 대가` ← 리베이트+상품권)이 손으로 적는 게 아니라
    # 데이터에서 나온다. 등급 문자열도 aliases.json 원문 그대로("위임판단" 아님).
    "terms": [
        ("리베이트", "자동", ["고객에게 지급할 대가"]),
        ("밀어내기", "자동", ["위탁약정"]),
        ("볼륨디스카운트", "자동", ["변동대가", "변동대가 추정치를 제약함"]),
        ("상품권", "자동(위임판단)", ["고객에게 지급할 대가", "고객이 행사하지 아니한 권리"]),
        ("반품의 회계처리", "자동", ["반품권이 있는 판매", "본인 대 대리인의 고려사항"]),
    ],
    # 범례는 **형태 언어만**, 수치는 캡션만. 2026-07-15 제3자 채점 FAIL(#3): "자동 316"·
    # "자동(위임판단) 86"이 범례와 캡션에 글자 그대로 두 번 있었다.
    "map_legend": "실선 = 자동    점선 = 자동(위임판단) — AI가 판단, 사람이 승인",
    "map_caption": "등재 423 중 발췌 5건 — 423 = 자동 316 + 자동(위임판단) 86 + 검토 18 + 사용자 확정 1 + 제외 2",
    "json_title": "aliases.json — 실제 엔트리",
    "json_lines": [  # 상품권 엔트리 실물 축약
        '{ "term": "상품권",',
        '  "sources": ["query-mapping"],',
        '  "grade": "자동(위임판단)",',
        '  "concepts": [',
        '    "고객에게 지급할 대가",',
        '    "고객이 행사하지 아니한 권리" ],',
        '  "decision": {',
        '    "by": "AI 위임 판단",',
        '    "reason": "미행사 상품권 =',
        '        B44~47 정면 조항" } }',
    ],
    "json_caption": "모든 엔트리가 결정 로그(누가·왜)를 갖는다 — 전건 추적.",
    "bar": "AI가 만드는 것은 색인 하나 — 틀려도 1종(놓침)으로 드러나는 자리에만 둔다",
    "source": "출처: 2_DATA-TAXONOMY.md §2.6 · 4_SEARCH-PIPELINE.md (00_factsheet.md §C·§D)",
}


# ── 이분 매핑 GRID (표를 걷어낸 자리 — 2026-07-15 사용자 반려 "이해가 안 가는 시각화") ──
#
# 물성 선언(design-rules P1): **잇기(N:M)**. 용어사전이 하는 일은 실무 언어를 기준서 언어에
#   잇는 것이다. 표는 그 관계를 4열 나열로 바꿔 쉼표 하나로 뭉갰다 — 한 용어가 여러 개념으로
#   갈라지는 것(1:N)도, 여러 용어가 한 개념으로 모이는 것(N:1)도 안 보였다. 그게 본질인데.
#   게다가 죽은 열이 절반이었다: `원천` 5행 중 4행이 같은 값, `등급` 4행이 "자동".
#   → 좌열(실무어) ─선─▶ 우열(기준서 개념). 갈래와 수렴이 **선의 모양 그 자체**로 보인다.
#   §6 "숫자·항목 단독 나열 금지 — 관계 인코딩과 결합" / §6 "나열은 도형에 가두기".
#
# 색(P4⑩): accent 예산은 **1**뿐이다 — 3칼럼 번호(01·02·03) 런이 이미 3을 쓴다(오딧 상한 4).
#   그 1을 `자동(위임판단)` 칩 테두리에 준다. 이 장의 결론 바가 "AI가 만드는 것은 색인 하나"이므로
#   **AI가 판단한 자리**가 accent를 가져가는 게 맞다. 선은 muted — 선까지 accent면 예산 초과다.
#   §5 점선 = 불확실 → 위임판단 매핑은 점선.
C, S, F = K["rgb"], K["sizes"], K["fonts"]

# 서사 3칼럼 — 원문(DEFAULT_C narratives)을 아이콘 카드용 2단 불릿으로 압축 (사실 보존)
# 하단 카드 3 = s06 표준 hero_card 그대로 (배지·태그·중앙 배너·히어로 아이콘·불릿·칩).
# (번호, 제목, 태그, 히어로 아이콘, 배너, 불릿3[(리드,설명)], 칩3[(아이콘,강조,라벨)])
# 카드 주제 = 용어사전의 고유 성질 (제네릭 왜/역할/구축 대신, 각 장 내용에 맞춤)
CARDS = [
    (
        "1",
        "언어 간극",
        "실무어 ↔ 기준서어",
        "link",
        "표현은 달라도 같은 개념",
        [
            ("실무 '리베이트'", " 기준서 '지급 대가'"),
            ("간극 미연결 시", " 검색 시작 불가"),
            ("여러 용어 → 한 개념", " N:M 수렴"),
        ],
        [
            ("link", "N:M", " 잇기"),
            ("book-open", "423", " 등재어"),
            ("circle-check-big", "사람 승인", " 경로"),
        ],
    ),
    (
        "2",
        "결정적 진입",
        "임베딩 0 · substring",
        "search",
        "유사도 점수 없는 진입",
        [
            ("질문 속 문자", " 그대로 포착"),
            ("substring 매칭", " 확률·유사도 없음"),
            ("Analyze 첫 관문", " 그래프 입구"),
        ],
        [
            ("search", "substring", " 결정적"),
            ("circle-x", "임베딩", " 0"),
            ("git-branch", "그래프", " 진입"),
        ],
    ),
    (
        "3",
        "후보 진입점",
        "확정 라우터 아님",
        "git-branch",
        "최종 선택은 질문 문맥",
        [
            ("진입점 제안만", " 확정은 안 함"),
            ("한 용어 → 여러 개념", " 다중 목적지"),
            ("최종 선택", " 질문 문맥이 담당"),
        ],
        [
            ("git-branch", "다중", " 목적지"),
            ("circle-x", "확정", " 아님"),
            ("circle-check-big", "문맥", " 선택"),
        ],
    ),
    (
        "4",
        "사람 검수 색인",
        "3원천 · AI 창작 0",
        "book-open",
        "등재 423 · 전건 결정 로그",
        [
            ("1차 자료 3종", " 288·123·9"),
            ("AI 초안 + 전수 검수", " 신규 창작 0"),
            ("전건 결정 로그", " 누가·왜 추적"),
        ],
        [
            ("book-open", "288·123·9", " 3원천"),
            ("circle-check-big", "423", " 등재"),
            ("file-text", "316", " 자동 등급"),
        ],
    ),
]

# ── 이분 매핑 지오메트리 — 압축 헤더에 맞춰 위로 당김 (원본 3.88~5.97 → 3.12~5.90) ──
# 상단 시그니처 밴드(매핑+JSON) — 하단은 hero_card 3장이 차지하므로 매핑을 위로 올림
MAP_X = G.MARGIN_L
MAP_HEAD_Y = 0.82
MAP_L_RIGHT = 2.45
MAP_R_X = 4.75
MAP_R_MAX = 8.20
CHIP_INSET = 0.44
CHIP_FLOOR = 0.95
MAP_TOP = 1.30
MAP_BOTTOM = 3.30
MAP_H = 0.24
DELEGATED = "자동(위임판단)"


# ── 이분 매핑 자리 (2026-07-26 부품화) ──────────────────────────────────────────
# 도해는 `figures.bipartite_map`이 그린다. 이 장이 정하는 것은 **자리**뿐이고, 자리는
# 골든 실측 좌표에서 역산한다 — 좌칩 우변이 MAP_L_RIGHT, 우칩 좌변이 MAP_R_X에 오도록.
#
# 부품은 좌열 폭을 글자에서 파생하므로 box의 좌변은 "MAP_L_RIGHT − 좌열 폭"이다.
# `halign=left`·`valign=top`으로 부품의 자동 정렬을 끄고 골든 좌표를 그대로 쓴다
# (정렬을 켜면 남는 자리만큼 도해가 밀려 픽셀 동일이 깨진다).
MAP_GAP = MAP_R_X - MAP_L_RIGHT  # 선 길이 2.30"
MAP_LAYOUT = {"gap": MAP_GAP, "halign": "left", "valign": "top", "chip_h": MAP_H}


def _map_data(terms):
    """골든 terms(3튜플) → 부품 계약. `자동(위임판단)`이 점선·accent 테두리를 켠다."""
    return {"pairs": [{"left": t, "right": cs, "soft": g == DELEGATED} for t, g, cs in terms]}


def bipartite_map(slide, terms):
    data = _map_data(terms)
    _, _, l_w, r_w = BIMAP.measure(data, K, MAP_LAYOUT)  # 열 폭은 글자에서 나온다
    box = Box(
        MAP_L_RIGHT - l_w,                 # 좌칩 우변이 MAP_L_RIGHT에 오도록 역산
        MAP_TOP,
        l_w + MAP_GAP + r_w,
        MAP_BOTTOM - MAP_TOP + MAP_H,      # 마지막 칩 top이 MAP_BOTTOM
    )
    return BIMAP.draw(slide, box, data, K, MAP_LAYOUT)


def _subhead(slide, x, y, w, text):
    add_text(slide, x, y, w, 0.30, text, S["head"], F["head"], C["primary"], bold=True)
    D.hrule(slide, x, y + 0.34, w, color=D.line_mid, weight=1.0)


def build(prs, c=None):
    c = {**DEFAULT_C, **(c or {})}
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    D.compact_header(slide, c["kicker"], c["headline"])

    # ── 상단 시그니처 밴드 좌: 이분 매핑 (주인공) ──
    map_w = MAP_R_MAX - MAP_X
    _subhead(slide, MAP_X, MAP_HEAD_Y, map_w, c["map_head"])
    add_text(
        slide,
        MAP_R_X,
        MAP_HEAD_Y + 0.05,
        MAP_R_MAX - MAP_R_X,
        0.22,
        c["map_legend"],
        S["foot"],
        F["body"],
        C["muted"],
        align=PP_ALIGN.RIGHT,
    )
    bipartite_map(slide, c["terms"])
    add_text(slide, MAP_X, 3.36, map_w, 0.18, c["map_caption"], S["foot"], F["body"], C["muted"])

    # ── 상단 시그니처 밴드 우: 실물 JSON 카드 (캡션 겹침 금지 — 카드 밖 아래) ──
    jx, jy, jw, jh = 8.7, MAP_HEAD_Y - 0.02, G.RIGHT_EDGE - 8.7, 2.50
    card = add_box(slide, jx, jy, jw, jh, fill=C["primary"], shape="round")
    try:
        card.adjustments[0] = 0.05
    except (IndexError, ValueError):
        pass
    add_text(
        slide,
        jx + 0.22,
        jy + 0.14,
        jw - 0.44,
        0.22,
        c["json_title"],
        S["foot"],
        F["head"],
        C["bg_alt"],
        bold=True,
    )
    add_text(
        slide,
        jx + 0.22,
        jy + 0.44,
        jw - 0.44,
        jh - 0.56,
        [[(ln, {})] for ln in c["json_lines"]],
        S["foot"],
        "Consolas",
        C["bg"],
        line_spacing=1.18,
    )
    add_text(slide, jx, 3.36, jw, 0.18, c["json_caption"], S["foot"], F["body"], C["muted"])

    # ── 하단: s06 표준 hero_card 4장 (s06과 동일 4칸 — 카드 밀도 일치, 휑함 해소) ──
    card_w = (D.FULL_W - 3 * 0.18) / 4
    for i, (num, title, tag, ic, banner, items, chips) in enumerate(CARDS):
        cx = G.MARGIN_L + i * (card_w + 0.18)
        D.hero_card(
            slide, cx, 3.62, card_w, 3.32, num, title, tag, ic, banner, items, chips, hero=0.44
        )
    # 결론("AI 개입 = 색인 1종")은 카드3 불릿에 흡수 — 하단 전폭 한줄평 바 금지(§8).
    D.source_line(slide, c["source"])
    return slide
