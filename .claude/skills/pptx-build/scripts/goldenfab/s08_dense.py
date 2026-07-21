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

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from . import dense as D
from . import grid as G
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit, set_shape_text
from .s08_variants import DEFAULT_C

K = load_kit()
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


def _pitch(n, top, bottom, h, cap=0.42, *, what="항목"):
    if n <= 1:
        return h
    p = min(cap, (bottom - top - h) / (n - 1))
    if p < h:
        raise ValueError(f"{what} {n}개는 이 구획에 안 들어간다 — 피치 {p:.3f} < 칩 {h}")
    return p


def chip_w(text, pt=None, inset=CHIP_INSET, floor=CHIP_FLOOR):
    pt = pt or S["caption"]
    return max(len(text.strip()) * pt * 0.8 / 72 + inset, floor)


def _map_line(slide, x1, y1, x2, y2, *, dashed=False):
    conn = slide.shapes.add_connector(2, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = C["muted"]
    conn.line.width = Pt(1.0)
    ln = conn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "sm", "len": "sm"}))


def bipartite_map(slide, terms):
    """실무어 → 기준서 개념 이분 매핑 (s08_variants 로직 이식, 지오메트리만 재정의)."""
    concepts = []
    for _t, _g, cs in terms:
        for c_ in cs:
            if c_ not in concepts:
                concepts.append(c_)
    r_pitch = _pitch(len(concepts), MAP_TOP, MAP_BOTTOM, MAP_H, what="연결 개념")
    r_y = {c_: MAP_TOP + i * r_pitch for i, c_ in enumerate(concepts)}
    r_span = (len(concepts) - 1) * r_pitch
    l_pitch = _pitch(len(terms), MAP_TOP, MAP_BOTTOM, MAP_H, what="용어")
    l_span = (len(terms) - 1) * l_pitch
    l_top = MAP_TOP + (r_span - l_span) / 2
    l_y = {t: l_top + i * l_pitch for i, (t, _g, _cs) in enumerate(terms)}
    fan_in = {}
    for _t, _g, cs in terms:
        for c_ in cs:
            fan_in[c_] = fan_in.get(c_, 0) + 1
    slot = dict.fromkeys(fan_in, 0)
    for t, g, cs in terms:
        for c_ in cs:
            n_in = fan_in[c_]
            off = (slot[c_] + 1) / (n_in + 1) if n_in > 1 else 0.5
            slot[c_] += 1
            _map_line(
                slide,
                MAP_L_RIGHT,
                l_y[t] + MAP_H / 2,
                MAP_R_X,
                r_y[c_] + MAP_H * off,
                dashed=(g == DELEGATED),
            )
    for t, g, _cs in terms:
        deleg = g == DELEGATED
        w = chip_w(t)
        b = add_box(
            slide,
            MAP_L_RIGHT - w,
            l_y[t],
            w,
            MAP_H,
            fill=C["bg"],
            line=C["accent"] if deleg else C["muted"],
            line_w=1.25 if deleg else 0.75,
            shape="round",
        )
        set_shape_text(b, t, S["caption"], F["head"], C["primary"], bold=True)
    for c_ in concepts:
        b = add_box(
            slide,
            MAP_R_X,
            r_y[c_],
            min(chip_w(c_), MAP_R_MAX - MAP_R_X),
            MAP_H,
            fill=C["bg_alt"],
            shape="round",
        )
        set_shape_text(b, c_, S["caption"], F["body"], C["primary"])


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
