"""S9 지식그래프 — 고밀도 개편본 (s06 공식, 2026-07-20). 골든 편입 대상.

원본 s09_variants.variant_a(28pt 헤더 + 3서사 + 위계 트리(좌) + 간선 카탈로그(우) + 검은 바)를
S8과 동일한 승리 공식으로 재배치:
- 압축 헤더(dense.compact_header)
- 상단 시그니처 밴드: 위계 트리(주인공, 좌) + 간선 7종 카탈로그(주인공, 우) — 검증된 자산 재사용
- 하단: s06 표준 hero_card 4장(왜/역할/구축/규모검증) — 카드 컴포넌트 단일(dense.hero_card)
- 검은 전폭 바 폐지 → 결론은 카드1 불릿에 흡수 (하단 한줄평 바 금지, design-rules §8)

트리는 원본 좌표를 top-band로 선형 압축(ty/TH), 카탈로그는 s09_variants.build_catalog를
상단 밴드 좌표로 재사용(모듈 상수 오버라이드). 로직·색·노드 클래스 인코딩은 원본 그대로.
"""

from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from . import dense as D
from . import grid as G
from . import s09_variants as S9
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit, set_shape_text

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]

# content_contract 계약 키 — registry가 이 dense를 golden.tech_tree로 부를 때 필수 키 출처.
DEFAULT = S9.DEFAULT

# 하단 카드 4 = s06 표준 hero_card (배지·태그·중앙 배너·히어로 아이콘·불릿·칩)
# (번호, 제목, 태그, 히어로 아이콘, 배너, 불릿3[(리드,설명)], 칩3[(아이콘,강조,라벨)])
# 카드 주제 = 지식그래프의 고유 성질 (제네릭 왜/역할/구축 대신, 각 장 내용에 맞춤)
CARDS = [
    (
        "1",
        "구조로 검색",
        "텍스트 아닌 관계",
        "workflow",
        "임베딩이 놓치는 법적 이웃",
        [
            ("문단 뭉치", " 순서·관계 없음"),
            ("기준서", " 계층·참조·사례 구조"),
            ("법적 이웃", " 관계가 잡는다"),
        ],
        [
            ("git-branch", "관계", " 로 연결"),
            ("search", "구조", " 검색"),
            ("circle-x", "임베딩", " 놓침"),
        ],
    ),
    (
        "2",
        "1홉 수집 지도",
        "Retrieve의 지도",
        "search",
        "수집 범위가 그래프로 고정",
        [
            ("진입 개념에서", " 관계 1단계만"),
            ("문단·사례·근거", " 결정적 수집"),
            ("수집 범위 자체", " 그래프로 고정"),
        ],
        [
            ("git-branch", "1홉", " 탐색"),
            ("layers", "문단·사례", " 수집"),
            ("circle-check-big", "결정적", " 범위"),
        ],
    ),
    (
        "3",
        "기계 생성 뼈대",
        "AI 판단 0",
        "git-branch",
        "공식 소제목 → 개념 80",
        [
            ("공식 소제목 80", " 그대로 개념 노드"),
            ("뼈대 간선", " 기준서에서 기계 생성"),
            ("AI 판단 0", " 사람 개입 최소"),
        ],
        [
            ("git-branch", "개념 80", ""),
            ("circle-x", "AI 판단", " 0"),
            ("book-open", "문단 250", ""),
        ],
    ),
    (
        "4",
        "규모·검증",
        "노드 929 · 간선 2,694",
        "layers",
        "고립 0 · 전수 감사 확인",
        [
            ("노드 929", " 개념+문단+사례+용어"),
            ("간선 2,694", " 7종 + BC 근거층"),
            ("고립 노드 0", " 감사로 확인"),
        ],
        [
            ("layers", "929", " 노드"),
            ("git-branch", "2,694", " 간선"),
            ("circle-check-big", "0", " 고립"),
        ],
    ),
]

# ── 트리 선형 압축 — 원본 y(3.4~6.2) → 상단 밴드(0.95~) ──
BAND_A, BAND_B = 0.80, 0.95  # new_y = (y-3.4)*0.80 + 0.95


def ty(y):
    return (y - 3.4) * BAND_A + BAND_B


TH = 0.34 * BAND_A  # 압축된 노드 높이


def _tree(slide, c):
    """원본 위계 트리(variant_a)를 상단 밴드로 선형 압축 이식. 노드는 NODE_STYLE 단일 출처."""
    add_text(
        slide,
        G.MARGIN_L,
        0.80,
        4.0,
        0.24,
        c["struct_head"],
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
    )

    def _node(x, y, w, name, kind):
        return S9._node_chip(slide, x, ty(y), w, name, kind, h=TH, size=S["foot"])

    root_x, root_w = 2.35, 1.6
    _node(root_x, 3.6, root_w, c["root"], "root")
    concept_pos = [(2.05, 1.0), (3.2, 0.7), (4.05, 0.65)]
    for name, (cx, cw) in zip(c["concepts"], concept_pos):
        _node(cx, 4.45, cw, name, "concept")
        S9._arrow(slide, root_x + root_w / 2, ty(3.94), cx + cw / 2, ty(4.45))
    para_pos = [(1.75, 0.8, 2.55), (2.95, 0.8, 2.55), (3.95, 0.6, 3.55)]
    for name, (px, pw, src_cx) in zip(c["paras"], para_pos):
        _node(px, 5.35, pw, name, "para")
        S9._arrow(slide, src_cx, ty(4.79), px + pw / 2, ty(5.35))
    add_text(
        slide,
        G.MARGIN_L,
        ty(4.5),
        1.15,
        0.20,
        c["layer_tag1"],
        S["foot"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_text(
        slide,
        G.MARGIN_L,
        ty(5.4),
        1.15,
        0.20,
        c["layer_tag2"],
        S["foot"],
        F["head"],
        C["primary"],
        bold=True,
    )
    _node(G.MARGIN_L, 3.6, 1.45, c["term_chip"], "term")
    add_text(
        slide, G.MARGIN_L, ty(3.98), 1.5, 0.16, c["term_cap"], S["foot"], F["body"], C["muted"]
    )
    S9._arrow(slide, G.MARGIN_L + 1.45, ty(3.77), 2.05, ty(4.62))
    _node(G.MARGIN_L, 5.85, 1.1, c["case_chip"], "case")
    S9._arrow(slide, G.MARGIN_L + 1.1, ty(6.02), 1.75, ty(5.69))
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(2.55), Inches(ty(5.52)), Inches(2.95), Inches(ty(5.52))
    )
    conn.line.color.rgb = C["muted"]
    conn.line.width = Pt(1.25)
    ln = conn.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "sm", "len": "sm"}))
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "sm", "len": "sm"}))
    add_text(
        slide,
        G.MARGIN_L,
        ty(6.24),
        4.3,
        0.14,
        c["node_caption"],
        S["foot"] - 1,
        F["body"],
        C["muted"],
    )


def build(prs, c=None):
    c = {**DEFAULT, **(c or {})}
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    D.compact_header(slide, c["kicker"], c["headline"])

    # ── 상단 시그니처 밴드 좌: 위계 트리 (주인공) ──
    _tree(slide, c)

    # ── 상단 시그니처 밴드 우: 간선 7종 카탈로그 (주인공) — 원본 로직 재사용, 좌표만 상단으로 ──
    S9.CAT_HEAD_Y, S9.CAT_TOP, S9.CAT_BOTTOM = 0.80, 1.22, 3.28
    add_text(
        slide,
        S9.CAT_X,
        0.80,
        S9.CAT_W,
        0.24,
        [
            [
                (c["cat_title"][0], {"bold": True, "color": C["primary"], "font": F["head"]}),
                (c["cat_title"][1], {}),
            ]
        ],
        S["caption"],
        F["body"],
        C["muted"],
    )
    S9.build_catalog(slide, c["cat_rows"])

    # ── 하단: s06 표준 hero_card 4장 ──
    card_w = (D.FULL_W - 3 * 0.18) / 4
    for i, (num, title, tag, ic, banner, items, chips) in enumerate(CARDS):
        cx = G.MARGIN_L + i * (card_w + 0.18)
        D.hero_card(
            slide, cx, 3.62, card_w, 3.32, num, title, tag, ic, banner, items, chips, hero=0.44
        )
    # 결론("법적 이웃을 구조로 검색")은 카드1에 흡수 — 하단 전폭 한줄평 바 금지(§8).
    D.source_line(slide, c["source"])
    return slide
