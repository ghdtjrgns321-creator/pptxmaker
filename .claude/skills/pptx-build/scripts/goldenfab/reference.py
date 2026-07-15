"""goldenfab.reference — 골든 레퍼런스 덱(19장) 정의. 회귀 게이트의 기준 산출물.

## 왜 이 파일이 생겼나 (2026-07-15 단일화)

원래 `golden/`(동결 원본)과 `goldenfab/`(공장)에 **같은 레이아웃 코드가 두 벌** 있었다.
Phase 2 "공장 역설계 이식"의 비계였다 — 이식이 충실한지 1회 검증하려고 원본을 남겨둔 것
(docs/user/07). 이식은 15/15·490/490으로 끝났는데 비계를 안 걷었고, "수정은 goldenfab에서만"
(05-plan 결정 #9) 규칙도 깨져 양쪽을 다 고치게 되면서 **1회성 검증이 영구 동기화 세금**이 됐다.
장 하나 고치려면 두 파일을 똑같이 고쳐야 했고, 한쪽만 고치면 게이트가 깨졌다.

## 지금 구조

    goldenfab/*.py (유일 소스) ─→ reference.build_reference() ─┬─→ compare_golden: 스냅샷 대조
                                                                └─→ golden/build_golden.py: pptx 렌더

골든 파일 테스트의 정석대로 **결과물(도형 서명)을 스냅샷으로 박아** 비교한다. 소스를 복사해두는
게 아니다 — 복사본은 독립 검증력이 없고(그냥 사본) 동기화 비용만 든다.
스냅샷: `assets/golden-snapshot.json`. 의도적 변경 시 `compare_golden.py --update-snapshot`으로
재생성하고 git diff로 리뷰한다.

## 이 덱의 텍스트에 대하여

내용은 K-IFRS 1115다(골든 제작 당시 소재). **납품물이 아니다** — 공장 문(content_contract)이
골든 기본값의 유출을 막으므로 실제 프로젝트 덱에는 한 글자도 안 나간다. 여기 텍스트의 쓸모는
**분량**이다: 실물급 길이여야 레이아웃(줄바꿈·박스 높이)이 제대로 검증된다.
"""

from pptx.util import Inches

from .kit import SLIDE_H, SLIDE_W, load_kit
from .registry import LAYOUTS

K = load_kit()

# ── 레퍼런스 덱 콘텐츠 (golden/build_golden.py에서 이관 — 00_factsheet.md 실측, 수치 창작 금지) ──
TITLE = "K-IFRS 1115 온톨로지 지식그래프 RAG"
VALUE_PROP = "확실할 때만 확정하고, 애매하면 근거를 보여주며 유보한다 — 2종 오류를 구조로 차단하는 수익인식 특화 챗봇"
KICKER = "TECHNICAL PROPOSAL · K-IFRS 1115"
YEAR = "2026"

PART_LEADS = [  # 간지 리드 — 00_factsheet.md(실데이터 v2) 부별 요약
    "일반 LLM의 허위 확정 리스크와, 초기 임베딩 RAG 운영에서 실측된 4가지 균열",
    "질문에서 답까지 4개 노드 전부가 결정적으로 동작합니다 — 진입부 임베딩 유사도 0",
    "기준서의 명시적 관계를 그대로 옮긴 지식그래프와, 회계 항등식으로 전수 검증한 데이터입니다",
    "실측된 네 균열마다 구조적 대체물을 배치했습니다 — 확률 신호의 전량 폐기",
    "점수 경쟁이 사라진 결정적 경로, AI 개입을 색인 하나로 제한한 신빙성 설계입니다",
    "정답 문서를 차단한 채 사람 전문가의 결론을 78/92 재현했습니다 — 수치가 실제 능력을 반영합니다",
]

TOC = [  # (제목, 서브카피 ' · ' 구분, 시작 페이지) — 00_factsheet.md §A~§H
    ("문제 정의", "치명적 오류는 2종(허위 확정) · 일반 LLM이 실패하는 4가지 이유", "03"),
    ("파이프라인", "Analyze → Retrieve → Generate → Format · 4개 결정 노드", "05"),
    ("기술 설명", "용어사전 · 지식그래프 · 판단트리 · 구조화 출력 — 노드 순서대로 4개 기술", "07"),
    ("문제와 해결", "확률 신호의 실패 실측(탈락 103/105) · 전·후 5개 지표 재현 72→78", "13"),
    ("차별점", "일반 임베딩 RAG 대비 7개 축 비교", "16"),
    ("성과·증거", "홀드아웃 78/92 · 에러 0건 · 미재현 원인 분해", "19"),
]

# 장 순서 — (레이아웃 키, 리포트용 이름). content가 None이면 그 레이아웃의 골든 기본값 사용.
SLIDE_ORDER = [
    ("cover", "S1 표지"),
    ("toc", "S2 목차"),
    ("part", "S3 간지1"),
    ("problem_grid", "S4 문제정의"),
    ("part", "S5 간지2"),
    ("exec_graph", "S6 실행그래프"),
    ("part", "S7 간지3"),
    ("tech_evidence", "S8 용어사전"),
    ("tech_tree", "S9 지식그래프"),
    ("screenshot", "S10 스크린샷"),
    ("tech_mechanism", "S11 판단트리"),
    ("tech_capture", "S12 구조화출력"),
    ("part", "S13 간지4"),
    ("ab_simulation", "S14 트러블슈팅"),
    ("validation", "S15 골든테스트"),
    ("part", "S16 간지5"),
    ("mirror_matrix", "S17 미러매트릭스"),
    ("boundary", "S18 경계"),
    ("closing", "S19 클로징"),
]


def cover_content():
    return {
        "kicker": KICKER,
        "title": TITLE,
        "value_prop": VALUE_PROP,
        "year": YEAR,
        "brand_name": K["brand"]["name"],
    }


def toc_content():
    return {"title": "목차", "items": TOC, "deck_title": TITLE, "deck_sub": "기술 제안서 · 2026"}


def part_content(no):
    return {"no": no, "title": TOC[no - 1][0], "lead": PART_LEADS[no - 1]}


def new_presentation():
    from pptx import Presentation

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
    return prs


def build_reference(prs=None):
    """레퍼런스 덱 19장 조립. content 없는 레이아웃은 골든 기본값으로 렌더된다.

    골든 기본값을 **의도적으로** 쓰는 유일한 경로다(회귀 기준 산출). 공장 문
    `build_pptx._render_golden`은 여기를 안 거치므로 content_contract가 계속 강제된다.
    """
    if prs is None:
        prs = new_presentation()
    part_no = 0
    for key, _name in SLIDE_ORDER:
        if key == "cover":
            LAYOUTS[key](prs, cover_content())
        elif key == "toc":
            LAYOUTS[key](prs, toc_content())
        elif key == "part":
            part_no += 1
            LAYOUTS[key](prs, part_content(part_no))
        else:
            LAYOUTS[key](prs, None)  # 골든 기본값
    return prs


def slide_names():
    return [name for _key, name in SLIDE_ORDER]
