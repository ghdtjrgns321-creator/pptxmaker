"""S10 스크린샷(지식그래프 3D 뷰어) — dense 개편(2026-07-20 재작업 v8). §F 스크린샷 장 = 실물 증거.

경계(사용자 확정): 그래프 **구조·타입·규모·검색 방식은 전부 S9(지식그래프 장)가 설명**한다.
s10이 그걸 또 그리면 자가복제다(스키마 노드-링크·타입 카드·규모 스탯·검색 순회 전부 S9 재탕).
s10의 고유 역할은 하나 — "이 그래프가 **실제로 돌아가는 실물 화면**"이라는 증거(§F).

그래서 이 장은:
- 큰 캡처(주인공)
- 곁들이는 건 **다이어그램이 증명 못 하는 것만** = 실물 도구성(실시간 탐색 UI·레이어 토글·
  한 화면 렌더·실운영 V1). 그래프 구조/타입/규모 재설명 금지(S9 경계).

콘텐츠는 SHOT_DEFAULTS(캡처·헤더) + 화면 UI 실측(§7). 카드·스키마·스탯 전부 제거.
"""

from . import dense as D
from . import grid as G
from . import s10_screenshot as S10
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, fit_picture, load_kit

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]

# content_contract 계약 키 — registry가 이 dense를 golden.screenshot으로 부를 때 필수 키 출처.
DEFAULT = S10.SHOT_DEFAULTS

SOURCE = "출처: 3_KNOWLEDGE-GRAPH.md · knowledge_graph_3d.png (실운영 초기 V1 스냅샷 · 수치는 그래프 v14) · 그래프 설계·규모는 S9"

# 하단 운영 증거 띠 — 다이어그램이 못 보여주는 '실물 도구'라는 것(§F). 그래프 구조 재설명 아님.
OPS = [
    ("cpu", "실시간 3D 탐색", "클릭 상세 · 드래그 회전 · 휠 확대 · 더블클릭 초기화"),
    ("layers", "레이어 토글", "개념 · 문단 · 사례 · 간선 종류별 on · off 표시"),
    ("workflow", "한 화면에 전체", "끊긴 섬 없이 하나의 구조로 렌더 — 눈으로 확인"),
    ("circle-check-big", "실운영 V1 화면", "설계 목업 · 재현이 아닌 실제 시스템 캡처"),
]

# 캡처 — 주인공. 가운데 큰 배치(2.05:1 와이드). 하단 증거 띠 자리를 남긴다.
IMG_W = 10.6
IMG_X = (SLIDE_W - IMG_W) / 2
IMG_Y = 0.88


def _op_card(slide, x, y, w, h, icon_name, label, sub):
    add_box(slide, x, y, w, h, fill=C["bg_alt"], shape="round")
    D.icon(slide, icon_name, "accent", x + 0.2, y + (h - 0.36) / 2, 0.36)
    add_text(
        slide,
        x + 0.68,
        y + 0.12,
        w - 0.82,
        0.24,
        label,
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_text(
        slide,
        x + 0.68,
        y + 0.38,
        w - 0.82,
        0.3,
        sub,
        S["foot"],
        F["body"],
        C["muted"],
        line_spacing=1.1,
    )


def build(prs, c=None):
    c = {**DEFAULT, **(c or {})}
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    D.compact_header(slide, c["kicker"], c["headline"])

    # ── 캡처(주인공) ── 캡션 생략(헤드라인·출처가 이미 함 · 하단 띠와 겹침 방지)
    fit_picture(slide, c["img"], IMG_X, IMG_Y, IMG_W, 5.2, align="left", line=C["muted"])

    # ── 하단 운영 증거 띠(4) — 실물 도구성. 그래프 구조/타입/규모는 S9 ──
    n = len(OPS)
    gap = 0.18
    cw = (D.FULL_W - (n - 1) * gap) / n
    oy, oh = 6.16, 0.74
    for i, (icon_name, label, sub) in enumerate(OPS):
        _op_card(slide, G.MARGIN_L + i * (cw + gap), oy, cw, oh, icon_name, label, sub)

    D.source_line(slide, SOURCE)
    return slide
