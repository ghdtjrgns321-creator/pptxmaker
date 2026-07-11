"""S10 시안 — 스크린샷 전용 장 (신규 어휘: 풀블리드 실물 화면. UI 패스).

다크 배경 + 3D 뷰어 캡처 대형 + 헤더·캡션만. 실물 화면이 주인공인 장.
실행: uv run python golden/s10_screenshot.py → golden/variants/s10_screenshot.pptx
"""

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from . import grid as G
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit, mix

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]

IMG_3D = r"C:\Users\ghdtj\workspace\portfolio\k-ifrs-1115\images\knowledge_graph_3d.png"
EMU_IN = 914400


SHOT_DEFAULTS = {
    "kicker": "3. 기술 설명 — TECH 02 · Retrieve",
    "headline": "노드 929개가 하나의 구조로 — 지식그래프 3D 뷰어",
    "img": IMG_3D,
    "caption": "실제 운영 화면 — 지식그래프 3D 뷰어 (초기 V1 스냅샷, 수치 표기는 그래프 v14 기준)",
    "points": None,  # None이면 골든 기본 4포인트
}


def variant_a(prs, c=None):
    c = {**SHOT_DEFAULTS, **(c or {})}
    """A — 흰 배경(본문 무드 통일) 풀블리드: 캡처가 주인공, 텍스트는 헤더·캡션만."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    add_text(
        slide,
        G.MARGIN_L,
        0.42,
        8.0,
        0.28,
        c["kicker"],
        S["caption"],
        F["head"],
        C["muted"],
        bold=True,
    )
    add_text(
        slide,
        G.MARGIN_L,
        0.72,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.55,
        c["headline"],
        S["section"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_box(
        slide,
        G.MARGIN_L,
        G.RULE_Y,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.014,
        fill=C["muted"],
    )
    img_h = 4.75
    img_w = img_h * 3200 / 2000  # 7.6
    pic = slide.shapes.add_picture(c["img"], Inches(G.MARGIN_L), Inches(1.8), height=Inches(img_h))
    pic.line.color.rgb = C["muted"]
    pic.line.width = Inches(0.01)
    add_box(slide, G.MARGIN_L, 6.78, 0.5, 0.035, fill=C["accent"])
    add_text(
        slide,
        G.MARGIN_L + 0.65,
        6.69,
        img_w - 0.65,
        0.3,
        c["caption"],
        S["caption"],
        F["body"],
        C["muted"],
    )
    # ── 우측 관찰 칼럼 — 캡처 범례 실제 항목과 1:1 대응 ──
    ox = G.MARGIN_L + img_w + 0.4  # 8.6
    ow = G.RIGHT_EDGE - ox
    add_text(
        slide,
        ox,
        1.8,
        ow,
        0.3,
        "이 화면에서 보이는 것",
        S["head"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_box(slide, ox, 2.15, ow, 0.012, fill=C["muted"])
    points = c["points"] or [
        ("큰 파란 노드 = 개념", "기준서 소제목 80개 — 클수록·밝을수록 상위 계층이다."),
        (
            "작은 점 = 문단·사례",
            "회색은 문단 250, 색 점은 QNA·감리·IE 사례 188 — 전부 개념 아래 매달린다.",
        ),
        (
            "선 = 관계 간선",
            "회색 계층(목차) · 노랑 상호참조(E3) · 빨강 선행판단(E2) — 검색이 걷는 길이 그대로 보인다.",
        ),
        ("고립 노드 0", "모든 노드가 최소 하나의 경로로 본문에 닿는다 — 감사로 확인."),
    ]
    for i, (head, body) in enumerate(points):
        py = 2.4 + i * 1.05
        add_text(
            slide, ox, py, 0.5, 0.28, f"0{i + 1}", S["head"], F["head"], C["accent"], bold=True
        )
        add_text(
            slide,
            ox + 0.45,
            py + 0.02,
            ow - 0.45,
            0.26,
            head,
            S["body"],
            F["head"],
            C["primary"],
            bold=True,
        )
        add_text(
            slide,
            ox + 0.45,
            py + 0.32,
            ow - 0.45,
            0.6,
            body,
            S["caption"],
            F["body"],
            C["muted"],
            line_spacing=1.2,
        )
    return slide


def audit(prs):
    fails = []
    for si, sl in enumerate(prs.slides):
        pics = sum(1 for sh in sl.shapes if sh.shape_type == 13)
        if pics != 1:
            fails.append((si, "picture≠1", pics))
    assert not fails, f"AUDIT FAIL {fails}"
    print("audit pass")


def main():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    variant_a(prs)
    audit(prs)
    out = Path(__file__).parent / "variants" / "s10_screenshot.pptx"
    prs.save(str(out))
    print(f"saved {out}")


if __name__ == "__main__":
    main()
