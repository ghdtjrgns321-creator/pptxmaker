"""goldenfab.common — build_golden 공용 헬퍼 이관본(content_header·source_line·pct_bar)."""

from . import grid as G  # noqa: F401
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]
MARGIN = K["layout"]["margin"]


def content_header(slide, part_no, part_title, headline):
    """본문 장 공통 헤더 — 킥커 + 헤드라인 1줄 + 헤어라인 (BCG 문서형)."""
    add_text(
        slide,
        MARGIN,
        0.42,
        6.0,
        0.28,
        f"{part_no}. {part_title}",
        S["caption"],
        F["head"],
        C["muted"],
        bold=True,
    )
    add_text(
        slide,
        MARGIN,
        0.72,
        SLIDE_W - 2 * MARGIN,
        0.55,
        headline,
        S["section"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_box(slide, MARGIN, 1.55, SLIDE_W - 2 * MARGIN, 0.014, fill=C["muted"])


def source_line(slide, text):
    """하단 출처선 — 골든 문법 고정 요소."""
    add_box(slide, MARGIN, SLIDE_H - 0.62, SLIDE_W - 2 * MARGIN, 0.01, fill=C["bg_alt"])
    add_text(
        slide,
        MARGIN,
        SLIDE_H - 0.55,
        SLIDE_W - 2 * MARGIN,
        0.3,
        text,
        S["foot"],
        F["body"],
        C["muted"],
    )


def pct_bar(slide, x, y, w, h, label, pct, pct_label, note):
    """100% 비례 바 — 회색 전체 = 100%, accent 구간 = pct(0~1)."""
    add_text(slide, x, y - 0.34, w, 0.3, label, S["body"], F["body"], C["primary"], bold=True)
    add_box(slide, x, y, w, h, fill=C["bg_alt"])
    seg_w = max(w * pct, 0.035)
    add_box(slide, x, y, seg_w, h, fill=C["accent"])
    add_text(
        slide,
        x + w + 0.15,
        y + h / 2 - 0.22,
        1.6,
        0.44,
        pct_label,
        S["title"],
        F["head"],
        C["accent"],
        bold=True,
    )
    add_text(slide, x, y + h + 0.06, w + 1.6, 0.3, note, S["caption"], F["body"], C["muted"])
