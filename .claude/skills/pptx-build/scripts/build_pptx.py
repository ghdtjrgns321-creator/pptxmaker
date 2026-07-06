#!/usr/bin/env python
"""deck-spec(JSON) + brand-kit(YAML) -> 네이티브 .pptx 빌드.

표는 네이티브 표(GraphicFrame.table), 차트는 네이티브 차트(GraphicFrame.chart)로
생성한다 — 이미지가 아니므로 파워포인트에서 데이터만 고치면 그대로 반영된다.

브랜드 일관성은 brand-kit.yaml 단일 출처에서만 나온다. 레이아웃 상수도 이 파일에
박제되어 있어, 같은 spec은 항상 같은 결과를 낸다("공장 일관성").

usage: python build_pptx.py <deck-spec.json> <out.pptx> [--brand brand-kit.yaml]
"""

import argparse
import json
import re
import sys
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

# 16:9 고정 캔버스 (inch)
EMU_W, EMU_H = 13.333, 7.5

# 마스터 틀(컨설팅 표준) 고정 좌표 — 모든 슬라이드에 동일 위치로 스탬프
FOOT_Y = EMU_H - 0.42  # 푸터 텍스트 y (회사명·페이지번호)
HAIR_Y = EMU_H - 0.5  # 푸터 헤어라인 y
HEADER_HAIR_Y = 1.42  # 제목(+부제) 아래 헤더 구분선 y
BODY_TOP = 1.72  # 본문 시작 y (헤더 이중룰 바로 아래 — 밀도 확보)
SRC_Y = EMU_H - 0.72  # 출처선 y (푸터 위)
FOOTNOTE_Y = EMU_H - 0.62  # 각주 y (푸터 헤어라인 위)
PAGE_W = 2.0  # 페이지번호 박스 폭 (우측 정렬)


def load_brand(path):
    b = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    b.setdefault("brand", {}).setdefault("name", "My Company")
    return b


def rgb(hex6):
    return RGBColor.from_string(hex6)


class Deck:
    """brand-kit을 들고 슬라이드 타입별 렌더러를 제공하는 빌더."""

    def __init__(self, brand):
        self.b = brand
        self.c = brand["colors"]
        self.f = brand["fonts"]
        self.s = brand["sizes"]
        self.margin = brand["layout"]["margin"]
        self.logo = brand["layout"].get("logo") or ""
        self.prs = Presentation()
        self.prs.slide_width = Inches(EMU_W)
        self.prs.slide_height = Inches(EMU_H)
        self.blank = self.prs.slide_layouts[6]
        self._toc_items = []  # 목차용: 본문 섹션 제목 수집

    # --- 저수준 헬퍼 ---
    def _slide(self, bg=None):
        s = self.prs.slides.add_slide(self.blank)
        fill = s.background.fill
        fill.solid()
        fill.fore_color.rgb = rgb(bg or self.c["bg"])
        return s

    def _text(
        self,
        slide,
        x,
        y,
        w,
        h,
        text,
        size,
        color,
        *,
        font=None,
        bold=False,
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,
    ):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = font or self.f["body"]
        run.font.color.rgb = rgb(color)
        return tb

    @staticmethod
    def _parse_emph(text):
        """`**x**` 구간을 강조(True)로 분리. BCG식 인라인 데이터 강조용."""
        parts, i = [], 0
        for mt in re.finditer(r"\*\*(.+?)\*\*", text):
            if mt.start() > i:
                parts.append((text[i : mt.start()], False))
            parts.append((mt.group(1), True))
            i = mt.end()
        if i < len(text):
            parts.append((text[i:], False))
        return parts or [(text, False)]

    def _rich_text(
        self, slide, x, y, w, h, text, size, color, *, font=None, align=PP_ALIGN.LEFT
    ):
        """본문 텍스트 — `**핵심**`은 그 run만 bold + accent색으로. 나머지는 color."""
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        for seg, emph in self._parse_emph(text):
            run = p.add_run()
            run.text = seg
            run.font.size = Pt(size)
            run.font.name = font or self.f["body"]
            run.font.bold = bool(emph)
            run.font.color.rgb = rgb(self.c["accent"] if emph else color)
        return tb

    def _accent_bar(self, slide, x, y, w=2.2, h=0.09, color=None):
        from pptx.enum.shapes import MSO_SHAPE

        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = rgb(color or self.c["accent"])
        bar.line.fill.background()
        return bar

    def _hline(self, slide, x1, x2, y, color, weight=0.75):
        """가는 수평 구분선(헤어라인)."""
        from pptx.enum.shapes import MSO_CONNECTOR

        ln = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y), Inches(x2), Inches(y)
        )
        ln.line.color.rgb = rgb(color)
        ln.line.width = Pt(weight)
        return ln

    def _title_block(self, slide, title, subtitle=None):
        """헤더(기관 문서형): 좌정렬 액션 제목 + 설명 부제 + 이중룰(굵은선+얇은선).
        BCG 실측: 제목 24pt, 부제 ~15pt. AI티 나는 '제목 옆 세로 accent 바' 미사용."""
        m = self.margin
        self._text(
            slide,
            m,
            0.32,
            EMU_W - 2 * m,
            0.6,
            title,
            self.s["section"],
            self.c["primary"],
            font=self.f["head"],
            bold=True,
        )
        if subtitle:
            self._text(
                slide,
                m,
                0.86,
                EMU_W - 2 * m,
                0.45,
                subtitle,
                self.s.get("sub", 15),
                self.c["muted"],
                font=self.f["body"],
            )
        self._hline(slide, m, EMU_W - m, HEADER_HAIR_Y, self.c["primary"], weight=2.5)
        self._hline(
            slide, m, EMU_W - m, HEADER_HAIR_Y + 0.05, self.c["muted"], weight=0.75
        )

    def _footnotes(self, slide, notes):
        """각주(7~8pt) — 참조·단서를 푸터 헤어라인 위에. BCG식 하단 각주."""
        txt = (
            "   ".join(f"{i}. {n}" for i, n in enumerate(notes, 1))
            if len(notes) > 1
            else str(notes[0])
        )
        self._text(
            slide,
            self.margin,
            FOOTNOTE_Y,
            EMU_W - 2 * self.margin,
            0.3,
            txt,
            self.s.get("foot", 8),
            self.c["muted"],
            font=self.f["body"],
        )

    def _frame(self, slide, page_no, total):
        """마스터 틀(기관 문서형 푸터): 헤어라인 + 좌 문서명 + 중앙 브랜드 워드마크
        + 우하단 페이지번호 + 우측 세로 저작권선. 상수 좌표라 전 장 동일 위치 스탬프."""
        m = self.margin
        self._hline(slide, m, EMU_W - m, HAIR_Y, self.c["muted"], weight=0.5)
        # 좌: 문서/프로젝트명
        self._text(
            slide,
            m,
            FOOT_Y,
            4.0,
            0.32,
            self.meta.get("project", ""),
            self.s["caption"],
            self.c["muted"],
        )
        # 중앙: 브랜드 워드마크
        self._text(
            slide,
            EMU_W / 2 - 3,
            FOOT_Y,
            6.0,
            0.32,
            self.b["brand"]["name"],
            self.s["caption"],
            self.c["muted"],
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        # 우: 페이지번호
        self._text(
            slide,
            EMU_W - m - PAGE_W,
            FOOT_Y,
            PAGE_W,
            0.32,
            f"{page_no:02d} / {total:02d}",
            self.s["caption"],
            self.c["muted"],
            align=PP_ALIGN.RIGHT,
        )
        self._vcopyright(slide)

    def _vcopyright(self, slide):
        """우측 여백에 세로로 흐르는 저작권선 — 실제 기관 덱의 대표 시그니처."""
        yr = self.meta.get("year", "")
        txt = f"© {yr} {self.b['brand']['name']}. All rights reserved.".replace(
            "©  ", "© "
        )
        tb = slide.shapes.add_textbox(
            Inches(EMU_W - 2.6), Inches(3.55), Inches(4.6), Inches(0.28)
        )
        tb.rotation = 270
        tf = tb.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = txt
        run.font.size = Pt(7)
        run.font.name = self.f["body"]
        run.font.color.rgb = rgb(self.c["muted"])

    def _source(self, slide, text):
        """출처선: 정량 슬라이드의 좌하단(푸터 위). 셀링 신뢰의 근거 표기."""
        self._text(
            slide,
            self.margin,
            SRC_Y,
            EMU_W - 2 * self.margin,
            0.28,
            f"출처: {text}",
            self.s["caption"],
            self.c["muted"],
        )

    def _logo_mark(self, slide, color=None):
        txt = self.b["brand"]["name"]
        self._text(
            slide,
            self.margin,
            EMU_H - 0.55,
            5,
            0.4,
            txt,
            self.s["caption"],
            color or self.c["muted"],
            bold=True,
        )

    # --- 슬라이드 타입별 렌더러 ---
    def cover(self, spec):
        s = self._slide(bg=self.c["primary"])
        m = self.margin
        self._accent_bar(s, m, 3.0, w=1.6, h=0.12)
        self._text(
            s,
            m,
            3.25,
            EMU_W - 2 * m,
            1.6,
            spec["title"],
            self.s["title"],
            "FFFFFF",
            font=self.f["head"],
            bold=True,
        )
        if spec.get("subtitle"):
            self._text(
                s,
                m,
                4.7,
                EMU_W - 2 * m,
                1.0,
                spec["subtitle"],
                self.s["head"],
                self.c["bg_alt"],
                font=self.f["body"],
            )
        self._logo_mark(s, color="FFFFFF")
        return s

    def toc(self, spec):
        s = self._slide()
        self._title_block(s, spec.get("title", "목차"))
        items = spec.get("items") or self._toc_items
        m = self.margin
        y = 2.0
        for i, it in enumerate(items, 1):
            self._text(
                s,
                m + 0.1,
                y,
                0.6,
                0.5,
                f"{i:02d}",
                self.s["head"],
                self.c["accent"],
                font=self.f["head"],
                bold=True,
            )
            self._text(
                s,
                m + 0.9,
                y,
                EMU_W - 2 * m - 0.9,
                0.5,
                it,
                self.s["body"],
                self.c["text"],
            )
            y += 0.62
        return s

    def section(self, spec):
        s = self._slide(bg=self.c["bg_alt"])
        m = self.margin
        self._accent_bar(s, m, 3.15, w=1.2, h=0.1)
        self._text(
            s,
            m,
            3.35,
            EMU_W - 2 * m,
            1.2,
            spec["title"],
            self.s["title"],
            self.c["primary"],
            font=self.f["head"],
            bold=True,
        )
        return s

    def bullets(self, spec):
        s = self._slide()
        self._title_block(s, spec["title"], spec.get("subtitle"))
        m = self.margin
        y = BODY_TOP
        for bt in spec.get("bullets", []):
            text = bt if isinstance(bt, str) else bt.get("text", "")
            subs = [] if isinstance(bt, str) else bt.get("sub", [])
            self._text(s, m, y, 0.3, 0.35, "•", self.s["body"], self.c["muted"])
            self._rich_text(
                s,
                m + 0.3,
                y,
                EMU_W - 2 * m - 0.3,
                0.5,
                text,
                self.s["body"],
                self.c["text"],
            )
            y += 0.44
            for sub in subs:  # 2단계 중첩 (BCG식 – 서브불릿)
                self._text(
                    s, m + 0.55, y, 0.3, 0.3, "–", self.s["body"] - 1, self.c["muted"]
                )
                self._rich_text(
                    s,
                    m + 0.85,
                    y,
                    EMU_W - 2 * m - 0.85,
                    0.4,
                    sub,
                    self.s["body"] - 1,
                    self.c["text"],
                )
                y += 0.36
            y += 0.08
        return s

    def two_column(self, spec):
        s = self._slide()
        self._title_block(s, spec["title"], spec.get("subtitle"))
        m = self.margin
        colw = (EMU_W - 2 * m - 0.6) / 2
        for idx, key in enumerate(("left", "right")):
            col = spec.get(key, {})
            x = m + idx * (colw + 0.6)
            self._text(
                s,
                x,
                1.72,
                colw,
                0.4,
                col.get("heading", ""),
                self.s["head"],
                self.c["primary"],
                font=self.f["head"],
                bold=True,
            )
            self._hline(s, x, x + colw, 2.16, self.c["primary"], weight=1.5)
            y = 2.34
            for it in col.get("items", []):
                text = it if isinstance(it, str) else it.get("text", "")
                subs = [] if isinstance(it, str) else it.get("sub", [])
                self._text(s, x, y, 0.3, 0.35, "•", self.s["body"], self.c["muted"])
                self._rich_text(
                    s,
                    x + 0.28,
                    y,
                    colw - 0.28,
                    0.5,
                    text,
                    self.s["body"],
                    self.c["text"],
                )
                y += 0.44
                for sub in subs:
                    self._text(
                        s,
                        x + 0.5,
                        y,
                        0.3,
                        0.3,
                        "–",
                        self.s["body"] - 1,
                        self.c["muted"],
                    )
                    self._rich_text(
                        s,
                        x + 0.8,
                        y,
                        colw - 0.8,
                        0.4,
                        sub,
                        self.s["body"] - 1,
                        self.c["text"],
                    )
                    y += 0.34
                y += 0.06
        return s

    def metrics(self, spec):
        s = self._slide()
        self._title_block(s, spec["title"], spec.get("subtitle"))
        items = spec.get("items", [])[:4]
        m = self.margin
        n = max(len(items), 1)
        colw = (EMU_W - 2 * m) / n
        for i, it in enumerate(items):
            x = m + i * colw
            # 숫자는 primary(잉크) — hero 오렌지 카드 대신 절제된 문서형
            self._text(
                s,
                x,
                2.6,
                colw,
                1.1,
                str(it.get("value", "")),
                self.s["title"],
                self.c["primary"],
                font=self.f["head"],
                bold=True,
                align=PP_ALIGN.CENTER,
            )
            # 값 아래 짧은 accent 밑줄 — 포인트는 여기 극소량만
            self._hline(
                s, x + colw * 0.32, x + colw * 0.68, 3.78, self.c["accent"], weight=2.0
            )
            self._text(
                s,
                x,
                3.9,
                colw,
                0.7,
                it.get("label", ""),
                self.s["body"],
                self.c["muted"],
                align=PP_ALIGN.CENTER,
            )
        return s

    def table(self, spec):
        s = self._slide()
        self._title_block(s, spec["title"], spec.get("subtitle"))
        headers = spec["headers"]
        rows = spec["rows"]
        m = self.margin
        nrows, ncols = len(rows) + 1, len(headers)
        gf = s.shapes.add_table(
            nrows,
            ncols,
            Inches(m),
            Inches(2.0),
            Inches(EMU_W - 2 * m),
            Inches(0.4 * nrows),
        )
        tbl = gf.table
        for j, h in enumerate(headers):
            cell = tbl.cell(0, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(self.c["primary"])
            self._cell_text(cell, h, self.c["bg"], bold=True)
        for i, row in enumerate(rows, 1):
            for j, val in enumerate(row):
                cell = tbl.cell(i, j)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(
                    self.c["bg"] if i % 2 else self.c["bg_alt"]
                )
                self._cell_text(cell, str(val), self.c["text"])
        return s

    def _cell_text(self, cell, text, color, bold=False):
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = text
        run.font.size = Pt(self.s["body"] - 3)
        run.font.bold = bold
        run.font.name = self.f["body"]
        run.font.color.rgb = rgb(color)

    def chart(self, spec):
        s = self._slide()
        self._title_block(s, spec["title"], spec.get("subtitle"))
        ct = {
            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "line": XL_CHART_TYPE.LINE,
            "pie": XL_CHART_TYPE.PIE,
        }.get(spec.get("chart_type", "bar"), XL_CHART_TYPE.COLUMN_CLUSTERED)
        data = CategoryChartData()
        data.categories = spec["categories"]
        for name, values in spec["series"].items():
            data.add_series(name, values)
        m = self.margin
        gf = s.shapes.add_chart(
            ct, Inches(m), Inches(2.0), Inches(EMU_W - 2 * m), Inches(4.6), data
        )
        ch = gf.chart
        ch.has_legend = len(spec["series"]) > 1 or ct == XL_CHART_TYPE.PIE
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.BOTTOM
            ch.legend.include_in_layout = False
        return s

    def cta(self, spec):
        s = self._slide(bg=self.c["primary"])
        m = self.margin
        self._accent_bar(s, m, 2.7, w=1.6, h=0.12)
        self._text(
            s,
            m,
            2.95,
            EMU_W - 2 * m,
            1.2,
            spec["title"],
            self.s["title"],
            "FFFFFF",
            font=self.f["head"],
            bold=True,
        )
        contact = spec.get("contact", {})
        lines = [f"{k}: {v}" for k, v in contact.items() if v]
        self._text(
            s,
            m,
            4.4,
            EMU_W - 2 * m,
            1.5,
            "\n".join(lines),
            self.s["body"],
            self.c["bg_alt"],
        )
        self._logo_mark(s, color="FFFFFF")
        return s

    RENDERERS = {
        "cover": cover,
        "toc": toc,
        "section": section,
        "bullets": bullets,
        "two_column": two_column,
        "metrics": metrics,
        "table": table,
        "chart": chart,
        "cta": cta,
    }

    def build(self, spec):
        self.meta = spec.get("meta", {})
        # 목차 자동 채움: 본문 섹션(section/bullets/two_column/table/chart/metrics) 제목 수집
        body_types = {"section", "bullets", "two_column", "table", "chart", "metrics"}
        self._toc_items = [
            sl["title"]
            for sl in spec["slides"]
            if sl["type"] in body_types and sl.get("title")
        ]
        total = len(spec["slides"])
        for i, sl in enumerate(spec["slides"], start=1):
            renderer = self.RENDERERS.get(sl["type"])
            if renderer is None:
                raise ValueError(f"unknown slide type: {sl['type']}")
            slide = renderer(self, sl)
            # 마스터 틀: 표지·CTA를 제외한 전 장에 푸터+페이지번호 스탬프
            if sl["type"] not in ("cover", "cta"):
                self._frame(slide, i, total)
            if sl.get("source"):
                self._source(slide, sl["source"])
            if sl.get("footnotes"):
                self._footnotes(slide, sl["footnotes"])
        return self.prs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("out")
    ap.add_argument(
        "--brand",
        default=str(Path(__file__).parent.parent / "assets" / "brand-kit.yaml"),
    )
    args = ap.parse_args()

    brand = load_brand(args.brand)
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    deck = Deck(brand)
    prs = deck.build(spec)
    prs.save(args.out)
    print(f"OK: {args.out} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    sys.exit(main())
