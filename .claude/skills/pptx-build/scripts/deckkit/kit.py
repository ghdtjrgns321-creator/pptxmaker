"""deckkit — 토큰·원시도구·페이지 골격. 형태는 정하지 않는다.

## 왜 이 모듈이 생겼나 (2026-08-03)

하네스를 못 보는 새 컨텍스트에 계약만 주고 같은 덱을 만들게 한 실험에서, 하네스로 만든 것보다
조밀하고(도형 70 대 50/장) 수치도 정확한 산출물이 나왔다. 사용자가 그쪽을 택했고, 골든덱은
기준선에서 내렸다. 이 모듈은 그 산출물에서 **토큰과 골격만** 뽑은 것이다.

## 무엇을 소유하고 무엇을 소유하지 않나

소유한다 — 색·글꼴 위계·여백(`brand.yaml`), 텍스트/사각/화살표/막대/이미지, 표지·간지·헤더·
푸터·패널. 즉 **일관성이 필요한 것**.

소유하지 않는다 — 장을 어떻게 구성할지. 그건 매번 내용에 맞게 정한다. 부품이 장 구성을 정하면
전 장이 같은 형식이 된다(2026-07-29 실측: 36장이 동일한 4카드).
"""

import math
import os
import struct

import yaml
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

_B = yaml.safe_load((__import__("pathlib").Path(__file__).parent / "brand.yaml").read_text("utf-8"))

COLORS = _B["colors"]
SIZES = _B["sizes"]
FONT = _B["font"]
FR = _B["frame"]
LIM = _B["limits"]

INK = COLORS["ink"]
# 구조는 근흑이 지고 오렌지는 포인트로만 — 무채색 + 포인트 하나(Muted Ember).
# NAVY/BLUE/AMBER 이름은 호출부 호환을 위해 남겼다. 값은 브랜드 토큰이 단일 출처다.
NAVY, NAVY_LT = COLORS["dark"], COLORS["dark_lt"]
BLUE, BLUE_LT = COLORS["struct"], COLORS["struct_lt"]
AMBER, AMBER_LT = COLORS["accent"], COLORS["accent_lt"]
GRAY, LINE, PANEL, WHITE = COLORS["muted"], COLORS["line"], COLORS["panel"], COLORS["white"]
DARK, DARK_LT = NAVY, NAVY_LT       # 뜻이 드러나는 이름 — 새 코드는 이쪽을 쓴다
STRUCT, STRUCT_LT = BLUE, BLUE_LT
ACCENT, ACCENT_LT = AMBER, AMBER_LT

SW, SH = FR["slide_w"], FR["slide_h"]
M, TOP, BOT = FR["margin"], FR["top"], FR["bottom"]
CW = SW - 2 * M
RIGHT = SW - M

_ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
_ANCH = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}

# 조판 중 쌓이는 경고 — 장 스크립트가 끝나고 읽는다
WARN: list[str] = []
CUR = {"n": 0, "title": ""}
IMG_DIR = ""


def set_context(n: int, title: str = "", img_dir: str = ""):
    CUR["n"], CUR["title"] = n, title
    if img_dir:
        global IMG_DIR
        IMG_DIR = img_dir


def C(h: str) -> RGBColor:
    return RGBColor.from_string(h)


def In(v: float) -> Inches:
    return Inches(v)


def _em(txt: str) -> float:
    """문자열의 대략적 폭(em). 전각 1.0 · 영숫자 0.55 · 공백 0.35."""
    w = 0.0
    for ch in txt:
        w += 0.35 if ch == " " else (1.0 if ord(ch) > 0x2000 else 0.55)
    return w


def _font(run, size, bold, color, italic=False):
    """라틴·동아시아·복합 자형을 모두 같은 글꼴로 — 하나라도 빠지면 한글이 다른 꼴로 나온다."""
    f = run.font
    f.size, f.bold, f.italic = Pt(size), bold, italic
    f.color.rgb = C(color)
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(f"{{http://schemas.openxmlformats.org/drawingml/2006/main}}{tag[2:]}")
        if el is None:
            el = rPr.makeelement(
                f"{{http://schemas.openxmlformats.org/drawingml/2006/main}}{tag[2:]}", {}
            )
            rPr.append(el)
        el.set("typeface", FONT)


def T(slide, x, y, w, h, lines, valign="t", pad=0.03, wrap=True, name=""):
    """텍스트 블록. lines = [{t, s, b, c, a, sb, sa, ls}] · t는 문자열 또는 런 목록.

    한글은 **어절 단위로만 줄바꿈**한다(`eaLnBrk=0`). 이걸 안 걸면 PowerPoint가 글자 단위로
    끊어 고아 글자가 생기고, 무하네스 실험에서 그걸 손으로 7곳 고쳐야 했다.
    """
    tb = slide.shapes.add_textbox(In(x), In(y), In(w), In(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = In(pad)
    tf.margin_top = tf.margin_bottom = In(0.01)
    tf.vertical_anchor = _ANCH[valign]
    used, inner = 0.0, max(w - 2 * pad, 0.2)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p._p.get_or_add_pPr().set("eaLnBrk", "0")
        p.alignment = _ALIGN[ln.get("a", "l")]
        sb, sa = ln.get("sb", 0), ln.get("sa", 0)
        p.space_before, p.space_after = Pt(sb), Pt(sa)
        ls = ln.get("ls", 1.0)
        p.line_spacing = ls
        t = ln["t"]
        runs = (
            t
            if isinstance(t, list)
            else [(t, ln.get("s", SIZES["body"]), ln.get("b", False), ln.get("c", INK))]
        )
        maxs, em = 0.0, 0.0
        for rt, rs, rb, rc in runs:
            for j, seg in enumerate(str(rt).split("\n")):
                if j:
                    em = math.ceil(em / inner - 1e-6) * inner
                r = p.add_run()
                r.text = ("\n" if j else "") + seg
                _font(r, rs, rb, rc)
                em += _em(seg) * rs / 72.0
            maxs = max(maxs, rs)
        used += max(1, math.ceil(em / inner - 1e-6)) * maxs * 1.22 * ls / 72.0 + (sb + sa) / 72.0
    if used > h + 0.03:
        WARN.append(f'S{CUR["n"]:02d} {name or "텍스트"} 넘침 추정 {used:.2f}" > {h:.2f}"')
    return tb


def R(slide, x, y, w, h, fill=None, line=None, lw=0.75, rounded=0.0):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, In(x), In(y), In(w), In(h)
    )
    if rounded:
        shp.adjustments[0] = rounded
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = C(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = C(line)
        shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    shp.text_frame.word_wrap = True
    if x + w > RIGHT + LIM["right_tolerance"]:
        WARN.append(f'S{CUR["n"]:02d} 우변 초과 {x + w - RIGHT:.2f}" (도형 x={x:.2f} w={w:.2f})')
    return shp


def ARROW(slide, x, y, w, h, color=BLUE):
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, In(x), In(y), In(w), In(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = C(color)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _arrowhead(shp, head=True, tail=False):
    ln = shp.line._get_or_add_ln()
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for tag, on in (("headEnd", tail), ("tailEnd", head)):
        if not on:
            continue
        el = ln.find(f"{ns}{tag}")
        if el is None:
            el = ln.makeelement(f"{ns}{tag}", {})
            ln.append(el)
        el.set("type", "triangle")
        el.set("w", "med")
        el.set("len", "med")


def CONNECT(slide, x1, y1, x2, y2, *, elbow=False, head=True, color=None, w=1.0, dash=False):
    """점과 점을 잇는다 — 직선 또는 꺾은선(elbow), 화살촉 선택.

    **도해 부품을 만들지 않는 이유.** 흐름도·관계도를 `flow()` `graph()` 같은 블록으로 굳히면
    안 맞는 내용에도 그 형식으로 그리게 된다(2026-07-29 실측: 전 장이 같은 형식). 그래서
    잉크만 준다 — 사각(`R`)과 이 커넥터면 머메이드·아스키로 적을 수 있는 노드·간선 도해는
    무엇이든 그릴 수 있고, 배치는 매번 내용이 정한다.

    언제 그릴지는 재료가 말한다 — `scan_material.py`가 절마다 화살표 수·순서 유무를 센다.
    """
    kind = MSO_CONNECTOR.ELBOW if elbow else MSO_CONNECTOR.STRAIGHT
    c = slide.shapes.add_connector(kind, In(x1), In(y1), In(x2), In(y2))
    c.line.color.rgb = C(color or LINE)
    c.line.width = Pt(w)
    if dash:
        c.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    if head:
        _arrowhead(c)
    return c


def _edge(shp, side):
    x, y = Emu(shp.left).inches, Emu(shp.top).inches
    w, h = Emu(shp.width).inches, Emu(shp.height).inches
    return {
        "l": (x, y + h / 2),
        "r": (x + w, y + h / 2),
        "t": (x + w / 2, y),
        "b": (x + w / 2, y + h),
        "c": (x + w / 2, y + h / 2),
    }[side]


def LINK(slide, a, b, *, sides=None, elbow=False, head=True, color=None, w=1.0, dash=False):
    """도형 a → 도형 b를 잇는다. sides=("r","l") 처럼 접점을 지정, 없으면 상대 위치로 고른다."""
    ax, ay = (
        Emu(a.left).inches + Emu(a.width).inches / 2,
        Emu(a.top).inches + Emu(a.height).inches / 2,
    )
    bx, by = (
        Emu(b.left).inches + Emu(b.width).inches / 2,
        Emu(b.top).inches + Emu(b.height).inches / 2,
    )
    if sides is None:
        sides = (
            ("r", "l")
            if abs(bx - ax) >= abs(by - ay) and bx >= ax
            else (
                ("l", "r")
                if abs(bx - ax) >= abs(by - ay)
                else (("b", "t") if by >= ay else ("t", "b"))
            )
        )
    p1, p2 = _edge(a, sides[0]), _edge(b, sides[1])
    return CONNECT(slide, *p1, *p2, elbow=elbow, head=head, color=color, w=w, dash=dash)


def BAR(slide, x, y, w, h, frac, color, track=None):
    """값을 길이로 인코딩. track을 주면 전체 대비가 보인다."""
    if track:
        R(slide, x, y, w, h, fill=track)
    frac = max(0.0, min(1.0, frac))
    if frac > 0:
        R(slide, x, y, max(w * frac, 0.02), h, fill=color)


def _png_size(path):
    with open(path, "rb") as f:
        head = f.read(33)
    return struct.unpack(">II", head[16:24])


def image_ar(fname) -> float:
    """그림의 가로/세로 비율. 칸을 그림에 맞출 때 쓴다."""
    path = fname if os.path.isabs(str(fname)) else os.path.join(IMG_DIR, str(fname))
    iw, ih = _png_size(path)
    return iw / ih


def image_h(fname, w, caption=False, cap_h=0.24) -> float:
    """**폭 w를 다 쓸 때 필요한 칸 높이.**

    칸을 먼저 정하고 그림을 우겨넣으면 비율이 안 맞는 만큼 빈다(2026-08-03 지적: "칸이
    안 맞아서 보기 싫다"). 순서를 뒤집는다 — 폭을 정하고 높이를 **그림에게 물어본다.**
    """
    return w / image_ar(fname) + (cap_h if caption else 0)


def IMAGE(slide, fname, x, y, w, h, caption=None, cap_h=0.24):
    """비율 유지로 칸에 맞춘다. 결과 폭이 판독 하한 미만이면 경고 — 개수를 줄이거나 크롭할 것."""
    path = fname if os.path.isabs(str(fname)) else os.path.join(IMG_DIR, str(fname))
    iw, ih = _png_size(path)
    ar = iw / ih
    box_h = h - (cap_h if caption else 0)
    pw, ph = (box_h * ar, box_h) if w / box_h > ar else (w, w / ar)
    px, py = x + (w - pw) / 2, y + (box_h - ph) / 2
    pic = slide.shapes.add_picture(path, In(px), In(py), In(pw), In(ph))
    pic.line.color.rgb = C(LINE)
    pic.line.width = Pt(0.75)
    if pw < LIM["image_min_w"]:
        WARN.append(
            f'S{CUR["n"]:02d} 이미지 폭 {pw:.2f}" < 판독 하한 {LIM["image_min_w"]}"'
            f" ({os.path.basename(path)}) — 개수를 줄이거나 크롭할 것"
        )
    fill = (pw * ph) / (w * box_h) if w * box_h > 0 else 1.0
    if fill < LIM["image_fill_min"]:
        WARN.append(
            f"S{CUR['n']:02d} 그림이 칸의 {fill:.0%}만 채운다 ({os.path.basename(path)})"
            f' — 칸 {w:.2f}×{box_h:.2f}", 그림 {pw:.2f}×{ph:.2f}".'
            f" `image_h(파일, 폭)`으로 칸 높이를 그림에 맞출 것"
        )
    if caption:
        T(
            slide,
            x,
            y + box_h + 0.02,
            w,
            cap_h,
            [{"t": caption, "s": SIZES["tiny"], "c": GRAY, "a": "c", "ls": 1.1}],
            name="caption",
        )
    return pic


# ── 페이지 골격 ──────────────────────────────────────────────────────────────
def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header(s, n, eyebrow, title, lead=None, footer_text=""):
    set_context(n, title)
    T(
        s,
        M,
        0.34,
        7.0,
        0.26,
        [{"t": eyebrow, "s": SIZES["small"], "b": True, "c": BLUE}],
        name="눈썹",
    )
    T(
        s,
        M - 0.02,
        0.56,
        5.9,
        0.52,
        [{"t": title, "s": SIZES["title"], "b": True, "c": INK}],
        valign="m",
        name="제목",
    )
    if lead:
        T(
            s,
            6.50,
            0.46,
            6.21,
            0.68,
            [{"t": lead, "s": SIZES["body"], "c": GRAY, "a": "r", "ls": 1.2}],
            valign="m",
            name="리드",
        )
    R(s, M, 1.18, CW, 0.018, fill=LINE)
    R(s, M, 1.18, 1.15, 0.018, fill=BLUE)
    footer(s, n, footer_text)


def footer(s, n, text=""):
    R(s, M, 7.00, CW, 0.012, fill=LINE)
    if text:
        T(s, M, 7.06, 8.0, 0.26, [{"t": text, "s": SIZES["tiny"], "c": GRAY}], name="꼬리")
    T(
        s,
        11.0,
        7.06,
        1.71,
        0.26,
        [{"t": f"{n:02d}", "s": SIZES["tiny"], "c": GRAY, "a": "r"}],
        name="쪽번호",
    )


def dark_page(prs, n):
    """어두운 전면 장 — 표지·간지의 바탕.

    배경은 **도형이 아니라 슬라이드 배경**으로 칠한다. 도형이면 폭 13.33"가 우변 경고를 낸다.
    """
    s = blank(prs)
    set_context(n)
    bg = s.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C(NAVY)
    T(
        s,
        11.0,
        7.02,
        1.71,
        0.26,
        [{"t": f"{n:02d}", "s": SIZES["tiny"], "c": GRAY, "a": "r"}],
        name="쪽번호",
    )
    return s


def cover(prs, kicker, title, subtitle=None, lead=None, facts=()):
    """표지. title은 줄바꿈(\\n)으로 행을 나눈다. facts = [(라벨, 값, 부연), ...]

    정형 장이라 덱마다 달라지면 안 된다 — 2026-08-03에 이게 템플릿에 없어서 매번 새로
    발명되고 있었다(그 과정에서 토큰 밖 색도 생겼다).
    """
    s = dark_page(prs, 1)
    R(s, 0, 0, 0.14, SH, fill=ACCENT)
    T(
        s,
        M,
        1.52,
        8.0,
        0.30,
        [{"t": kicker, "s": SIZES["small"], "b": True, "c": ACCENT}],
        name="킥커",
    )
    T(
        s,
        M,
        1.94,
        11.4,
        1.66,
        [
            {"t": ln, "s": SIZES["cover"], "b": True, "c": WHITE, "ls": 1.16}
            for ln in title.split("\n")
        ],
        name="표제",
    )
    y = 3.88
    if subtitle:
        T(
            s,
            M,
            y,
            9.0,
            0.56,
            [{"t": subtitle, "s": SIZES["div"], "b": True, "c": ACCENT}],
            name="부제",
        )
        y += 0.64
    if lead:
        T(
            s,
            M,
            y,
            10.6,
            0.40,
            [{"t": lead, "s": SIZES["lead"], "c": LINE, "ls": 1.25}],
            name="리드",
        )
    if facts:
        n = len(facts)
        gap = 0.24
        fw = (CW - gap * (n - 1)) / n
        fy = 5.42
        for i, (lab, val, sub) in enumerate(facts):
            fx = M + i * (fw + gap)
            R(s, fx, fy, fw, 0.86, fill=NAVY_LT)
            T(
                s,
                fx + 0.20,
                fy + 0.10,
                fw - 0.40,
                0.66,
                [
                    {"t": lab, "s": SIZES["tiny"], "c": GRAY},
                    {"t": val, "s": SIZES["h"], "b": True, "c": WHITE, "sb": 2},
                    {"t": sub, "s": SIZES["tiny"], "c": GRAY, "sb": 2},
                ],
                name=f"표지칩:{lab}",
            )
    return s


def divider(prs, n, roman, title, lead=None, chapters=()):
    """간지. chapters = ["S4 문제 인식", ...] — 그 부에 무엇이 들었는지 미리 보인다."""
    s = dark_page(prs, n)
    R(s, 0, 2.05, 2.30, 0.05, fill=ACCENT)
    T(
        s,
        M,
        2.30,
        3.0,
        0.62,
        [{"t": roman, "s": SIZES["div"], "b": True, "c": ACCENT}],
        name="로마자",
    )
    T(
        s,
        M,
        3.00,
        10.0,
        0.70,
        [{"t": title, "s": SIZES["cover"], "b": True, "c": WHITE}],
        name="간지제목",
    )
    if lead:
        T(
            s,
            M,
            3.86,
            8.4,
            0.44,
            [{"t": lead, "s": SIZES["lead"], "c": LINE, "ls": 1.25}],
            name="간지리드",
        )
    if chapters:
        x = M
        for ch in chapters:
            w = 0.30 + len(ch) * 0.115
            R(s, x, 4.62, w, 0.34, fill=NAVY_LT)
            T(
                s,
                x,
                4.62,
                w,
                0.34,
                [{"t": ch, "s": SIZES["tiny"], "c": LINE, "a": "c"}],
                valign="m",
                name="간지칩",
            )
            x += w + 0.14
    return s


def toc(prs, n, parts, footer_text=""):
    """목차. parts = [(로마자, 부 제목, "S4 · S5", 요약), ...] — 부마다 한 행."""
    s = blank(prs)
    header(s, n, "CONTENTS", "목차", f"{len(parts)}부 구성", footer_text)
    rh = (BOT - TOP) / len(parts)
    for i, (roman, title, span, summary) in enumerate(parts):
        y = TOP + i * rh
        R(s, M, y + rh - 0.02, CW, 0.012, fill=LINE)
        R(s, M, y + 0.10, 0.52, 0.52, fill=BLUE)
        T(
            s,
            M,
            y + 0.10,
            0.52,
            0.52,
            [{"t": roman, "s": SIZES["h"], "b": True, "c": WHITE, "a": "c"}],
            valign="m",
            name="부번호",
        )
        T(
            s,
            M + 0.74,
            y,
            4.4,
            rh - 0.06,
            [{"t": title, "s": SIZES["title"] - 6, "b": True, "c": INK}],
            valign="m",
            name=f"부:{title}",
        )
        T(
            s,
            M + 5.30,
            y,
            5.6,
            rh - 0.06,
            [{"t": summary, "s": SIZES["small"], "c": GRAY, "ls": 1.2}],
            valign="m",
            name="부요약",
        )
        T(
            s,
            M + CW - 1.5,
            y,
            1.5,
            rh - 0.06,
            [{"t": span, "s": SIZES["small"], "b": True, "c": BLUE, "a": "r"}],
            valign="m",
            name="쪽범위",
        )
    return s


class Panel:
    """소제목으로 가둔 구역. **세로 슬롯을 소유한다.**

    `row(h)`로 아래로 쌓으면 폭이 구조적으로 같아진다 — 쌓은 블록의 좌우가 어긋나는 사고
    (2026-08-03 지적)를 규칙이 아니라 구조로 막는다.
    """

    def __init__(self, s, x, y, w, h, title=None, sub=None, fill=PANEL, accent=BLUE, pad=0.18):
        R(s, x, y, w, h, fill=fill, line=LINE)
        self.s, self.x, self.y, self.w, self.h, self.pad = s, x, y, w, h, pad
        head = 0.0
        if title:
            R(s, x, y, 0.055, h, fill=accent)
            lines = [{"t": title, "s": SIZES["h"], "b": True, "c": INK}]
            if sub:
                lines.append({"t": sub, "s": SIZES["tiny"], "c": GRAY, "sb": 2})
            head = 0.34 + (0.20 if sub else 0)
            T(s, x + pad, y + 0.08, w - 2 * pad, head, lines, name="패널 제목")
        self._cur = y + (0.08 + head + 0.06 if title else pad)

    @property
    def inner_x(self):
        return self.x + self.pad

    @property
    def inner_w(self):
        return self.w - 2 * self.pad

    @property
    def left(self):
        return self._cur

    def row(self, h, gap=0.06):
        """다음 슬롯 (x, y, w, h)를 준다 — x·w는 패널이 정하므로 어긋날 수 없다."""
        y = self._cur
        self._cur = y + h + gap
        if self._cur > self.y + self.h + 0.02:
            WARN.append(f"S{CUR['n']:02d} 패널 세로 초과 — 슬롯이 패널 밖으로 나갔다")
        return self.inner_x, y, self.inner_w, h
