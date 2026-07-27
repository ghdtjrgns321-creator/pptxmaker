#!/usr/bin/env python
"""익스히빗 사양서 목업 갤러리 생성기 — 사용자 승인 게이트(파이프라인 ②.5)의 산출물.

입력: 후보 사양서 JSON — deck-spec 슬라이드에 visual_candidates[] 를 붙인 형태
  {"slides": [{"no": 7, "title": "...", "message": "so-what 한 문장",
               "visual_candidates": [<chart/diagram spec>, ...], "ref": "ref/p18.png"}]}
출력: <out_dir>/sNN_a.png … + gallery.html (후보마다 편집 가능성 라벨 명시)

사용자는 브라우저에서 비교 후 "s07=B" 식으로 회신 → deck-composer가 deck-spec 확정.
usage: python make_mockups.py <candidates.json> <out_dir>
"""

import html
import json
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import yaml  # noqa: E402
from matplotlib.patches import FancyArrow, Rectangle  # noqa: E402

sys.path.insert(0, str(Path(__file__).parent))
import mpl_exhibits as mx  # noqa: E402

NATIVE_TYPES = {"bar", "hbar", "line", "pie", "doughnut", "stacked_bar"}
LABELS = {
    "native": "네이티브 차트 — PPT에서 수치편집가능",
    "image": "matplotlib 이미지 — 수치편집불가(재빌드로 수정)",
    "shape": "pptx 도형 — PPT에서 편집가능",
}


def edit_label(cand):
    if cand.get("type") == "diagram" or cand.get("layout"):
        return LABELS["shape"]
    if cand.get("type") in ("table", "bullets", "two_column", "metrics"):
        return "네이티브 표/텍스트 — PPT에서 편집가능"
    ct = cand.get("chart_type", "")
    if ct in mx.MPL_TYPES or cand.get("render") == "image":
        return LABELS["image"]
    return LABELS["native"]


def _native_mock(ax, spec, pal):
    """네이티브 유형의 mpl 근사 — 실데이터로 룩을 미리 보여준다."""
    ct = spec["chart_type"]
    cats = spec.get("categories", [])
    series = spec.get("series", {})
    first = list(series.values())[0] if series else []
    colors = mx._emph_colors(cats, spec.get("emphasis"), pal)
    if ct == "bar":
        ax.bar(cats, first, color=colors, width=0.6)
    elif ct == "hbar":
        ax.barh(cats, first, color=colors, height=0.6)
        ax.invert_yaxis()
    elif ct == "line":
        for i, (name, vals) in enumerate(series.items()):
            ax.plot(
                cats,
                vals,
                marker="o",
                lw=1.8,
                color=pal["accent"] if i == 0 else pal["gray"],
                label=name,
            )
        if len(series) > 1:
            ax.legend(frameon=False, fontsize=8)
    elif ct in ("pie", "doughnut"):
        n = len(cats)
        cs = [
            pal["accent"] if i == 0 else [pal["gray"], pal["gray_dark"], pal["bg_alt"]][i % 3]
            for i in range(n)
        ]
        wedge = dict(width=0.42) if ct == "doughnut" else {}
        ax.pie(
            first,
            labels=cats,
            colors=cs,
            wedgeprops=wedge,
            textprops=dict(fontsize=8, color=pal["ink"]),
        )
    elif ct == "stacked_bar":
        bottom = [0.0] * len(cats)
        shades = [pal["accent"], pal["gray_dark"], pal["gray"], pal["bg_alt"]]
        for i, (name, vals) in enumerate(series.items()):
            ax.bar(cats, vals, bottom=bottom, color=shades[i % 4], width=0.6, label=name)
            bottom = [b + v for b, v in zip(bottom, vals)]
        ax.legend(frameon=False, fontsize=8)


def _table_mock(ax, spec, pal):
    """표 스케치 — 헤더 행 강조 + 실데이터 셀(matplotlib table)."""
    ax.axis("off")
    headers = [str(h) for h in spec.get("headers", [])]
    rows = [[str(v) for v in r] for r in spec.get("rows", [])][:6]
    tbl = ax.table(cellText=rows, colLabels=headers, loc="center", cellLoc="left")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(7)
    matrix = spec.get("style") == "matrix"
    ncols = len(headers)
    for (r, c), cell in tbl.get_celld().items():
        cell.set_edgecolor("#DDDDDD")
        if r == 0:
            cell.set_facecolor(pal["accent"] if (matrix and c == ncols - 1) else pal["primary"])
            cell.get_text().set_color("white")
        elif r % 2 == 0:
            cell.set_facecolor(pal["bg_alt"])


def _text_mock(ax, spec, pal):
    """bullets/two_column/metrics 스케치 — 텍스트 요약(형식 감만 전달)."""
    ax.axis("off")
    t = spec.get("type", "")
    if t == "metrics":
        items = spec.get("items", [])[:4]
        for i, it in enumerate(items):
            x = (i + 0.5) / max(len(items), 1)
            ax.text(
                x,
                0.62,
                str(it.get("value", "")),
                ha="center",
                fontsize=17,
                fontweight="bold",
                color=pal["primary"],
                transform=ax.transAxes,
            )
            ax.text(
                x,
                0.38,
                str(it.get("label", ""))[:16],
                ha="center",
                fontsize=7.5,
                color=pal["gray_dark"],
                transform=ax.transAxes,
            )
        return
    lines = []
    if t == "two_column":
        left, right = spec.get("left", {}), spec.get("right", {})
        lines.append(f"[{left.get('heading', '좌')}]  vs  [{right.get('heading', '우')}]")
        for it in (left.get("items") or [])[:3]:
            lines.append("• " + (it if isinstance(it, str) else it.get("text", ""))[:38])
    else:
        for it in (spec.get("bullets") or [])[:5]:
            lines.append("• " + (it if isinstance(it, str) else it.get("text", ""))[:44])
    for i, ln in enumerate(lines):
        ax.text(0.04, 0.9 - i * 0.17, ln, fontsize=8.5, color=pal["ink"], transform=ax.transAxes)


def _diagram_mock(ax, spec, pal):
    """다이어그램 스케치 — 노드 라벨을 박스로 배치(레이아웃 감만 전달).

    2026-07-25 재작성: 옛 분기는 폐기된 박스 어휘 이름(`process_band`·`layers`·`cards`·`branch`·
    `from_to`)으로 갈라졌고, **남은 9종 중 timeline 말고는 어느 이름도 안 걸려 빈 스케치가
    나왔다**(matrix_2x2·venn·icon_rows…). 목업은 "이 자리에 이런 덩치가 온다"만 전하면 되므로
    이름 분기를 버리고 **가로 진행(timeline) vs 세로/격자(그 외)** 둘로 줄인다."""
    layout = spec.get("layout")
    rows = spec.get("rows") or []
    nodes = spec.get("nodes") or [
        {"label": r.get("lead") or r.get("label") or r.get("text") or ""}
        for r in rows
        if isinstance(r, dict)
    ]
    n = max(len(nodes), 1)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis("off")

    def box(x, y, w, h, label, fc):
        ax.add_patch(Rectangle((x, y), w, h, fc=fc, ec=pal["primary"], lw=1))
        ax.text(
            x + w / 2,
            y + h / 2,
            label,
            ha="center",
            va="center",
            fontsize=7.5,
            color=pal["ink"],
            wrap=True,
        )

    if layout == "timeline":  # 가로 진행 — 시간축이 유일한 진행 어휘로 남았다
        w = 10 / n - 0.5
        for i, nd in enumerate(nodes):
            box(i * (w + 0.5), 2, w, 1.4, nd["label"], pal["bg_alt"])
            if i:
                ax.add_patch(
                    FancyArrow(i * (w + 0.5) - 0.45, 2.7, 0.32, 0, width=0.06, color=pal["accent"])
                )
    elif n <= 4:  # 세로 행 — icon_rows·stat_split·spectrum 등 행 기반 어휘의 덩치
        h = 4.5 / n - 0.25
        for i, nd in enumerate(nodes):
            box(1.5, 4.5 - (i + 1) * (h + 0.25), 7, h, nd["label"], pal["bg_alt"])
    else:  # 격자 — 셀 기반 어휘(harvey_table·check_matrix·matrix_2x2)
        cols = 3
        w, h = 10 / cols - 0.6, 1.6
        for i, nd in enumerate(nodes):
            box(
                (i % cols) * (w + 0.6),
                3.2 - (i // cols) * (h + 0.4),
                w,
                h,
                nd["label"],
                pal["bg_alt"],
            )


def render_candidate(cand, brand, out_png):
    """후보 1안 → 목업 PNG. 차트는 실데이터, 다이어그램은 스케치."""
    ct = cand.get("chart_type", "")
    if ct in mx.MPL_TYPES:
        return mx.render(cand, brand, out_png, w_in=5.2, h_in=2.9)
    fig, ax, pal = mx._fig(brand, 5.2, 2.9)
    t = cand.get("type", "")
    if ct in NATIVE_TYPES:
        _native_mock(ax, cand, pal)
    elif t == "table":
        _table_mock(ax, cand, pal)
    elif t in ("bullets", "two_column", "metrics"):
        _text_mock(ax, cand, pal)
    else:
        _diagram_mock(ax, cand, pal)
    return mx._save(fig, ax, cand, pal, out_png)


CSS = """body{font-family:Pretendard,'Malgun Gothic',sans-serif;background:#F2F3F5;margin:24px;color:#1F2329}
h1{font-size:20px} .slide{background:#fff;border:1px solid #ddd;padding:16px;margin-bottom:20px}
.msg{color:#555;margin:4px 0 12px} .cands{display:flex;gap:14px;flex-wrap:wrap}
.card{border:1px solid #ccc;padding:10px;width:420px} .card img{width:100%}
.badge{display:inline-block;font-size:11px;padding:2px 8px;margin-bottom:6px;border-radius:3px}
.native{background:#e8f0e8;color:#2c6e2c}.image{background:#fdeee4;color:#B14D1F}.shape{background:#e9edf5;color:#2c4a6e}
.pick{font-weight:700;font-size:14px;color:#D66E3A} .ref{font-size:11px;color:#888}"""


def main():
    cand_path, out_dir = Path(sys.argv[1]), Path(sys.argv[2])
    out_dir.mkdir(parents=True, exist_ok=True)
    brand_path = Path(__file__).resolve().parents[2] / "pptx-build" / "assets" / "brand-kit.yaml"
    brand = yaml.safe_load(brand_path.read_text(encoding="utf-8"))
    data = json.loads(cand_path.read_text(encoding="utf-8"))
    parts = [
        f"<style>{CSS}</style><h1>익스히빗 사양서 — 슬라이드별 시각화 후보</h1>",
        "<p>각 슬라이드에서 원하는 안을 골라 <b>s07=B</b> 형식으로 회신하세요.</p>",
    ]
    for sl in data["slides"]:
        no = sl["no"]
        parts.append(f'<div class="slide"><h2>s{no:02d}. {html.escape(sl.get("title", ""))}</h2>')
        parts.append(f'<div class="msg">주장: {html.escape(sl.get("message", ""))}</div>')
        parts.append('<div class="cands">')
        for i, cand in enumerate(sl.get("visual_candidates", [])):
            tag = chr(ord("A") + i)
            png = out_dir / f"s{no:02d}_{tag.lower()}.png"
            render_candidate(cand, brand, png)
            label = edit_label(cand)
            cls = (
                "image"
                if "수치편집불가" in label
                else ("native" if "네이티브" in label else "shape")
            )
            kind = cand.get("chart_type") or f"diagram {cand.get('layout', '')}"
            ref = (
                f'<div class="ref">레퍼런스: {html.escape(cand["ref"])}</div>'
                if cand.get("ref")
                else ""
            )
            parts.append(
                f'<div class="card"><span class="pick">{tag}</span> — {html.escape(kind)}<br>'
                f'<span class="badge {cls}">{label}</span>{ref}'
                f'<img src="{png.name}" alt="{tag}"></div>'
            )
        parts.append("</div></div>")
    out = out_dir / "gallery.html"
    out.write_text("\n".join(parts), encoding="utf-8")
    print(f"OK: {out}")


if __name__ == "__main__":
    main()
