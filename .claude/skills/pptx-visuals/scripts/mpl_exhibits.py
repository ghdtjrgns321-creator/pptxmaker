"""BCG풍 복합 익스히빗 렌더러 — matplotlib 고해상도 PNG (pptx-visuals 스킬 단일 출처).

네이티브 차트(visuals.py)가 표현 못 하는 유형만 담당한다: waterfall·heatmap·dumbbell·
slope·funnel·annotated_scatter·histogram. 산출물은 이미지라 PPT에서 수치 편집이 불가하며,
수정은 deck-spec 경유 재빌드가 공식 경로다(목업 갤러리에 "수치편집불가" 라벨로 고지).

스타일 단일 경로: 모든 렌더러는 _fig()로 시작해 _save()로 끝나고, 색은 반드시
_palette(brand)가 주는 회색 베이스 + accent 1색만 쓴다 — "전부 회색, 강조 데이터만
브랜드색 1개 + 콜아웃 주석 + 하단 출처 각주"가 BCG 룩의 최대 지렛대.
"""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from matplotlib.colors import LinearSegmentedColormap  # noqa: E402
from matplotlib.patches import Circle  # noqa: E402

DPI = 220

# 렌더 경로 판별용 — build_pptx가 chart_type이 여기 있으면 이미지 경로로 보낸다
MPL_TYPES = {
    "waterfall",
    "heatmap",
    "dumbbell",
    "slope",
    "funnel",
    "annotated_scatter",
    "histogram",
    "bubble",
    "waffle",
}


def _korean_font():
    """설치된 한글 폰트 중 브랜드 우선순위 첫 매치(글자 깨짐 방지)."""
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in ("Pretendard", "Malgun Gothic", "NanumGothic"):
        if name in installed:
            return name
    return "sans-serif"


def _palette(brand):
    """brand-kit 색을 mpl 팔레트로 — 회색 베이스 + accent 1색 원칙의 단일 출처."""
    c = brand["colors"]
    return {
        "accent": f"#{c['accent']}",
        "ink": f"#{c['text']}",
        "primary": f"#{c['primary']}",
        "muted": f"#{c['muted']}",
        "gray": "#C7CBD1",  # 데이터 베이스 회색(muted보다 밝게 — 강조가 튀도록)
        "gray_dark": f"#{c['muted']}",
        "bg_alt": f"#{c['bg_alt']}",
    }


def _fig(brand, w_in, h_in):
    """공통 진입점 — 폰트·스파인·눈금 스타일을 한곳에서 강제."""
    pal = _palette(brand)
    plt.rcParams.update(
        {
            "font.family": _korean_font(),
            "axes.unicode_minus": False,
            "font.size": 9,
            "text.color": pal["ink"],
            "axes.edgecolor": pal["muted"],
            "axes.labelcolor": pal["ink"],
            "xtick.color": pal["gray_dark"],
            "ytick.color": pal["gray_dark"],
        }
    )
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.tick_params(length=0)
    return fig, ax, pal


def _save(fig, ax, spec, pal, out_path):
    """공통 종점 — 콜아웃 주석 + 각주를 얹고 220dpi PNG 저장."""
    for ann in spec.get("annotations", []):
        _callout(ax, ann, pal)
    if spec.get("note"):
        fig.text(0.01, 0.01, spec["note"], fontsize=7, color=pal["muted"], ha="left")
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_path


def _callout(ax, ann, pal):
    """데이터 좌표(xy)를 지목하는 BCG식 콜아웃 박스 — accent 테두리 + 지시선."""
    xy = ann.get("xy")
    xytext = ann.get("at", (0.98, 0.95))  # axes 분수 좌표
    kw = dict(
        fontsize=8.5,
        color=pal["ink"],
        ha="right",
        va="top",
        bbox=dict(boxstyle="round,pad=0.4", fc="white", ec=pal["accent"], lw=1.2),
    )
    if xy is not None:
        ax.annotate(
            ann["text"],
            xy=xy,
            xycoords="data",
            xytext=xytext,
            textcoords="axes fraction",
            arrowprops=dict(arrowstyle="-", color=pal["accent"], lw=1.0),
            **kw,
        )
    else:
        ax.annotate(ann["text"], xy=xytext, xycoords="axes fraction", **kw)


def render(spec, brand, out_path, w_in=8.0, h_in=4.2):
    """spec['chart_type']에 맞는 렌더러로 PNG 생성, 경로 반환."""
    kind = spec["chart_type"]
    fn = {
        "waterfall": _waterfall,
        "heatmap": _heatmap,
        "dumbbell": _dumbbell,
        "slope": _slope,
        "funnel": _funnel,
        "annotated_scatter": _annotated_scatter,
        "histogram": _histogram,
        "bubble": _bubble,
        "waffle": _waffle,
    }.get(kind)
    if fn is None:
        raise ValueError(f"unknown mpl exhibit type: {kind}")
    fig, ax, pal = _fig(brand, w_in, h_in)
    fn(ax, spec, pal)
    return _save(fig, ax, spec, pal, out_path)


def _emph_colors(labels, emphasis, pal, base=None):
    """emphasis에 든 라벨만 accent, 나머지는 회색 — 강조 1색 원칙."""
    emph = set(emphasis or [])
    return [pal["accent"] if lb in emph else (base or pal["gray"]) for lb in labels]


def _waterfall(ax, spec, pal):
    """증분 누적(구성 분해). values는 증감분, 마지막에 합계 막대 자동 추가."""
    cats = list(spec["categories"])
    vals = list(spec["values"])
    cum, bases = 0.0, []
    for v in vals:
        bases.append(cum if v >= 0 else cum + v)
        cum += v
    cats.append(spec.get("total_label", "합계"))
    bases.append(0)
    vals.append(cum)
    colors = _emph_colors(cats[:-1], spec.get("emphasis"), pal) + [pal["accent"]]
    ax.bar(cats, [abs(v) for v in vals], bottom=bases, color=colors, width=0.62)
    for i, (b, v) in enumerate(zip(bases, vals)):
        ax.text(
            i,
            b + abs(v) + cum * 0.02,
            f"{v:+g}" if i < len(vals) - 1 else f"{v:g}",
            ha="center",
            fontsize=8.5,
            color=pal["ink"],
        )
        if i:  # 단차 연결 점선
            prev_top = bases[i - 1] + abs(vals[i - 1]) if vals[i - 1] >= 0 else bases[i - 1]
            ax.plot(
                [i - 1 + 0.31, i - 0.31],
                [prev_top, prev_top],
                ls=":",
                lw=0.9,
                color=pal["gray_dark"],
            )
    ax.set_yticks([])


def _heatmap(ax, spec, pal):
    """행×열 밀도 매트릭스 — 흰색→accent 단색 그라데이션 + 셀 값 표기."""
    rows, cols, vals = spec["rows"], spec["cols"], spec["values"]
    cmap = LinearSegmentedColormap.from_list("brand", ["#FFFFFF", pal["accent"]])
    vmax = max(max(r) for r in vals) or 1
    ax.imshow(vals, cmap=cmap, aspect="auto", vmin=0, vmax=vmax)
    ax.set_xticks(range(len(cols)), cols)
    ax.set_yticks(range(len(rows)), rows)
    for i, row in enumerate(vals):
        for j, v in enumerate(row):
            ax.text(
                j,
                i,
                f"{v:g}",
                ha="center",
                va="center",
                fontsize=8.5,
                color="white" if v > vmax * 0.55 else pal["ink"],
            )
    for side in ("top", "right", "bottom", "left"):
        ax.spines[side].set_visible(False)


def _dumbbell(ax, spec, pal):
    """전후 비교(카테고리별 두 시점) — 회색 점(전) → accent 점(후) 연결선."""
    cats = spec["categories"]
    (n1, v1), (n2, v2) = list(spec["series"].items())[:2]
    ys = range(len(cats))
    for y, a, b in zip(ys, v1, v2):
        ax.plot([a, b], [y, y], color=pal["gray"], lw=2.2, zorder=1)
    ax.scatter(v1, list(ys), s=64, color=pal["gray_dark"], zorder=2, label=n1)
    ax.scatter(v2, list(ys), s=64, color=pal["accent"], zorder=2, label=n2)
    for y, a, b in zip(ys, v1, v2):
        ax.text(b, y - 0.28, f"{b:g}", ha="center", fontsize=8, color=pal["accent"])
    ax.set_yticks(list(ys), cats)
    ax.invert_yaxis()
    ax.legend(loc="lower right", frameon=False, fontsize=8)


def _slope(ax, spec, pal):
    """두 시점 순위·수준 변화 — 강조 시리즈만 accent 굵은 선."""
    left, right = spec["columns"][:2]
    emph = set(spec.get("emphasis") or [])
    for name, (a, b) in spec["series"].items():
        is_e = name in emph
        color = pal["accent"] if is_e else pal["gray"]
        ax.plot(
            [0, 1],
            [a, b],
            color=color,
            lw=2.4 if is_e else 1.4,
            marker="o",
            markersize=5 if is_e else 3.5,
        )
        ax.text(
            -0.03,
            a,
            f"{name}  {a:g}",
            ha="right",
            va="center",
            fontsize=8.5,
            color=pal["ink"] if is_e else pal["gray_dark"],
            fontweight="bold" if is_e else "normal",
        )
        ax.text(
            1.03,
            b,
            f"{b:g}",
            ha="left",
            va="center",
            fontsize=8.5,
            color=pal["accent"] if is_e else pal["gray_dark"],
        )
    ax.set_xticks([0, 1], [left, right])
    ax.set_xlim(-0.45, 1.25)
    ax.set_yticks([])
    for side in ("left", "bottom"):
        ax.spines[side].set_visible(False)


def _funnel(ax, spec, pal):
    """단계 전환·이탈 — 중앙 정렬 가로 막대 + 단계 간 전환율 표기."""
    stages, vals = spec["stages"], spec["values"]
    colors = _emph_colors(stages, spec.get("emphasis"), pal)
    top = vals[0] or 1
    for i, (st, v) in enumerate(zip(stages, vals)):
        ax.barh(i, v, left=(top - v) / 2, color=colors[i], height=0.62)
        # 좁은 막대(상단의 35% 미만)는 라벨이 막대 밖으로 나가므로 잉크색 고정
        wide = v >= top * 0.35
        ax.text(
            top / 2,
            i,
            f"{st}  {v:g}",
            ha="center",
            va="center",
            fontsize=9,
            color="white" if (wide and colors[i] != pal["gray"]) else pal["ink"],
        )
        if i:
            rate = vals[i] / vals[i - 1] * 100 if vals[i - 1] else 0
            ax.text(top * 1.01, i - 0.5, f"↓ {rate:.0f}%", fontsize=8, color=pal["accent"])
    ax.invert_yaxis()
    ax.set_xticks([])
    ax.set_yticks([])
    for side in ("left", "bottom"):
        ax.spines[side].set_visible(False)


def _annotated_scatter(ax, spec, pal):
    """상관+예외 주장 — 회색 점 구름, emphasis 점만 accent + 라벨."""
    pts = spec["points"]
    xs = [p["x"] for p in pts]
    ys = [p["y"] for p in pts]
    colors = [pal["accent"] if p.get("emphasis") else pal["gray"] for p in pts]
    sizes = [70 if p.get("emphasis") else 36 for p in pts]
    ax.scatter(xs, ys, c=colors, s=sizes, zorder=2)
    for p in pts:
        if p.get("emphasis") or p.get("always_label"):
            ax.annotate(
                p.get("label", ""),
                (p["x"], p["y"]),
                xytext=(6, 6),
                textcoords="offset points",
                fontsize=8.5,
                color=pal["ink"],
                fontweight="bold",
            )
    ax.set_xlabel(spec.get("x_label", ""), fontsize=8.5)
    ax.set_ylabel(spec.get("y_label", ""), fontsize=8.5)
    ax.grid(True, lw=0.4, color=pal["bg_alt"])


def _bubble(ax, spec, pal):
    """원 면적 비교(Deloitte p7 실측) — 값 내림차순 가로 배열, 면적 ∝ 값.
    spec: {"items": [{"label", "value", "emphasis"?}]} 4~9개."""
    items = sorted(spec["items"], key=lambda i: -i["value"])
    vmax = max(i["value"] for i in items) or 1
    x = 0.0
    for it in items:
        r = 0.5 * (it["value"] / vmax) ** 0.5  # 면적 비례
        cx = x + r
        color = pal["accent"] if it.get("emphasis") else pal["gray"]
        ax.add_patch(Circle((cx, 0), r, color=color))
        ax.text(
            cx,
            0,
            f"{it['value']:g}",
            ha="center",
            va="center",
            fontsize=9,
            fontweight="bold",
            color="white" if (it.get("emphasis") or r > 0.3) else pal["ink"],
        )
        ax.text(
            cx,
            -0.62,
            it["label"],
            ha="center",
            va="top",
            fontsize=8,
            color=pal["ink"] if it.get("emphasis") else pal["gray_dark"],
        )
        x += 2 * r + 0.12
    ax.set_xlim(-0.1, x)
    ax.set_ylim(-1.0, 0.62)
    ax.set_aspect("equal")
    ax.axis("off")


def _waffle(ax, spec, pal):
    """와플(픽토그램) 비율(Deloitte p24 실측) — 10×10 도트 그리드, 채움 = %.
    spec: {"items": [{"label", "percent", "emphasis"?}]} 1~3개(나란히)."""
    items = spec["items"][:3]
    gap_x = 12.5
    for gi, it in enumerate(items):
        filled = round(it["percent"])
        color = pal["accent"] if it.get("emphasis", True) else pal["gray_dark"]
        ox = gi * gap_x
        for k in range(100):
            row, col = divmod(k, 10)
            on = k < filled
            ax.add_patch(Circle((ox + col, 9 - row), 0.38, color=color if on else pal["bg_alt"]))
        ax.text(
            ox + 4.5,
            -1.5,
            f"{it['percent']:g}%",
            ha="center",
            fontsize=13,
            fontweight="bold",
            color=color,
        )
        ax.text(ox + 4.5, -3.0, it["label"], ha="center", fontsize=8, color=pal["ink"])
    ax.set_xlim(-1, (len(items) - 1) * gap_x + 10.5)
    ax.set_ylim(-4, 10.5)
    ax.set_aspect("equal")
    ax.axis("off")


def _histogram(ax, spec, pal):
    """분포 주장 — 회색 히스토그램 + accent 평균선."""
    vals = spec["values"]
    ax.hist(vals, bins=spec.get("bins", 10), color=pal["gray"], edgecolor="white")
    mean = sum(vals) / len(vals)
    ax.axvline(mean, color=pal["accent"], lw=1.8)
    ax.text(
        mean,
        ax.get_ylim()[1] * 0.96,
        f" 평균 {mean:g}",
        fontsize=8.5,
        color=pal["accent"],
        va="top",
    )
    ax.set_xlabel(spec.get("x_label", ""), fontsize=8.5)
