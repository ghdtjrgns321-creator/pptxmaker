"""compare_golden — 골든 원본 vs goldenfab 공장 출력의 도형 전수 비교 (Phase 2 회귀 게이트).

비교 속성: 도형 수 / shape_type / left·top·width·height(±0.005") / 전체 텍스트 /
솔리드 채움색 / 첫 텍스트 런의 폰트 크기·볼드.
사용: uv run python .claude/skills/pptx-build/scripts/compare_golden.py [리포트경로]
"""

import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.util import Inches

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]  # scripts -> pptx-build -> skills -> .claude -> pptmaker
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "golden"))

EMU = 914400
TOL = 0.005


def _color_str(color_fmt):
    """ColorFormat -> 비교용 문자열. 테마색·미지정도 안전하게 처리."""
    try:
        if color_fmt.type is None:
            return None
        return str(color_fmt.rgb)
    except Exception:
        try:
            return str(color_fmt.theme_color)
        except Exception:
            return None


def _fill_str(fill_fmt):
    try:
        if fill_fmt.type == MSO_FILL_TYPE.SOLID:
            return _color_str(fill_fmt.fore_color)
    except Exception:
        pass
    return None


def _run_colors(sh):
    """모든 런의 색 — 런 단위 강조색(01 accent 등)을 잡는다. 첫 런만 보면 놓친다."""
    if not sh.has_text_frame:
        return None
    out = []
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            out.append(_color_str(r.font.color))
    return tuple(out) if out else None


def _cells(sh):
    """표 셀의 (텍스트, 채움) — 표는 GraphicFrame 1덩어리라 text_frame/fill로는 안 잡힌다."""
    if not getattr(sh, "has_table", False):
        return None
    out = []
    for row in sh.table.rows:
        for cell in row.cells:
            out.append((cell.text, _fill_str(cell.fill)))
    return tuple(out)


def shape_sig(sh):
    sig = {
        "type": str(sh.shape_type),
        "l": sh.left / EMU if sh.left is not None else None,
        "t": sh.top / EMU if sh.top is not None else None,
        "w": sh.width / EMU if sh.width is not None else None,
        "h": sh.height / EMU if sh.height is not None else None,
        "text": sh.text_frame.text if sh.has_text_frame else "",
        "fill": None,
        "run": None,
        "run_colors": _run_colors(sh),
        "cells": _cells(sh),
    }
    # str(fill.type) 비교 금지 — 실제 repr은 "<MSO_FILL_TYPE.SOLID: 1>"이라 문자열 대조는
    # 영원히 거짓이 되고 게이트가 색을 못 본 채 통과한다(2026-07-15 실측: 225개 중 0개 비교).
    try:
        sig["fill"] = _fill_str(sh.fill)
    except Exception:
        pass
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            if p.runs:
                r = p.runs[0]
                sig["run"] = (r.font.size.pt if r.font.size else None, r.font.bold)
                break
    return sig


def diff_shapes(a, b):
    """도형 서명 비교 — 불일치 필드 목록 반환(빈 리스트 = 일치)."""
    bad = []
    for k in ("type", "text", "fill", "run", "run_colors", "cells"):
        if a[k] != b[k]:
            bad.append(f"{k}: {a[k]!r} != {b[k]!r}")
    for k in ("l", "t", "w", "h"):
        av, bv = a[k], b[k]
        if (av is None) != (bv is None) or (av is not None and abs(av - bv) > TOL):
            bad.append(f"{k}: {av} != {bv}")
    return bad


def compare_slide(golden_slide, fab_slide, name):
    gs = [shape_sig(s) for s in golden_slide.shapes]
    fs = [shape_sig(s) for s in fab_slide.shapes]
    rows, mism = [], 0
    if len(gs) != len(fs):
        rows.append(f"| {name} | 도형 수 | {len(gs)} != {len(fs)} | FAIL |")
        mism += 1
    n = min(len(gs), len(fs))
    ok = 0
    for i in range(n):
        bad = diff_shapes(gs[i], fs[i])
        if bad:
            mism += 1
            rows.append(f"| {name} | 도형#{i} | {'; '.join(bad)[:120]} | FAIL |")
        else:
            ok += 1
    rows.insert(0, f"| {name} | 전체 | {ok}/{len(gs)} 일치 | {'PASS' if mism == 0 else 'FAIL'} |")
    return rows, mism, ok, len(gs)


def build_full():
    """골든 19장 조립 vs 공장(registry) 19장 조립."""
    import build_golden as BG
    from _variant_k import variant_k as g_s04
    from s06_variants import variant_c as g_s06
    from s08_variants import variant_c as g_s08
    from s09_variants import variant_a as g_s09
    from s10_screenshot import variant_a as g_s10
    from s11_variants import variant_d as g_s11
    from s12_variants import variant_b as g_s12
    from s14_variants import variant_c as g_s14
    from s15_variants import variant_c as g_s15
    from s17_variants import variant_c as g_s17
    from s18_variants import variant_b as g_s18
    from s21_closing import variant_a as g_s21

    from goldenfab.registry import LAYOUTS as FAB
    from goldenfab import layouts as FL

    g = Presentation()
    g.slide_width, g.slide_height = Inches(13.333), Inches(7.5)
    BG.build_s01_cover(g)
    BG.build_s02_toc(g)
    BG.build_part(g, 1)
    g_s04(g)
    BG.build_part(g, 2)
    g_s06(g)
    BG.build_part(g, 3)
    g_s08(g)
    g_s09(g)
    g_s10(g)
    g_s11(g)
    g_s12(g)
    BG.build_part(g, 4)
    g_s14(g)
    g_s15(g)
    BG.build_part(g, 5)
    g_s17(g)
    g_s18(g)
    g_s21(g)

    f = Presentation()
    f.slide_width, f.slide_height = Inches(13.333), Inches(7.5)
    cover_c = {
        "kicker": BG.KICKER,
        "title": BG.TITLE,
        "value_prop": BG.VALUE_PROP,
        "year": BG.YEAR,
        "brand_name": FL.K["brand"]["name"],
    }
    toc_c = {
        "title": "목차",
        "items": BG.TOC,
        "deck_title": BG.TITLE,
        "deck_sub": "기술 제안서 · 2026",
    }

    def part_c(no):
        return {"no": no, "title": BG.TOC[no - 1][0], "lead": BG.PART_LEADS[no - 1]}

    FAB["cover"](f, cover_c)
    FAB["toc"](f, toc_c)
    FAB["part"](f, part_c(1))
    FAB["problem_grid"](f)
    FAB["part"](f, part_c(2))
    FAB["exec_graph"](f)
    FAB["part"](f, part_c(3))
    FAB["tech_evidence"](f)
    FAB["tech_tree"](f)
    FAB["screenshot"](f)
    FAB["tech_mechanism"](f)
    FAB["tech_capture"](f)
    FAB["part"](f, part_c(4))
    FAB["ab_simulation"](f)
    FAB["validation"](f)
    FAB["part"](f, part_c(5))
    FAB["mirror_matrix"](f)
    FAB["boundary"](f)
    FAB["closing"](f)

    names = [
        "S1 표지",
        "S2 목차",
        "S3 간지1",
        "S4 문제정의",
        "S5 간지2",
        "S6 실행그래프",
        "S7 간지3",
        "S8 용어사전",
        "S9 지식그래프",
        "S10 스크린샷",
        "S11 판단트리",
        "S12 구조화출력",
        "S13 간지4",
        "S14 트러블슈팅",
        "S15 골든테스트",
        "S16 간지5",
        "S17 미러매트릭스",
        "S18 경계",
        "S19 클로징",
    ]
    return g, f, names


def param_efficacy():
    """파라미터 유효성 — 상이 입력이 출력 텍스트에 반영되는지 2케이스."""
    from goldenfab.registry import LAYOUTS as FAB

    results = []
    f = Presentation()
    f.slide_width, f.slide_height = Inches(13.333), Inches(7.5)
    FAB["closing"](
        f,
        {
            "value_prop": "테스트 가치제안 문장 A" + chr(10) + "둘째 줄 B",
            "kicker": "CLOSING — 테스트덱",
        },
    )
    texts = " ".join(sh.text_frame.text for sh in f.slides[0].shapes if sh.has_text_frame)
    results.append(
        ("closing 커스텀", "테스트 가치제안 문장 A" in texts and "CLOSING — 테스트덱" in texts)
    )

    f2 = Presentation()
    f2.slide_width, f2.slide_height = Inches(13.333), Inches(7.5)
    FAB["screenshot"](f2, {"headline": "커스텀 헤드라인 검사용", "caption": "커스텀 캡션 — 검사"})
    texts2 = " ".join(sh.text_frame.text for sh in f2.slides[0].shapes if sh.has_text_frame)
    results.append(
        ("screenshot 커스텀", "커스텀 헤드라인 검사용" in texts2 and "커스텀 캡션 — 검사" in texts2)
    )

    # Phase B 파라미터화 10종 — headline override가 출력에 반영되는지(내장 아님) 검증
    for t in (
        "problem_grid",
        "exec_graph",
        "tech_evidence",
        "tech_tree",
        "tech_mechanism",
        "tech_capture",
        "ab_simulation",
        "validation",
        "mirror_matrix",
        "boundary",
    ):
        ft = Presentation()
        ft.slide_width, ft.slide_height = Inches(13.333), Inches(7.5)
        sent = "SENT_" + t.upper()
        FAB[t](ft, {"headline": sent})
        tt = " ".join(sh.text_frame.text for sh in ft.slides[0].shapes if sh.has_text_frame)
        results.append((f"{t} override", sent in tt))
    return results


def main():
    out_path = (
        Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "golden" / "variants" / "compare_full.md"
    )
    g, f, names = build_full()
    lines = [
        "# compare_golden — 골든 19장 vs goldenfab 도형 전수 대조",
        "",
        f'허용오차 ±{TOL}" · 비교 속성: 수/type/text/fill/run(pt·bold)/l·t·w·h',
        "",
        "| 슬라이드 | 대상 | 내용 | 판정 |",
        "|---|---|---|---|",
    ]
    total_mism = total_ok = total_n = 0
    for gs, fs2, name in zip(g.slides, f.slides, names):
        rows, mism, ok, n = compare_slide(gs, fs2, name)
        lines.extend(rows)
        total_mism += mism
        total_ok += ok
        total_n += n
    lines.append("")
    lines.append(
        f"총계: 슬라이드 {len(names)}/19, 도형 일치 {total_ok}/{total_n}, 불일치 {total_mism}건"
    )
    lines.append("")
    lines.append("## 파라미터 유효성 (상이 입력 반영)")
    par = param_efficacy()
    for name, ok in par:
        lines.append(f"- {name}: {'PASS' if ok else 'FAIL'}")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(chr(10).join(lines) + chr(10), encoding="utf-8")
    print(chr(10).join(lines[-30:]))
    assert total_mism == 0, f"COMPARE FAIL {total_mism}건"
    assert all(ok for _, ok in par), "PARAM FAIL"
    print(f"COMPARE PASS -> {out_path}")


if __name__ == "__main__":
    main()
