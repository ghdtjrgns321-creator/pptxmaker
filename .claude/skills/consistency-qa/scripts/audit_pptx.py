#!/usr/bin/env python
"""빌드된 pptx를 재파싱해 네이티브 객체·골격·폰트 일관성을 기계 검증한다.

deck-spec과 대조하여 표/차트가 이미지로 새지 않았는지, 골격이 지켜졌는지 확인한다.
판단(셀링 흐름 등)은 사람/에이전트 몫 — 여기선 기계로 셀 수 있는 것만 센다.

usage: python audit_pptx.py <out.pptx> <deck-spec.json> [--brand brand-kit.yaml]
exit code 0=PASS, 1=FAIL
"""

import argparse
import json
import sys
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def collect_fonts(prs):
    fonts = set()
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.name:
                        fonts.add(run.font.name)
    return fonts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("spec")
    ap.add_argument("--brand", default="")
    args = ap.parse_args()

    prs = Presentation(args.pptx)
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    slides = spec["slides"]

    want_tables = sum(1 for s in slides if s["type"] == "table")
    want_charts = sum(1 for s in slides if s["type"] == "chart")

    got_tables = got_charts = got_pics = 0
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_table:
                got_tables += 1
            elif sh.has_chart:
                got_charts += 1
            elif sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                got_pics += 1

    checks = []
    checks.append(
        (
            "slide_count",
            len(prs.slides) == len(slides),
            f"{len(prs.slides)}/{len(slides)}",
        )
    )
    checks.append(
        ("native_tables", got_tables == want_tables, f"{got_tables}/{want_tables}")
    )
    checks.append(
        ("native_charts", got_charts == want_charts, f"{got_charts}/{want_charts}")
    )
    # 골격: 첫 두 장 cover/toc, 마지막 cta
    types = [s["type"] for s in slides]
    skeleton_ok = types[:2] == ["cover", "toc"] and types[-1] == "cta"
    checks.append(
        ("skeleton(cover→toc..cta)", skeleton_ok, str(types[:2] + ["..", types[-1]]))
    )

    # 폰트 일관성 (brand-kit 제공 시)
    if args.brand:
        brand = yaml.safe_load(Path(args.brand).read_text(encoding="utf-8"))
        allowed = {v for v in brand["fonts"].values()}
        used = collect_fonts(prs)
        stray = used - allowed
        checks.append(
            (
                "font_consistency",
                not stray,
                f"stray={sorted(stray)} allowed={sorted(allowed)}",
            )
        )

    print(f"pptx: {args.pptx}")
    all_ok = True
    for name, ok, detail in checks:
        all_ok &= ok
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}: {detail}")
    if got_pics:
        print(f"  [info] pictures={got_pics} (표/차트가 그림으로 샜는지 확인 필요)")
    print("RESULT:", "PASS" if all_ok else "FAIL")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
