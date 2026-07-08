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
    # 멀티패널 chart는 패널 수만큼 네이티브 차트를 만든다
    want_charts = sum(len(s.get("panels") or [None]) for s in slides if s["type"] == "chart")

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
    checks.append(("native_tables", got_tables == want_tables, f"{got_tables}/{want_tables}"))
    checks.append(("native_charts", got_charts == want_charts, f"{got_charts}/{want_charts}"))
    # 골격: 첫 두 장 cover/toc, 마지막 cta
    types = [s["type"] for s in slides]
    skeleton_ok = types[:2] == ["cover", "toc"] and types[-1] == "cta"
    checks.append(("skeleton(cover→toc..cta)", skeleton_ok, str(types[:2] + ["..", types[-1]])))

    # 정보 밀도: 본문 슬라이드(표지·목차·섹션·CTA 제외) 단어수 하한 — 빈 슬라이드 차단
    MIN_BODY_WORDS = 60
    frame_types = {"cover", "toc", "part", "section", "cta"}
    thin = []
    for idx, (sl, spec_sl) in enumerate(zip(prs.slides, slides), 1):
        if spec_sl["type"] in frame_types:
            continue
        words = sum(len(sh.text_frame.text.split()) for sh in sl.shapes if sh.has_text_frame)
        if words < MIN_BODY_WORDS:
            thin.append(f"s{idx}({spec_sl['type']}:{words})")
    n_body = sum(1 for s in slides if s["type"] not in frame_types)
    checks.append(
        (
            f"density(본문 {MIN_BODY_WORDS}단어 미만 0장)",
            not thin,
            f"미달 {len(thin)}/{n_body}" + (f" {thin}" if thin else ""),
        )
    )

    # 구성 규칙(deck-compose 계약을 기계로 강제 — 프롬프트 규칙은 요동, 게이트는 불변)
    n_bullets = types.count("bullets")
    checks.append(("bullets_slides(≤2)", n_bullets <= 2, f"{n_bullets}/2"))
    n_parts = types.count("part")
    part_ok = n_parts >= 3 or len(slides) < 12
    checks.append(
        ("part_navigation(≥3 when slides≥12)", part_ok, f"part={n_parts}, slides={len(slides)}")
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
