#!/usr/bin/env python
"""Lucide(MIT) SVG 아이콘을 brand-kit 색상 PNG로 변환해 assets/icons/에 박제한다.

브랜드 색이 바뀌면 이 스크립트만 다시 돌리면 전체 아이콘이 재생성된다(ripple 단일 경로).
SVG 렌더는 pymupdf를 일회성으로 사용: `uv run --with pymupdf python make_icons.py`

usage: uv run --with pymupdf python make_icons.py [--brand <brand-kit.yaml>]
"""

import argparse
import re
import urllib.request
from pathlib import Path

import fitz  # pymupdf
import yaml

# B2B 셀링 덱 범용 아이콘 큐레이션 — 추가는 이 목록에만
ICONS = [
    "search",
    "database",
    "shield-check",
    "triangle-alert",
    "circle-check-big",
    "circle-x",
    "git-branch",
    "workflow",
    "layers",
    "gauge",
    "trending-up",
    "chart-column",
    "file-text",
    "book-open",
    "scale",
    "lock",
    "settings",
    "users",
    "target",
    "rocket",
    "calendar",
    "coins",
    "link",
    "cpu",
]
CDN = "https://unpkg.com/lucide-static@0.462.0/icons/{name}.svg"
PNG_PX = 256  # 삽입 크기(0.3~0.5in) 대비 고해상도


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--brand",
        default=str(Path(__file__).parents[2] / "pptx-build" / "assets" / "brand-kit.yaml"),
    )
    args = ap.parse_args()
    brand = yaml.safe_load(Path(args.brand).read_text(encoding="utf-8"))
    colors = {
        "primary": brand["colors"]["primary"],
        "accent": brand["colors"]["accent"],
        "white": "FFFFFF",
    }
    out_dir = Path(args.brand).parent / "icons"
    out_dir.mkdir(exist_ok=True)
    svg_cache = out_dir / "_svg"
    svg_cache.mkdir(exist_ok=True)

    made = 0
    failed = []
    for name in ICONS:
        svg_path = svg_cache / f"{name}.svg"
        if not svg_path.exists():
            try:
                urllib.request.urlretrieve(CDN.format(name=name), svg_path)
            except Exception as e:
                failed.append(f"{name}({e})")
                continue
        svg = svg_path.read_text(encoding="utf-8")
        for cname, hex6 in colors.items():
            colored = re.sub(r'stroke="currentColor"', f'stroke="#{hex6}"', svg)
            doc = fitz.open(stream=colored.encode(), filetype="svg")
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(PNG_PX / 24, PNG_PX / 24), alpha=True)
            pix.save(out_dir / f"{name}_{cname}.png")
            made += 1
    print(f"OK: {made}개 PNG → {out_dir}")
    if failed:
        print(f"FAILED {len(failed)}종(이름 확인 필요): {failed}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
