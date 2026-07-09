#!/usr/bin/env python
"""실물 목업 갤러리 — 후보 전수를 실제 build_pptx로 빌드하고 PowerPoint로 렌더한다.

mpl 스케치(make_mockups.py)와 달리 최종 산출물과 동일한 룩(마스터 틀·폰트·색·다이어그램
도형)을 보여준다. 승인 게이트(②.5)의 기본 경로.

usage: python render_real_mockups.py <candidates.json> <out_dir>
요구: Windows + PowerPoint 설치(COM). 실패 시 make_mockups.py(스케치)로 폴백하라.
"""

import html
import json
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "pptx-build" / "scripts"))
from make_mockups import CSS, edit_label  # noqa: E402

import build_pptx  # noqa: E402


def candidate_slide(sl, tag, cand):
    """후보 1안 → 빌더용 슬라이드 spec (실데이터 그대로, 제목에 선택 코드)."""
    out = dict(cand)
    if "type" not in out:
        out["type"] = "chart" if out.get("chart_type") else "diagram"
    out["title"] = f"[{tag}] {sl.get('title', '')}"
    out.setdefault("subtitle", sl.get("message", ""))
    return out


def main():
    cand_path, out_dir = Path(sys.argv[1]), Path(sys.argv[2])
    out_dir.mkdir(parents=True, exist_ok=True)
    data = json.loads(cand_path.read_text(encoding="utf-8"))
    brand_path = Path(__file__).resolve().parents[2] / "pptx-build" / "assets" / "brand-kit.yaml"
    brand = build_pptx.load_brand(brand_path)

    order = []  # [(no, tag, cand)]
    slides = []
    for sl in data["slides"]:
        for i, cand in enumerate(sl.get("visual_candidates", [])):
            tag = chr(ord("A") + i)
            order.append((sl["no"], tag, sl, cand))
            slides.append(candidate_slide(sl, f"s{sl['no']:02d}-{tag}", cand))
    spec = {"meta": {"project": "exhibit-mockups", "year": "2026"}, "slides": slides}

    deck = build_pptx.Deck(brand)
    deck.render_dir = out_dir / "render"
    prs = deck.build(spec)
    pptx_path = out_dir / "candidates.pptx"
    prs.save(str(pptx_path))
    print(f"built: {pptx_path} ({len(prs.slides)} slides)")

    # PowerPoint COM으로 전장 PNG 렌더
    png_dir = out_dir / "png"
    png_dir.mkdir(exist_ok=True)
    ps = f"""
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open((Resolve-Path '{pptx_path}').Path, $true, $false, $false)
$pres.SaveAs((Resolve-Path '{png_dir}').Path, 18)
$pres.Close(); $ppt.Quit()
"""
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
    if r.returncode != 0:
        print("COM render FAIL:", r.stderr[-500:])
        sys.exit(1)
    pngs = sorted(
        png_dir.glob("*.PNG"),
        key=lambda p: int(re.sub(r"\D", "", p.stem) or 0),
    ) or sorted(png_dir.glob("*.png"), key=lambda p: int(re.sub(r"\D", "", p.stem) or 0))
    if len(pngs) != len(order):
        print(f"PNG {len(pngs)} != 후보 {len(order)} — FAIL")
        sys.exit(1)
    for (no, tag, _, _), png in zip(order, pngs):
        (out_dir / f"s{no:02d}_{tag.lower()}.png").write_bytes(png.read_bytes())

    # 갤러리 HTML (실물 렌더 카드 + 인터랙티브 세트 심사)
    def vkey(c):
        if c.get("panels"):
            return "chart:panels"
        if c.get("chart_type"):
            return f"chart:{c['chart_type']}"
        if c.get("layout") or c.get("type") == "diagram":
            return f"diagram:{c.get('layout', 'flow')}"
        return c.get("type", "?")

    BOX = {"diagram:flow", "diagram:layers", "diagram:cards", "diagram:branch", "diagram:from_to"}
    meta = [
        {"no": no, "tag": tag, "key": vkey(cand), "box": vkey(cand) in BOX}
        for no, tag, _, cand in order
    ]

    # 기본 선택: 게이트 4종을 통과하는 조합을 그리디로 사전 계산(초기 로드부터 전부 ✓)
    from collections import Counter

    by_no = {}
    for m in meta:
        by_no.setdefault(m["no"], []).append(m)
    nos = sorted(by_no)
    default_sel, picked_keys = {}, []
    kcount, boxes = Counter(), 0
    for no in nos:
        best, best_pen = by_no[no][0], 10**9
        for m in by_no[no]:
            pen = 0
            if m["key"] in picked_keys[-2:]:
                pen += 100  # 쿨다운(간격<3) 위반
            if kcount[m["key"]] >= 2:
                pen += 100  # 덱 전체 ≤2회 위반
            if m["box"] and (boxes + 1) / len(nos) > 0.30:
                pen += 50  # 박스 30% 초과
            if kcount[m["key"]]:
                pen += 3  # 새 유형 선호(≥5종 달성 유도)
            if pen < best_pen:
                best, best_pen = m, pen
        default_sel[no] = best["tag"]
        picked_keys.append(best["key"])
        kcount[best["key"]] += 1
        boxes += 1 if best["box"] else 0
    parts = [
        f"<style>{CSS}"
        ".bar{display:flex;align-items:center;gap:6px;margin:2px 0}"
        ".bar i{display:inline-block;height:12px;background:#D66E3A;border-radius:2px}"
        ".gate{font-size:13px;margin:3px 0}.ok{color:#2c6e2c}.bad{color:#B14D1F;font-weight:700}"
        "#panel{position:sticky;top:0;background:#fff;border:2px solid #15171B;padding:6px 16px;"
        "z-index:9;margin-bottom:18px}#panel[open]{padding:12px 16px}"
        "#panel summary{cursor:pointer;font-weight:700;padding:6px 0}"
        "#sel{font-family:monospace;font-size:12px}"
        "label.card{cursor:pointer}input[type=radio]:checked+span.pick{background:#D66E3A;color:#fff;"
        "padding:1px 7px;border-radius:3px}</style>",
        "<h1>익스히빗 사양서 — 실물 목업 (실제 빌더 렌더)</h1>",
        '<details id="panel"><summary>세트 심사 — 펼쳐서 히스토그램·게이트 확인 (카드 클릭 시 갱신)</summary>'
        '<div id="hist"></div><div id="gates"></div>'
        '<button onclick="copySel()">선택 문자열 복사</button> <span id="sel"></span></details>',
    ]
    cur = None
    for no, tag, sl, cand in order:
        if no != cur:
            if cur is not None:
                parts.append("</div></div>")
            cur = no
            parts.append(
                f'<div class="slide"><h2>s{no:02d}. {html.escape(sl.get("title", ""))}</h2>'
            )
            parts.append(f'<div class="msg">주장: {html.escape(sl.get("message", ""))}</div>')
            parts.append('<div class="cands">')
        label = edit_label(cand)
        cls = "image" if "수치편집불가" in label else ("native" if "네이티브" in label else "shape")
        kind = vkey(cand)
        ref = (
            f'<div class="ref">레퍼런스: {html.escape(cand["ref"])}</div>'
            if cand.get("ref")
            else ""
        )
        checked = "checked" if tag == default_sel.get(no, "A") else ""
        parts.append(
            f'<label class="card"><input type="radio" name="s{no}" value="{tag}" {checked} '
            f'onchange="update()" style="margin-right:6px"><span class="pick">{tag}</span>'
            f" — {html.escape(str(kind))}<br>"
            f'<span class="badge {cls}">{label}</span>{ref}'
            f'<img src="s{no:02d}_{tag.lower()}.png" alt="{tag}"></label>'
        )
    parts.append("</div></div>")
    parts.append(
        "<script>\nconst META = "
        + json.dumps(meta, ensure_ascii=False)
        + ";\n"
        + """
function selected(){
  const byNo = {};
  META.forEach(m => { byNo[m.no] = byNo[m.no] || {}; byNo[m.no][m.tag] = m; });
  return Object.keys(byNo).map(Number).sort((a,b)=>a-b).map(no => {
    const r = document.querySelector(`input[name="s${no}"]:checked`);
    return byNo[no][r ? r.value : 'A'];
  });
}
function update(){
  const sel = selected();
  const keys = sel.map(m => m.key);
  const counts = {};
  keys.forEach(k => counts[k] = (counts[k]||0)+1);
  const max = Math.max(...Object.values(counts));
  document.getElementById('hist').innerHTML = Object.entries(counts)
    .sort((a,b)=>b[1]-a[1])
    .map(([k,n]) => `<div class="bar"><i style="width:${n/max*160}px"></i>${k} ×${n}</div>`).join('');
  const gates = [];
  let cool = [];
  for(let i=0;i<keys.length;i++)
    for(let j=i+1;j<Math.min(i+3,keys.length);j++)
      if(keys[i]===keys[j]) cool.push(`s${sel[i].no}·s${sel[j].no}=${keys[i]}`);
  gates.push(['쿨다운(동일 유형 간격 ≥3)', cool.length===0, cool.join(', ')]);
  const kinds = new Set(keys).size;
  gates.push([`최소 5종 (현재 ${kinds}종)`, kinds>=5, '']);
  const boxes = sel.filter(m=>m.box).length;
  const ratio = Math.round(boxes/keys.length*100);
  gates.push([`박스 ≤30% (현재 ${boxes}/${keys.length}=${ratio}%)`, ratio<=30, '']);
  const over = Object.entries(counts).filter(([k,n])=>n>2).map(([k,n])=>`${k}×${n}`);
  gates.push(['동일 유형 ≤2회', over.length===0, over.join(', ')]);
  document.getElementById('gates').innerHTML = gates.map(([nm,ok,d]) =>
    `<div class="gate ${ok?'ok':'bad'}">${ok?'✓':'✗'} ${nm}${d?' — '+d:''}</div>`).join('');
  document.getElementById('sel').textContent = sel.map(m=>`s${String(m.no).padStart(2,'0')}=${m.tag}`).join(', ');
}
function copySel(){ navigator.clipboard.writeText(document.getElementById('sel').textContent); }
update();
</script>"""
    )
    out = out_dir / "gallery.html"
    out.write_text("\n".join(parts), encoding="utf-8")
    print(f"OK: {out} (후보 {len(order)}안, 인터랙티브)")


if __name__ == "__main__":
    main()
