"""S9 시안 — 기술 2: 지식그래프 (§8d 기술 장 문법: 서사 3칼럼 + 실물 증거. UI 패스).

재료: 00_factsheet.md §C·§D (원문 2_DATA-TAXONOMY·3_KNOWLEDGE-GRAPH).
실물 증거: knowledge_graph_3d.png(V1 스냅샷 — 수치 표기는 v14만 사용) + 노드·간선 등록부 표.
실행: uv run python golden/s09_variants.py → golden/variants/s09_variants.pptx
"""

from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from . import grid as G
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit, set_shape_text

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]

IMG_3D = r"C:\Users\ghdtj\workspace\portfolio\k-ifrs-1115\images\knowledge_graph_3d.png"

# 콘텐츠 기본값(골든 내용) — c=None이면 이 값, override 시 텍스트만 교체(좌표·색 고정)
DEFAULT = {
    "kicker": "3. 기술 설명 — TECH 02 · Retrieve",
    "headline": "지식그래프 — 기준서의 구조를 그대로 옮긴 결정적 지도",
    "narratives": [
        (
            "왜 필요한가",
            "문단 뭉치에는 순서도 관계도 없다. 기준서는 계층·상호참조·사례가 얽힌 구조 — 그 구조를 보존해야 근거가 근거로 이어진다.",
        ),
        (
            "파이프라인에서의 역할",
            "Retrieve의 지도. 진입 개념에서 관계를 한 단계만 따라 문단·사례·근거를 수집한다 — 수집 범위 자체가 그래프로 고정된다.",
        ),
        (
            "어떻게 만들었나",
            "기준서 공식 소제목 80개를 그대로 개념 노드로, 뼈대 간선은 기준서에서 기계 생성 — AI 판단 0. 고립 노드 0을 감사로 확인했다.",
        ),
    ],
    "struct_head": "구조 — 기준서의 위계 그대로",
    "root": "기준서 1115",
    "concepts": ["변동대가", "보증", "⋯ 80"],
    "paras": ["문단 50", "문단 56", "B33"],
    "layer_tag1": "개념 80",
    "layer_tag2": "문단 250",
    "term_chip": "“볼륨디스카운트”",
    "term_cap": "용어 423 — 진입",
    "case_chip": "사례 188",
    "xref_cap": "사례는 인용 문단으로 연결 · 문단끼리는 상호참조(E3 244)",
    "table_title": ("무엇을 어떻게 만들었나", "    그래프 v14 — 노드 929 · 간선 2,697 · 고립 0"),
    "build_rows": [  # 유형별 구축 방법 — 2_DATA-TAXONOMY·3_KNOWLEDGE-GRAPH 사실만
        ("구성 요소", "어떻게 만들었나", "수"),
        ("개념", "기준서 공식 소제목을 그대로 노드로 — 이름도 경계도 창작하지 않음", "80"),
        (
            "문단·계층",
            "본문의 위계(목차 구조)를 그대로 간선으로 — 내 판단이 아니라 기준서의 질서",
            "250 · 79",
        ),
        (
            "사례",
            "QNA·감리는 제목이 정제돼 있어 제목을 색인으로, 본문이 인용한 문단 번호로 그래프에 연결",
            "188 (연결 1,220)",
        ),
        (
            "상호참조",
            "문단이 서로를 인용하는 관계를 간선으로 — 임베딩이 못 잡는 '법적 이웃'",
            "244",
        ),
        ("용어", "사람 1차 자료 3종에서 전수 검수로 등재 (TECH 01)", "423"),
    ],
    "bar": "임베딩이 놓치는 '법적 이웃'을 관계가 잡는다 — 텍스트가 아니라 구조로 검색한다",
    "source": "출처: 2_DATA-TAXONOMY.md · 3_KNOWLEDGE-GRAPH.md (00_factsheet.md §C·§D)",
}


def header(slide, headline, kicker):
    add_text(
        slide, G.MARGIN_L, 0.42, 8.0, 0.28, kicker, S["caption"], F["head"], C["muted"], bold=True
    )
    add_text(
        slide,
        G.MARGIN_L,
        0.72,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.55,
        headline,
        S["section"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_box(slide, G.MARGIN_L, G.RULE_Y, G.RIGHT_EDGE - G.MARGIN_L, 0.014, fill=C["muted"])


def narrative_row(slide, narratives):
    nar_w = (G.RIGHT_EDGE - G.MARGIN_L - 2 * 0.4) / 3
    for i, (head, body) in enumerate(narratives):
        nx = G.MARGIN_L + i * (nar_w + 0.4)
        add_text(
            slide,
            nx,
            G.CONTENT_TOP,
            nar_w,
            0.28,
            [(f"0{i + 1}", {"color": C["accent"]}), (f"  {head}", {})],
            S["head"],
            F["head"],
            C["primary"],
            bold=True,
        )
        add_text(
            slide,
            nx,
            G.CONTENT_TOP + 0.36,
            nar_w,
            0.95,
            body,
            S["caption"],
            F["body"],
            C["muted"],
            line_spacing=1.25,
        )
        if i < 2:
            add_box(slide, nx + nar_w + 0.2, G.CONTENT_TOP, 0.012, 1.2, fill=C["bg_alt"])


def build_table(slide, rows, x, y, w):
    """유형별 구축 방법 3열 표 — 방법 서술이 본문(넓은 가운데 열), 수치는 우측 보조."""
    n = len(rows)
    c1, c3 = 1.15, 1.35
    gf = slide.shapes.add_table(n, 3, Inches(x), Inches(y), Inches(w), Inches(0.46 * n))
    tbl = gf.table
    for ci, cw in enumerate((c1, w - c1 - c3, c3)):
        tbl.columns[ci].width = Inches(cw)
    for ri, row in enumerate(rows):
        tbl.rows[ri].height = Inches(0.34 if ri == 0 else 0.44)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = (
                C["accent"] if ri == 0 else (C["bg_alt"] if ri % 2 == 0 else C["bg"])
            )
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p._p.get_or_add_pPr().set("eaLnBrk", "0")
            if ci == 2:
                p.alignment = PP_ALIGN.RIGHT
            r = p.add_run()
            r.text = val
            r.font.name = F["head"] if ri == 0 or ci == 0 else F["body"]
            r.font.size = Pt(S["caption"] if ci != 1 else S["foot"] + 1)
            r.font.bold = ri == 0 or ci == 0
            r.font.color.rgb = C["bg"] if ri == 0 else (C["primary"] if ci != 1 else C["text"])


def bar_and_source(slide, text, source):
    bar = add_box(slide, G.MARGIN_L, G.BAR_Y, G.RIGHT_EDGE - G.MARGIN_L, G.BAR_H, fill=C["primary"])
    tf = bar.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name, run.font.bold = F["head"], True
    run.font.size = Pt(S["body"])
    run.font.color.rgb = C["bg"]
    add_text(
        slide,
        G.MARGIN_L,
        G.SOURCE_Y,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.3,
        source,
        S["foot"],
        F["body"],
        C["muted"],
    )


def variant_a(prs, c=None):
    """A — 서사 3칼럼 + 실물 증거(3D 캡처 + 노드·간선 등록부 표). c: 텍스트 override(None=골든 기본)."""
    c = {**DEFAULT, **(c or {})}
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    header(slide, c["headline"], c["kicker"])
    narrative_row(slide, c["narratives"])

    # ── 하단 좌: 위계 도해 (기준서의 질서 그대로) + 3D 캡처 썸네일 ──
    def _arrow(x1, y1, x2, y2):
        from pptx.oxml.ns import qn

        conn = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
        )
        conn.line.color.rgb = C["muted"]
        conn.line.width = Pt(1.5)
        ln = conn.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))

    add_text(
        slide,
        G.MARGIN_L,
        3.3,
        4.0,
        0.24,
        c["struct_head"],
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
    )

    # 4층 트리 — 루트 → 개념(fan) → 문단(fan), 용어·사례 칩 부착. 예시 라벨은 실데이터.
    def _node(x, y, w, name, *, fill=None, ink=None, bold=False, line=None):
        b = add_box(
            slide,
            x,
            y,
            w,
            0.34,
            fill=fill or C["bg_alt"],
            line=line,
            line_w=0.75 if line else None,
            shape="round",
        )
        set_shape_text(b, name, S["caption"], F["head"], ink or C["primary"], bold=bold)
        return b

    root_x, root_w = 2.35, 1.6
    _node(root_x, 3.6, root_w, c["root"], fill=C["primary"], ink=C["bg"], bold=True)
    # 개념층 (y 4.45): 변동대가 · 보증 · ⋯ (80) — 좌표 고정, 텍스트만 c에서
    concept_pos = [(2.05, 1.0), (3.2, 0.7), (4.05, 0.65)]  # (cx, cw)
    for name, (cx, cw) in zip(c["concepts"], concept_pos):
        _node(cx, 4.45, cw, name)
        _arrow(root_x + root_w / 2, 3.94, cx + cw / 2, 4.45)
    # 문단층 (y 5.35): 변동대가→50·56, 보증→B33 (상호참조 화살표 공간 확보용 간격)
    para_pos = [(1.75, 0.8, 2.55), (2.95, 0.8, 2.55), (3.95, 0.6, 3.55)]  # (px, pw, src_cx)
    for name, (px, pw, src_cx) in zip(c["paras"], para_pos):
        _node(px, 5.35, pw, name)
        _arrow(src_cx, 4.79, px + pw / 2, 5.35)
    # 층 태그 (좌측 열)
    add_text(
        slide,
        G.MARGIN_L,
        4.5,
        1.15,
        0.22,
        c["layer_tag1"],
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_text(
        slide,
        G.MARGIN_L,
        5.4,
        1.15,
        0.22,
        c["layer_tag2"],
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
    )
    # 용어 칩 → 변동대가 (aliases 실물: 볼륨디스카운트 → 변동대가)
    _node(G.MARGIN_L, 3.6, 1.45, c["term_chip"], fill=C["bg"], line=C["muted"])
    add_text(slide, G.MARGIN_L, 3.96, 1.5, 0.18, c["term_cap"], S["foot"], F["body"], C["muted"])
    _arrow(G.MARGIN_L + 1.45, 3.77, 2.05, 4.62)  # 용어는 개념 좌변으로 진입(루트 fan과 분리)
    # 사례 칩 → 문단층 (인용 문단으로 연결)
    _node(G.MARGIN_L, 5.85, 1.1, c["case_chip"], fill=C["bg"], line=C["muted"])
    _arrow(G.MARGIN_L + 1.1, 6.02, 1.75, 5.69)
    # 상호참조 양방향 (문단 50 ↔ 문단 56)
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(2.55), Inches(5.52), Inches(2.95), Inches(5.52)
    )
    conn.line.color.rgb = C["muted"]
    conn.line.width = Pt(1.25)
    _ln = conn.line._get_or_add_ln()
    from pptx.oxml.ns import qn as _qn

    _ln.append(_ln.makeelement(_qn("a:headEnd"), {"type": "triangle", "w": "sm", "len": "sm"}))
    _ln.append(_ln.makeelement(_qn("a:tailEnd"), {"type": "triangle", "w": "sm", "len": "sm"}))
    add_text(
        slide,
        1.85,
        6.14,
        3.0,
        0.2,
        c["xref_cap"],
        S["foot"],
        F["body"],
        C["muted"],
    )
    # ── 하단 우: 유형별 구축 방법 표 (방법이 본문, 수치는 보조) ──
    tx = 5.0
    tw = G.RIGHT_EDGE - tx
    add_text(
        slide,
        tx,
        3.3,
        tw,
        0.24,
        [
            [
                (
                    c["table_title"][0],
                    {"bold": True, "color": C["primary"], "font": F["head"]},
                ),
                (c["table_title"][1], {}),
            ]
        ],
        S["caption"],
        F["body"],
        C["muted"],
    )
    build_table(slide, c["build_rows"], tx, 3.62, tw)
    bar_and_source(slide, c["bar"], c["source"])
    return slide


def audit(prs):
    EMU = 914400
    fails = []
    for si, sl in enumerate(prs.slides):
        for sh in sl.shapes:
            L, T, W, H = sh.left / EMU, sh.top / EMU, sh.width / EMU, sh.height / EMU
            if L == 0 and T == 0 and W > 13:
                continue
            if T + H > G.CONTENT_BOTTOM + 0.02 and T < G.BAR_Y - 0.02:
                fails.append((si, "bottom>6.35", round(T + H, 3)))
            if 11.5 < W < 13 and abs((L + W) - G.RIGHT_EDGE) > 0.02:
                fails.append((si, "풀폭 right≠12.733", round(L + W, 3)))
    assert not fails, f"AUDIT FAIL {fails}"
    print("audit pass")


def main():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    variant_a(prs)
    audit(prs)
    out = Path(__file__).parent / "variants" / "s09_variants.pptx"
    prs.save(str(out))
    print(f"saved {out}")


if __name__ == "__main__":
    main()
