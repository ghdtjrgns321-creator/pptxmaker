# 시안 K — 사용자 GRID 스냅판 (grid.py 상수만 사용, 임의 좌표 금지)
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from . import grid as G
from ._variant_h import _shape_step
from .kit import SLIDE_H, SLIDE_W, add_box, add_text, load_kit, mix

K = load_kit()
C, S, F = K["rgb"], K["sizes"], K["fonts"]

QW, QH = 1.45, 1.3
CONN_W = 0.4  # 질문→밴드 커넥터 구간
FX = G.COL_R_X + QW + CONN_W  # 셰브런 행 시작 5.95
OVERLAP = G.FLOW_H * 0.55 / 2
STEP_W = (G.RIGHT_EDGE - FX + 3 * OVERLAP) / 4
Q_CY = (G.FLOW_ROW1_Y + G.FLOW_ROW2_Y + G.FLOW_H) / 2  # 3.56 — 덤프 기준 재계산식


def _subhead(slide, y, text):
    add_text(
        slide, G.COL_R_X, y, G.COL_R_W, 0.32, text, S["head"], F["head"], C["primary"], bold=True
    )
    add_box(slide, G.COL_R_X, y + 0.35, G.COL_R_W, 0.012, fill=C["muted"])


def variant_k(prs):
    """K안(v12) — GRID 스냅: 좌 0.6 / 우 4.2 / 셰브런 right 12.733 / 하한 6.35."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=C["bg"])
    # 헤더 (룰 = RULE_Y)
    add_text(
        slide,
        G.MARGIN_L,
        0.42,
        6.0,
        0.28,
        "1. 문제 정의",
        S["caption"],
        F["head"],
        C["muted"],
        bold=True,
    )
    add_text(
        slide,
        G.MARGIN_L,
        0.72,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.55,
        "일반 LLM을 회계에 못 쓰는 이유는 틀려서가 아니라, 틀리는 방향 때문이다",
        S["section"],
        F["head"],
        C["primary"],
        bold=True,
    )
    add_box(slide, G.MARGIN_L, G.RULE_Y, G.RIGHT_EDGE - G.MARGIN_L, 0.014, fill=C["muted"])

    # ── 좌 컬럼 (left = COL_L_X 통일, bottom ≤ 6.35) ──
    lx, lw = G.COL_L_X, G.COL_L_W
    add_text(
        slide,
        lx,
        G.SUBHEAD_Y,
        lw,
        0.26,
        "핵심 리스크",
        S["caption"],
        F["head"],
        C["muted"],
        bold=True,
    )
    add_text(slide, lx, 2.08, lw, 0.7, "2종 오류", S["display"], F["head"], C["accent"], bold=True)
    add_text(
        slide, lx, 2.78, lw, 0.48, "= 허위 확정", S["title"], F["head"], C["primary"], bold=True
    )
    add_text(
        slide,
        lx,
        3.38,
        lw,
        0.8,
        "근거 없는 확신은 부실감사와 재무제표 왜곡이라는 실제 손해로 이어진다.",
        S["sub"],
        F["body"],
        C["text"],
        line_spacing=1.25,
    )
    nums = [
        (
            "01",
            "검증 불가",
            "그럴듯한 답을 즉시 내놓지만, 근거 문단을 확인할 방법이 없어 감사 조서에 쓸 수 없다.",
            4.4,
        ),
        ("02", "판단 창작", "기준서에 실재하지 않는 기준을 확신에 찬 어조로 창작한다.", 5.42),
    ]
    for no, head, desc, ny in nums:
        add_text(slide, lx, ny, 0.5, 0.3, no, S["head"], F["head"], C["muted"], bold=True)
        add_text(
            slide, lx + 0.5, ny, lw - 0.5, 0.28, head, S["body"], F["head"], C["primary"], bold=True
        )
        add_text(
            slide,
            lx + 0.5,
            ny + 0.3,
            lw - 0.5,
            0.6,
            desc,
            S["caption"],
            F["body"],
            C["muted"],
            line_spacing=1.2,
        )
    add_box(
        slide, G.V_RULE_X, G.CONTENT_TOP, 0.014, G.CONTENT_BOTTOM - G.CONTENT_TOP, fill=C["bg_alt"]
    )

    # ── 우 컬럼 구획 1 (left = COL_R_X 통일) ──
    _subhead(slide, G.SUBHEAD_Y, "같은 질문, 두 개의 결말")
    q = add_box(
        slide,
        G.COL_R_X,
        Q_CY - QH / 2,
        QW,
        QH,
        fill=C["bg_alt"],
        line=C["muted"],
        line_w=0.75,
        shape="round",
    )
    tf = q.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE  # 도형 내부 중앙 — 프레임 정렬은 GRID로 고정됨
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = '"이 계약, 수익을\n지금 인식해도\n됩니까?"'
    r.font.name, r.font.bold = F["head"], True
    r.font.size = Pt(S["caption"])
    r.font.color.rgb = C["primary"]
    for row_y in (G.FLOW_ROW1_Y, G.FLOW_ROW2_Y):
        conn = slide.shapes.add_connector(
            MSO_CONNECTOR.ELBOW,
            Inches(G.COL_R_X + QW),
            Inches(Q_CY),
            Inches(FX),
            Inches(row_y + G.FLOW_H / 2),
        )
        conn.line.color.rgb = C["muted"]
        conn.line.width = Pt(1.75)

    lanes = [
        (
            G.FLOW_ROW1_Y,
            "일반 LLM",
            C["accent"],
            [
                ("즉시 답변", "근거 조문 없음"),
                ("확신형 어조", "판단 창작"),
                ("검증 불가", "인용 불가"),
            ],
            "치명",
        ),
        (
            G.FLOW_ROW2_Y,
            "본 시스템",
            C["primary"],
            [("근거 먼저 검색", "문단·사례 인용"), ("확정", "결론+인용"), ("유보", "정보 되물음")],
            "안전",
        ),
    ]
    for row_y, name, ink, steps, verdict in lanes:
        add_text(
            slide,
            FX,
            row_y - 0.28,
            3.0,
            0.24,
            name,
            S["caption"],
            F["head"],
            C["primary"],
            bold=True,
        )
        for i, (head, desc) in enumerate(steps):
            _shape_step(
                slide,
                FX + i * (STEP_W - OVERLAP),
                row_y,
                STEP_W,
                G.FLOW_H,
                "pentagon" if i == 0 else "chevron",
                C["bg_alt"],
                head,
                desc,
                C["primary"],
                mix(C["primary"], C["bg"], 0.35),
            )
        _shape_step(
            slide,
            FX + 3 * (STEP_W - OVERLAP),
            row_y,
            STEP_W,
            G.FLOW_H,
            "chevron",
            ink,
            verdict,
            "",
            C["bg"],
            C["bg"],
        )

    # ── 우 컬럼 구획 2: 하단 2박스 (top 5.30, h 1.05, bottom 6.35) ──
    _subhead(slide, G.BOX_Y - 0.6, "오류의 두 방향 — 왜 2종만 치명적인가")  # 룰=4.70+0.35=5.05, 박스 5.30
    panels = [
        (
            G.COL_R_X,
            C["bg_alt"],
            mix(C["muted"], C["bg"], 0.3),
            "1종 · 놓침 — 안전한 실패",
            C["primary"],
            '근거가 있는데 못 찾아 "모른다"고 답하는 실패는 추가 검토로 해소된다. 실측된 유보 편향(되물음 9건)도 프롬프트 재균형으로 관리하고 있다.',
            C["text"],
        ),
        (
            G.COL_R_X + G.BOX_W + G.BOX_GAP,
            C["primary"],
            C["accent"],
            "2종 · 허위 확정 — 치명적 실패",
            C["bg"],
            "근거가 없거나 틀렸는데 확신에 찬 답은 곧 부실감사로 직결된다. 설계 전체를 이 방향의 차단에 편향시켰다.",
            C["bg_alt"],
        ),
    ]
    for px, fill, spine, head, head_ink, desc, desc_ink in panels:
        add_box(slide, px, G.BOX_Y, G.BOX_W, G.BOX_H, fill=fill, shape="round")
        add_box(slide, px, G.BOX_Y, 0.05, G.BOX_H, fill=spine)  # 박스와 top/height 완전 일치, 내부 포함
        add_text(
            slide,
            px + 0.2,
            G.BOX_Y + 0.12,
            G.BOX_W - 0.4,
            0.28,
            head,
            S["body"],
            F["head"],
            head_ink,
            bold=True,
        )
        add_text(
            slide,
            px + 0.2,
            G.BOX_Y + 0.44,
            G.BOX_W - 0.4,
            0.55,
            desc,
            S["caption"],
            F["body"],
            desc_ink,
            line_spacing=1.2,
        )

    # ── 바 + 출처 ──
    bar = add_box(slide, G.MARGIN_L, G.BAR_Y, G.RIGHT_EDGE - G.MARGIN_L, G.BAR_H, fill=C["primary"])
    tfb = bar.text_frame
    tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
    pb = tfb.paragraphs[0]
    pb.alignment = PP_ALIGN.CENTER
    rb = pb.add_run()
    rb.text = "확실할 때만 확정하고, 애매하면 근거를 보여주며 유보한다"
    rb.font.name, rb.font.bold = F["head"], True
    rb.font.size = Pt(S["body"])
    rb.font.color.rgb = C["bg"]
    add_text(
        slide,
        G.MARGIN_L,
        G.SOURCE_Y,
        G.RIGHT_EDGE - G.MARGIN_L,
        0.3,
        "출처: 1_OVERVIEW.md · PROJECT_OVERVIEW.md · 5_INTERFACE.md (00_factsheet.md §A·§B)",
        S["foot"],
        F["body"],
        C["muted"],
    )


def audit(pptx_path):
    """오딧 4종 — 위반 리스트 반환."""
    from pptx import Presentation

    EMU = 914400
    prs = Presentation(pptx_path)
    s = list(prs.slides)[-1]
    tol = 0.02
    viol = {"좌 left≠0.6": [], "우 left≠4.2": [], "bottom>6.35": [], "셰브런 right≠12.733": [],
            "룰-도형 겹침": []}
    rules, solids = [], []
    for i, sh in enumerate(s.shapes):
        L, T = sh.left / EMU, sh.top / EMU
        W, H = sh.width / EMU, sh.height / EMU
        if W > 12 or (L == 0 and T == 0):
            continue
        if H <= 0.02 and W > 1:          # 수평 룰
            rules.append((i, L, T, W))
        elif H > 0.3 and W > 0.5 and sh.shape_type != 17:  # 채움 도형만(TEXT_BOX=17 제외)
            solids.append((i, L, T, W, H))
    for ri, rl, rt, rw_ in rules:
        for si, sl, st, sw, shh in solids:
            if rl < sl + sw and rl + rw_ > sl and st - 0.005 <= rt <= st + shh + 0.005:
                viol["룰-도형 겹침"].append((ri, si, round(rt, 3), round(st, 3)))
    for i, sh in enumerate(s.shapes):
        L, T = sh.left / EMU, sh.top / EMU
        W, H = sh.width / EMU, sh.height / EMU
        text = (
            sh.text_frame.text.strip().replace("\n", "|")[:18]
            if sh.has_text_frame and sh.text_frame.text.strip()
            else f"shape{i}"
        )
        if W > 12 or (L == 0 and T == 0):  # 배경·헤더 풀폭 제외
            continue
        if L < 3.7 - tol and abs(L - G.COL_L_X) > tol and L > 0.5:
            pass  # 01 번호(0.6)와 헤드(1.1)는 의도적 들여쓰기 — 블록 좌변은 0.6
        if abs(L - 1.1) < tol or abs(L - (G.COL_L_X + 0.5)) < tol:
            pass
        elif L < 3.7 and abs(L - G.COL_L_X) > tol:
            viol["좌 left≠0.6"].append((i, text, round(L, 3)))
        if W < 0.1 or H < 0.05:  # 세로 룰·스파인 바·헤어라인은 정렬 대상 아님
            continue
        if 3.9 < L < 4.5 and abs(L - G.COL_R_X) > tol and abs(L - 4.4) > tol:  # 4.4 = 박스 텍스트 시작선(사용자 지정)
            viol["우 left≠4.2"].append((i, text, round(L, 3)))
        if T + H > G.CONTENT_BOTTOM + tol and T < G.BAR_Y - tol:
            viol["bottom>6.35"].append((i, text, round(T + H, 3)))
        if text in ("치명", "안전"):
            if abs((L + W) - G.RIGHT_EDGE) > tol:
                viol["셰브런 right≠12.733"].append((i, text, round(L + W, 3)))
    return viol
