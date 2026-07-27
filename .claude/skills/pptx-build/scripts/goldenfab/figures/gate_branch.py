"""gate_branch — 판정이 결과를 가르는 인과도 (골든 s04 상단 구획에서 꺼냄, 2026-07-26).

출처: `_variant_k`(골든 S4)의 게이트 인과도를 부품 계약 3조건으로 옮긴 것. 좌표·도형·색은
골든 실측 그대로여서 `compare_golden` 픽셀 동일로 무손실을 증명한다.

## 이 도해가 무엇을 말하나

같은 질문이 ◇판정 하나를 지나면서 **결과가 갈린다**. 갈래마다 그 뒤에 무슨 일이 일어나는지를
3단계(행동 → 오류 → 종착)로 펼치고, 한 갈래만 종착까지 도달하고 다른 갈래는 **질문으로
되돌아온다**. 되돌아옴을 글자가 아니라 실제 선으로 그리는 것이 이 도해의 핵심이다 — 글자로만
쓰면 비교의 두 항이 서로 다른 층위(시각 vs 텍스트)에 놓여 대비가 성립하지 않는다.

## 갈래가 왜 2 고정인가

두 갈래의 차이가 자리가 아니라 **의미**다 — 하나는 종착(치명)이고 하나는 복귀(회복 가능)라
색·종착 도형·복귀 호가 전부 다르다. 항목 수를 늘리는 축이 아니므로 3개가 들어오면 조용히
자르지 않고 시끄럽게 죽는다. "갈래가 몇이냐 자체가 메시지"인 내용은 다른 부품(n_branch)이
맡는다.
"""

from pptx.enum.text import PP_ALIGN

from ..kit import add_box, add_text, mix
from . import elements as E
from . import Box, merged, require

META = {
    "name": "gate_branch",
    "어휘": "G01",  # archetype-catalog 등재 어휘(gate_causality)와 대응
    "물성": "같은 입력이 판정 하나에서 갈리고, 갈래마다 결말이 다르다(하나는 종착·하나는 복귀)",
    "arity": "갈래 2 고정(의미 고정) · 갈래 뒤 3단계는 자리 파생",
    # n 범위는 **실측**이다(주력 밴드 2.16" 실측) — 지어낸 값이 아니다. 다시 재려면
    # `verify_selection.py --measure`. 실제 자리에서의 정확한 한계는 draw가 다시 본다.
    # 셀 수 있는 선택 조건 — 컴포저는 재료를 세어 이것과 **비교만** 한다(산문 해석 없음).
    "accepts": {'sets': 1, 'flow': '1:N', 'order': False, 'extra': True, 'n': (2, 2), 'ends': False},
    "keys": ["question", "gate", "branch_ok", "branch_lack", "ok_chip", "lack_chip", "rows"],
    "optional": ["ok_note", "back_label"],
}

LAYOUT = {
    # box 원점 기준 offset(inch). 골든 s04 실측값 — 바꾸면 compare_golden이 잡는다.
    "q_w": 1.55,
    "q_h": 1.05,
    "dia_dx": 1.85,
    "dia_w": 1.35,
    "dia_h": 0.95,
    "row_h": 0.48,  # 행 노드 높이 — 같은 계열은 균일(높이가 섞이면 위계로 오독된다)
    "ok_dy": 0.12,  # 상단 갈래(확정) 행 top
    "chip_dx": 3.70,  # 2단계 칩(확정 / 부족) 좌변
    "chip_w": 1.25,
    "act_dx": 5.35,  # 3단계 [행동] 좌변
    "act_w": 2.35,
    "err_dx": 8.10,
    "err_w": 1.90,
    "term_dx": 10.40,
    "cap_dy": 0.04,  # 행 bottom → 하위 캡션 top
    "cap_h": 0.22,
    "row_air": 0.14,  # 캡션 bottom → 다음 행 top
    "row1_dy": 0.76,  # 하단 갈래 첫 행 top
    "back_dy": 2.48,  # 복귀 호의 수평 구간
    "back_label_dx": 1.80,
    "back_label_dy": 2.24,
    "back_label_w": 3.50,
    "label_gap": 0.06,  # ◇ 우변 → 갈래 라벨
}


def measure(data, kit, layout=None):  # noqa: ARG001 — 계약 시그니처 고정
    """필요 폭·높이(inch). 배치하는 쪽이 자리를 손으로 재지 않게 한다."""
    L = merged(LAYOUT, layout)
    return L["term_dx"] + L["q_w"], L["back_dy"] + 0.02


def draw(slide, box, data, kit, layout=None):
    require(data, META["keys"], META["name"])
    box = Box(*box)  # 4튜플도 받는다 — 목업 러너·장이 둘 다 쓴다
    L = merged(LAYOUT, layout)
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    ink_note = mix(C["muted"], C["primary"], 0.55)

    rows = data["rows"]
    if len(rows) != 2:
        raise ValueError(
            f"gate_branch: 갈래는 2 고정인데 {len(rows)}개가 들어왔다 — "
            "두 갈래의 차이가 자리가 아니라 의미(종착 vs 복귀)라 늘릴 수 없다"
        )

    x0, y0 = box.x, box.y
    q_h, row_h = L["q_h"], L["row_h"]
    dia_right = x0 + L["dia_dx"] + L["dia_w"]
    chip_x = x0 + L["chip_dx"]
    chip_right = chip_x + L["chip_w"]

    # 행 y — 캡션·공기에서 파생한다(좌표 리터럴 2개를 zip으로 물리지 않는다)
    row_pitch = row_h + L["cap_dy"] + L["cap_h"] + L["row_air"]
    row_ys = [y0 + L["row1_dy"] + i * row_pitch for i in range(2)]
    row_cys = [ry + row_h / 2 for ry in row_ys]

    ok_cy = y0 + L["ok_dy"] + row_h / 2
    lack_cy = sum(row_cys) / 2
    dia_cy = (ok_cy + lack_cy) / 2

    # ── 질문 → ◇판정 ──
    q = add_box(
        slide,
        x0,
        dia_cy - q_h / 2,
        L["q_w"],
        q_h,
        fill=C["bg_alt"],
        line=C["muted"],
        line_w=0.75,
        shape="round",
    )
    E.center_text(q, kit, data["question"], S["caption"], C["primary"])
    E.arrow(slide, kit, x0 + L["q_w"], dia_cy, x0 + L["dia_dx"], dia_cy)
    dia = add_box(
        slide,
        x0 + L["dia_dx"],
        dia_cy - L["dia_h"] / 2,
        L["dia_w"],
        L["dia_h"],
        fill=C["bg"],
        line=C["primary"],
        line_w=1.25,
        shape="diamond",
    )
    E.center_text(dia, kit, data["gate"], S["caption"], C["primary"])

    # ── 두 갈래 ──
    E.arrow(slide, kit, dia_right, dia_cy, chip_x, ok_cy, elbow=True)
    E.arrow(slide, kit, dia_right, dia_cy, chip_x, lack_cy, elbow=True)
    lx = dia_right + L["label_gap"]
    add_text(
        slide,
        lx,
        ok_cy - 0.30,
        0.62,
        0.22,
        data["branch_ok"],
        S["caption"],
        F["head"],
        C["muted"],
        bold=True,
    )
    add_text(
        slide,
        lx,
        lack_cy + L["label_gap"],
        0.62,
        0.22,
        data["branch_lack"],
        S["caption"],
        F["head"],
        C["primary"],
        bold=True,
    )

    ok = add_box(
        slide,
        chip_x,
        y0 + L["ok_dy"],
        L["chip_w"],
        row_h,
        fill=C["bg_alt"],
        line=C["muted"],
        line_w=0.75,
        shape="round",
    )
    E.center_text(ok, kit, data["ok_chip"], S["body"], C["primary"])
    if data.get("ok_note"):
        add_text(
            slide,
            x0 + L["act_dx"],
            y0 + L["ok_dy"] - 0.02,
            box.w - L["chip_dx"] - L["chip_w"] - 0.4,
            row_h,
            data["ok_note"],
            S["caption"],
            F["body"],
            ink_note,
            line_spacing=1.3,
        )
    lack = add_box(
        slide,
        chip_x,
        lack_cy - row_h / 2,
        L["chip_w"],
        row_h,
        fill=C["bg"],
        line=C["primary"],
        line_w=1.25,
        shape="round",
    )
    E.center_text(lack, kit, data["lack_chip"], S["body"], C["primary"])
    for cy in row_cys:
        E.arrow(slide, kit, chip_right, lack_cy, x0 + L["act_dx"], cy, elbow=True)

    # ── 갈래 뒤 3단계: 행동 → 오류 → (종착 | 복귀) ──
    # 첫 행만 종착까지 간다. 그 대비가 이 도해의 주장이므로 색·도형도 함께 갈린다.
    act_x, err_x, term_x = x0 + L["act_dx"], x0 + L["err_dx"], x0 + L["term_dx"]
    term_w = box.w - L["term_dx"]
    for i, (ry, row) in enumerate(zip(row_ys, rows, strict=True)):
        fatal = i == 0
        err_fill = C["accent"] if fatal else C["bg_alt"]
        err_ink = C["bg"] if fatal else C["primary"]
        act = add_box(
            slide,
            act_x,
            ry,
            L["act_w"],
            row_h,
            fill=C["bg_alt"],
            line=C["muted"],
            line_w=0.75,
            shape="round",
        )
        E.center_text(act, kit, row["act"], S["caption"], C["primary"])
        rcy = ry + row_h / 2
        E.arrow(slide, kit, act_x + L["act_w"], rcy, err_x, rcy)
        err = add_box(slide, err_x, ry, L["err_w"], row_h, fill=err_fill, shape="round")
        E.center_text(err, kit, row["err"], S["caption"], err_ink)
        if not fatal:
            continue
        E.arrow(slide, kit, err_x + L["err_w"], rcy, term_x, rcy)
        term = add_box(slide, term_x, ry, term_w, row_h, fill=C["accent"], shape="rect")
        E.center_text(term, kit, row["term"], S["caption"], C["bg"])
        cap_y = ry + row_h + L["cap_dy"]
        add_text(
            slide,
            term_x,
            cap_y,
            term_w,
            L["cap_h"],
            row["mark"],
            S["caption"],
            F["head"],
            C["primary"],
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            err_x,
            cap_y,
            term_x - err_x - 0.1,
            L["cap_h"],
            row["note"],
            S["caption"],
            F["body"],
            ink_note,
        )

    # ── 복귀 호 — 종착하지 않은 갈래는 질문으로 되돌아온다 ──
    back_x = err_x + L["err_w"] / 2
    q_cx = x0 + L["q_w"] / 2
    back_y = y0 + L["back_dy"]
    E.hairline(slide, kit, back_x, row_ys[-1] + row_h, back_x, back_y)
    E.hairline(slide, kit, back_x, back_y, q_cx, back_y)
    E.arrow(slide, kit, q_cx, back_y, q_cx, dia_cy + q_h / 2)
    if data.get("back_label"):
        add_text(
            slide,
            x0 + L["back_label_dx"],
            y0 + L["back_label_dy"],
            L["back_label_w"],
            L["cap_h"],
            data["back_label"],
            S["caption"],
            F["head"],
            C["primary"],
            bold=True,
        )


SAMPLES = [
    (
        "최소",
        {
            "question": "질문 한 건",
            "gate": "근거 충분?",
            "branch_ok": "충분",
            "branch_lack": "부족",
            "ok_chip": "확정",
            "ok_note": "근거를 붙여 확정 답변을 낸다.",
            "lack_chip": "근거 부족",
            "rows": [
                {
                    "act": "그래도 답한다",
                    "err": "허위 확정",
                    "term": "되돌릴 수 없음",
                    "mark": "종착",
                    "note": "틀린 줄 모르고 그대로 나간다",
                },
                {"act": "유보하고 묻는다", "err": "판단 보류"},
            ],
            "back_label": "다시 묻는다 — 고리가 닫힌다",
        },
    ),
]
