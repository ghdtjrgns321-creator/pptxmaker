"""deckkit 블록 8종 — 사용자가 2026-08-03에 "좋다"고 고른 것에서 뽑았다.

공통 문법 하나: **모든 것을 칸에 가두고, 칸마다 정해진 슬롯에 필요한 것만 넣는다.**

블록은 `(s, x, y, w, h, 데이터)`를 받아 그 사각형 안만 그린다. **자기 자리를 모른다** —
어디에 놓을지는 장이 정한다. 그래서 같은 블록을 써도 장마다 다른 배치가 나온다.

한 유형에 변형이 하나뿐이면 전 장이 같아진다(2026-07-29 실측: 36장 동일 카드). 변형이
필요하면 인자로 열되, 인자가 형태를 바꾸는지 치수만 바꾸는지 구분할 것.
"""

from .kit import (
    AMBER,
    AMBER_LT,
    BAR,
    BLUE,
    BLUE_LT,
    GRAY,
    INK,
    LIM,
    LINE,
    PANEL,
    R,
    SIZES,
    T,
    WARN,
    WHITE,
    ARROW,
    CUR,
)

S = SIZES


# ① KPI 스트립 — 큰 수치를 나란히. 분모·부연을 같이 적어 수치가 혼자 서지 않게 한다.
def kpi_strip(s, x, y, w, h, items, gap=0.10, fill=BLUE_LT, accent=INK):
    """items = [(라벨, 값, 부연|None), ...]"""
    n = len(items)
    cw = (w - gap * (n - 1)) / n
    for i, it in enumerate(items):
        lab, val = it[0], it[1]
        sub = it[2] if len(it) > 2 else None
        cx = x + i * (cw + gap)
        R(s, cx, y, cw, h, fill=fill)
        lines = [
            {"t": lab, "s": S["tiny"], "c": GRAY},
            {"t": val, "s": S["kpi"], "b": True, "c": accent, "sb": 2},
        ]
        if sub:
            lines.append({"t": sub, "s": S["tiny"], "c": GRAY, "sb": 2})
        T(s, cx + 0.12, y + 0.10, cw - 0.24, h - 0.20, lines, name=f"kpi:{lab}")


# ② 슬롯 카드 목록 — 카드마다 [제목 · 부제 · 산출물 · 근거]. 슬롯이 고정이라 비교가 된다.
def slot_cards(
    s, x, y, w, h, items, gap=0.10, cols=("묻는 것", "산출물", "근거"), ratios=(1.0, 1.0, 0.95)
):
    """items = [(제목, 묻는 것, 산출물, 근거), ...] — **행 = 항목, 칸 = 슬롯.**

    한 줄씩 쌓으면 카드 오른쪽이 통째로 빈다(2026-08-03 지적). 2열로 나눠도 같은 문제가
    반씩 생겼다. 그래서 **폭 전체를 슬롯으로 쪼개고 슬롯마다 라벨을 단다** — 칸이 내용보다
    넓어질 수 없고, 같은 슬롯끼리 세로로 비교된다.
    """
    n = len(items)
    ch = (h - gap * (n - 1)) / n
    inner = w - 0.36
    tw = inner * 0.30  # 제목 몫
    sw = inner - tw
    unit = sw / sum(ratios)
    for i, it in enumerate(items):
        cy = y + i * (ch + gap)
        R(s, x, cy, w, ch, fill=WHITE, line=LINE)
        R(s, x, cy, 0.05, ch, fill=BLUE)
        T(
            s,
            x + 0.18,
            cy + 0.06,
            tw - 0.12,
            ch - 0.12,
            [{"t": it[0], "s": S["h"], "b": True, "c": INK, "ls": 1.18}],
            valign="m",
            name=f"card:{it[0]}",
        )
        cx = x + 0.18 + tw
        for j, lab in enumerate(cols):
            val = it[j + 1] if len(it) > j + 1 else ""
            cwid = unit * ratios[j]
            if j:
                R(s, cx - 0.06, cy + 0.14, 0.008, ch - 0.28, fill=LINE)
            T(
                s,
                cx,
                cy + 0.06,
                cwid - 0.14,
                ch - 0.12,
                [
                    {"t": lab, "s": S["tiny"], "c": GRAY},
                    {
                        "t": val,
                        "s": S["small"],
                        "b": j == 1,
                        "c": INK if j == 1 else GRAY,
                        "sb": 3,
                        "ls": 1.2,
                    },
                ],
                valign="m",
                name=f"슬롯:{lab}",
            )
            cx += cwid


# ③ 번호 절차 — 순서 자체가 메시지일 때. 배지 + 제목 + 설명 한 줄.
def numbered_steps(s, x, y, w, h, items, lab_w=1.55, hot=()):
    """items = [(제목, 설명), ...]"""
    n = len(items)
    rh = h / n
    for i, (title, desc) in enumerate(items):
        ry = y + i * rh
        on = i in hot
        R(s, x, ry + 0.02, 0.30, min(0.30, rh - 0.06), fill=BLUE if on else INK)
        T(
            s,
            x,
            ry + 0.02,
            0.30,
            min(0.30, rh - 0.06),
            [{"t": str(i + 1), "s": S["small"], "b": True, "c": WHITE, "a": "c"}],
            valign="m",
            name="번호",
        )
        T(
            s,
            x + 0.42,
            ry,
            lab_w,
            rh,
            [{"t": title, "s": S["h"], "b": True, "c": INK}],
            valign="m",
            name=f"step:{title}",
        )
        T(
            s,
            x + 0.42 + lab_w,
            ry,
            w - 0.42 - lab_w,
            rh,
            [{"t": desc, "s": S["small"], "c": GRAY, "ls": 1.15}],
            valign="m",
            name="설명",
        )


# ④ 분기 구조 — 하나에서 N으로. 갈래마다 [제목 · 질문→산출]을 붙여 갈래가 뭘 하는지 보인다.
def branch(s, x, y, w, h, root, leaves, node_w=1.85, gap=0.10):
    """leaves = [(제목, 설명), ...]"""
    n = len(leaves)
    rh = (h - gap * (n - 1)) / n
    R(s, x, y + h / 2 - 0.34, node_w, 0.68, fill=INK)
    T(
        s,
        x + 0.08,
        y + h / 2 - 0.34,
        node_w - 0.16,
        0.68,
        [{"t": root, "s": S["h"], "b": True, "c": WHITE, "a": "c", "ls": 1.15}],
        valign="m",
        name="root",
    )
    bx = x + node_w + 0.30
    R(s, x + node_w + 0.14, y + rh / 2, 0.012, h - rh, fill=LINE)
    for i, (title, desc) in enumerate(leaves):
        ly = y + i * (rh + gap)
        R(s, x + node_w + 0.14, ly + rh / 2, 0.16, 0.012, fill=LINE)
        T(
            s,
            bx,
            ly,
            w - (bx - x),
            rh,
            [
                {"t": title, "s": S["h"], "b": True, "c": INK},
                {"t": desc, "s": S["small"], "c": GRAY, "sb": 2, "ls": 1.15},
            ],
            valign="m",
            name=f"leaf:{title}",
        )


# ⑤ 문제→해결 2박스 — 장 머리에서 대립을 세운다.
def problem_solution(s, x, y, w, h, problem, solution, gap=0.42):
    bw = (w - gap - 0.55) / 2
    for i, (cap, txt, fill, line, col) in enumerate(
        [("문제 인식", problem, PANEL, LINE, INK), ("해결 방향", solution, BLUE_LT, BLUE, BLUE)]
    ):
        bx = x + i * (bw + gap + 0.55)
        R(s, bx, y, bw, h, fill=fill, line=line)
        T(
            s,
            bx + 0.16,
            y + 0.08,
            bw - 0.32,
            h - 0.16,
            [
                {"t": cap, "s": S["tiny"], "c": GRAY},
                {"t": txt, "s": S["h"], "b": True, "c": col, "sb": 3, "ls": 1.2},
            ],
            valign="m",
            name=f"ps:{cap}",
        )
    ARROW(s, x + bw + 0.16, y + h / 2 - 0.11, 0.55, 0.22)


# ⑥ 비교표 — 행 높이에 상한이 있다. 행이 높아지면 셀 안이 비어 보인다(2026-08-03 지적).
def compare_table(s, x, y, w, h, header, rows, lab_r=0.26, right_hot=True):
    """header = (구분, 좌, 우) · rows = [(축, 좌, 우), ...]"""
    hh = 0.34
    n = len(rows)
    rh = (h - hh) / n
    if rh > LIM["table_row_warn"]:
        # 누르지 않는다 — 강제로 줄이면 긴 셀 텍스트가 잘린다. 알리기만 하고 판단은 눈이 한다.
        WARN.append(
            f'S{CUR["n"]:02d} 표 행 높이 {rh:.2f}" > {LIM["table_row_warn"]}"'
            f' — 셀 안이 비어 보인다. 표 높이를 {hh + n * LIM["table_row_warn"]:.2f}"로 줄이고'
            " 남는 세로를 다른 블록에 줄 것"
        )
    lw = w * lab_r
    cw = (w - lw) / 2
    R(s, x, y, w, hh, fill=INK)
    for i, t in enumerate(header):
        T(
            s,
            x + (0 if i == 0 else lw + (i - 1) * cw) + 0.14,
            y,
            (lw if i == 0 else cw) - 0.28,
            hh,
            [{"t": t, "s": S["small"], "b": True, "c": WHITE}],
            valign="m",
            name="표머리",
        )
    for r, (ax, lt, rt) in enumerate(rows):
        ry = y + hh + r * rh
        if r % 2 == 0:
            R(s, x, ry, w, rh, fill=PANEL)
        if right_hot:
            R(s, x + lw + cw, ry, cw, rh, fill=BLUE_LT)
        for i, (t, bold, col) in enumerate([(ax, True, INK), (lt, False, GRAY), (rt, True, INK)]):
            T(
                s,
                x + (0 if i == 0 else lw + (i - 1) * cw) + 0.14,
                ry,
                (lw if i == 0 else cw) - 0.28,
                rh,
                [{"t": t, "s": S["body"], "b": bold, "c": col, "ls": 1.15}],
                valign="m",
                name=f"셀:{ax}",
            )


# ⑦ 값 막대 + 설명 — 값은 길이로, 옆에 항목을 편다.
def value_bars(s, x, y, w, h, items, lab_w=1.35, val_w=0.42, bar_w=1.60, note_w=None):
    """items = [(라벨, 값표시, 비율 0~1, 설명|None), ...]"""
    n = len(items)
    rh = h / n
    nw = note_w if note_w is not None else w - lab_w - val_w - bar_w - 0.24
    for i, it in enumerate(items):
        lab, val, frac = it[0], it[1], it[2]
        note = it[3] if len(it) > 3 else None
        ry = y + i * rh
        hot = frac >= 0.999
        T(
            s,
            x,
            ry,
            lab_w,
            rh,
            [{"t": lab, "s": S["small"], "b": True, "c": INK}],
            valign="m",
            name=f"bar:{lab}",
        )
        BAR(s, x + lab_w, ry + rh / 2 - 0.07, bar_w, 0.14, frac, AMBER if hot else BLUE, track=LINE)
        T(
            s,
            x + lab_w + bar_w + 0.10,
            ry,
            val_w,
            rh,
            [{"t": val, "s": S["small"], "b": True, "c": AMBER if hot else BLUE, "a": "r"}],
            valign="m",
            name="값",
        )
        if note and nw > 0.3:
            T(
                s,
                x + lab_w + bar_w + val_w + 0.24,
                ry,
                nw,
                rh,
                [{"t": note, "s": S["tiny"], "c": GRAY, "ls": 1.15}],
                valign="m",
                name="설명",
            )


# ⑧ 격자 카드 — 항목이 많을 때. 코드 + 이름 + 태그. 전수를 보이는 게 목적이라 접지 않는다.
def card_grid(s, x, y, w, h, items, ncol=2, gap=0.10, tag_colors=None):
    """items = [(코드, 이름, 태그|None), ...]"""
    n = len(items)
    nrow = -(-n // ncol)
    cw = (w - gap * (ncol - 1)) / ncol
    ch = (h - gap * (nrow - 1)) / nrow
    for i, it in enumerate(items):
        code, name = it[0], it[1]
        tag = it[2] if len(it) > 2 else None
        r, c = divmod(i, ncol)
        cx, cy = x + c * (cw + gap), y + r * (ch + gap)
        R(s, cx, cy, cw, ch, fill=WHITE, line=LINE)
        col = (tag_colors or {}).get(tag, BLUE)
        R(s, cx, cy, 0.045, ch, fill=col)
        T(
            s,
            cx + 0.14,
            cy,
            cw * 0.30,
            ch,
            [{"t": code, "s": S["tiny"], "b": True, "c": col}],
            valign="m",
            name="코드",
        )
        T(
            s,
            cx + 0.14 + cw * 0.30,
            cy,
            cw * 0.44,
            ch,
            [{"t": name, "s": S["small"], "c": INK}],
            valign="m",
            name=f"격자:{code}",
        )
        if tag:
            T(
                s,
                cx + cw - 0.14 - cw * 0.22,
                cy,
                cw * 0.22,
                ch,
                [{"t": tag, "s": S["tiny"], "c": GRAY, "a": "r"}],
                valign="m",
                name="태그",
            )
