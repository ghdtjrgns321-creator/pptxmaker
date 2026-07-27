"""G05 이분 매핑 — A 집합의 말이 B 집합의 말로 어떻게 갈라지고 모이나(N:M).

출처: `s08_variants.bipartite_map`(골든 s08 실물)을 부품 계약 3조건으로 옮긴 것.
좌표 기본값은 골든 실측 그대로라 장 교체 후에도 스냅샷이 동일해야 한다.

## 이 도해가 지키는 것 (원본에서 계승 · 이유까지)

- **정렬은 폭이 아니라 한쪽 변으로** — 좌열은 우변 고정, 우열은 좌변 고정. 균일 폭으로
  그렸더니 4자짜리가 3.45" 칩에서 채움률 12%였다(죽은 회색). 폭은 글자에서 파생한다.
- **우열은 데이터에서 파생** — 손으로 적으면 N:1 수렴이 "내가 그렇게 그렸다"가 된다.
  파생하면 두 선이 한 개념에 모이는 게 데이터가 그렇게 생겼기 때문이 된다.
- **N:1 도착점 분산** — 한 개념에 두 선이 오면 도착 y를 칩 높이 안에서 등분한다.
  한 점에 모으면 점선 화살촉이 실선 화살촉에 삼켜져 "위임판단"이 실선으로 보인다
  (2026-07-15 제3자 채점 FAIL #6 — 수렴이 요점인 장에서 바로 그 지점의 형태가 무너졌다).
- **선 먼저 → 칩 나중** — 칩이 화살촉을 덮어 칩 변에서 깔끔히 끝난다.
"""

from ..kit import add_box, set_shape_text
from . import Box, merged, require
from .elements import chip, link

META = {
    "name": "bipartite_map",
    "어휘": "G05",  # archetype-catalog.md 등재 어휘와 대응(유일하게 정합한 건)
    "물성": "A 집합과 B 집합이 어떻게 갈라지고 모이나(N:M 대응·수렴)",
    "arity": "파생",  # 좌·우 항목 수 모두 파생. 상한은 pitch가 시끄럽게 알린다
    "keys": ["pairs"],
    "optional": ["soft"],  # pairs[i].soft = 점선+accent(불확실·위임 판단)
}

# 이 도해는 **고유 크기**를 갖는다: 칩 폭은 글자에서, 높이는 항목 수에서 나온다.
# 자리를 넓힌다고 칩이 커지지 않으므로 폭을 자리에 비례시키면 **선만 길어지고 좌우가 빈다**
# (2026-07-26 목업 눈검증). 그래서 자리는 "얼마나 크게 그리나"가 아니라 **"어디에 앉히나"**를
# 정한다 — 고유 크기로 그리고 남는 자리는 정렬로 처리하며, 모자라면 시끄럽게 죽는다.
LAYOUT = {
    "l_col": None,  # 좌열 폭. None=좌칩 최대폭에서 파생 · 숫자=고정(골든 장 재현용)
    "r_col": None,  # 우열 폭. None=우칩 최대폭에서 파생
    "gap": None,  # 좌열↔우열 거리(=선 길이). None=자리에서 파생(gap_min~gap_max) · 숫자=고정
    "gap_min": 1.10,  # 이보다 짧으면 관계선이 아니라 붙은 두 칩으로 읽힌다
    "gap_max": 3.50,  # 이보다 길면 좌우가 따로 노는 두 덩이가 된다(눈이 선을 못 따라간다)
    "halign": "center",  # center | left — 자리가 남을 때 가로 위치
    "valign": "center",  # center | top   — 자리가 남을 때 세로 위치
    "chip_h": 0.24,  # 칩 높이 — 계열 균일(글자 크기에 묶인 치수라 자리와 무관)
    "inset": 0.44,  # 칩 글자 좌우 여백
    "floor": 0.95,  # 최소 칩 폭 — 3자짜리도 칩으로 보이게
    "line_w": 1.0,  # 잇기 선 굵기(pt)
    # 방향. "h"=좌→우(라벨이 길 때 — 글자가 가로로 자라므로 열 폭에 여유가 있다)
    #        "v"=위→아래(항목이 많고 라벨이 짧을 때 — 가로로 늘어놓아야 개수가 보인다)
    # 같은 문법의 방향 변형이라 별도 부품으로 두지 않는다(부품이 갈라지면 규칙도 갈라진다).
    "orient": "h",
    "row_gap": 0.30,  # (세로) 두 행 사이 최소 거리
    "row_gap_max": 3.00,
    "col_gap": 0.18,  # (세로) 같은 행 안 칩 사이 최소 간격
    "chip_h_v": 0.34,  # (세로) 칩 높이 — 가로판보다 두껍다. 세로판은 칩이 크게 보이는 문법이다
}


def chip_w(text, pt, inset, floor):
    """칩 폭을 글자에서 파생. 근사식은 오딧과 같은 자(한국어 1자 ≈ 0.8×pt).

    폭을 자리에 맞춰 **자르지 않는다**. 옛 코드는 `min(w, limit)`으로 잘랐는데 폭만 줄고
    글자는 그대로라 칩이 2줄이 되어 **글자가 칩 높이(0.24") 밖으로 넘쳤다**(2026-07-26 목업
    눈검증 — 상자는 자리 안이라 기하 게이트를 통과한다). 계열 균일이 이 도해의 규칙이라
    높이를 늘리거나 폰트를 줄이는 것도 답이 아니다. 안 들어가면 `measure`가 죽인다.
    """
    return max(len(text.strip()) * pt * 0.8 / 72 + inset, floor)


def measure(data, kit, layout=None):
    """이 데이터를 그리려면 얼마나 필요한가 — (필요 폭, 필요 높이, 좌열 폭, 우열 폭).

    자리보다 먼저 크기가 정해진다. 배치하는 쪽(장·컴포저)이 "이만큼 주면 되겠다"를 알 수 있어야
    자리를 손으로 재지 않는다.
    """
    from .. import grid as G

    L = merged(LAYOUT, layout)
    pt, h = kit["sizes"]["caption"], L["chip_h"]
    pairs = data["pairs"]
    concepts = []
    for p in pairs:
        for c_ in p["right"]:
            if c_ not in concepts:
                concepts.append(c_)
    l_w = L["l_col"] or max(chip_w(p["left"], pt, L["inset"], L["floor"]) for p in pairs)
    r_w = L["r_col"] or max(chip_w(c_, pt, L["inset"], L["floor"]) for c_ in concepts)
    n = max(len(concepts), len(pairs))
    need_h = (n - 1) * G.CAP_PITCH + h
    return l_w + (L["gap"] or L["gap_min"]) + r_w, need_h, l_w, r_w


def _line(slide, kit, x1, y1, x2, y2, *, dashed, width):
    """잇기 선. 화살촉은 목적지 쪽 — 방향이 곧 '진입'이다."""
    from pptx.enum.shapes import MSO_CONNECTOR
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt

    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    conn.line.color.rgb = kit["rgb"]["muted"]
    conn.line.width = Pt(width)
    ln = conn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "sm", "len": "sm"}))
    return conn


def _row(kit, L, labels, x0, w_total):
    """한 행을 가로로 나열하고 {라벨: (중심 x, 폭)}을 준다 — 세로판의 위·아래 행 공용."""
    pt = kit["sizes"]["caption"]
    ws = [chip_w(t, pt, L["inset"], L["floor"]) for t in labels]
    need = sum(ws) + L["col_gap"] * (len(ws) - 1)
    if need > w_total:
        raise ValueError(
            f'세로판 한 행에 {len(labels)}개는 안 들어간다 — 필요 {need:.2f}" > 자리 {w_total:.2f}". '
            "라벨을 줄이거나 가로 방향(orient='h')으로 바꿀 것."
        )
    x = x0 + (w_total - need) / 2
    out = {}
    for t, w in zip(labels, ws):
        out[t] = (x + w / 2, w)
        x += w + L["col_gap"]
    return out


def _spread(desired, ws, x0, w_total, gap):
    """원하는 중심 x에서 출발해 **겹침만** 푼다 — 안 겹치면 원하는 자리 그대로 둔다.

    아래 행을 위 행과 따로 가운데 정렬하면 칩 폭이 달라 열 중심이 어긋나고, 1:1인데도
    선이 전부 사선으로 떨어진다(2026-07-26 사용자 반려). 아래 칩의 자리는 **자기를 가리키는
    위 칩들의 무게중심**에서 나와야 한다 — 그러면 1:1은 저절로 수직이 되고, N:1은 대칭 수렴이 된다.
    """
    cx = list(desired)
    for i in range(1, len(cx)):  # 왼→오: 앞 칩 밀어내기
        lo = cx[i - 1] + ws[i - 1] / 2 + gap + ws[i] / 2
        cx[i] = max(cx[i], lo)
    right = x0 + w_total
    if cx[-1] + ws[-1] / 2 > right:  # 오른쪽 벽을 넘었으면 오→왼으로 되민다
        cx[-1] = right - ws[-1] / 2
        for i in range(len(cx) - 2, -1, -1):
            cx[i] = min(cx[i], cx[i + 1] - ws[i + 1] / 2 - gap - ws[i] / 2)
    if cx[0] - ws[0] / 2 < x0 - 1e-6:
        need = sum(ws) + gap * (len(ws) - 1)
        raise ValueError(
            f'세로판 아래 행 {len(ws)}개가 자리에 안 들어간다 — 필요 {need:.2f}" > '
            f'자리 {w_total:.2f}". 라벨을 줄이거나 가로 방향(orient="h")으로 바꿀 것.'
        )
    return cx


def _draw_v(slide, box, pairs, lefts, concepts, kit, L):
    """세로판 — 위 행에서 아래 행으로. 항목이 많고 라벨이 짧을 때 개수가 눈에 들어온다."""
    h = L["chip_h_v"]
    gap = min(L["row_gap_max"], max(L["row_gap"], box.h - 2 * h))
    y_top = box.y + (box.h - (2 * h + gap)) / 2 if L["valign"] == "center" else box.y
    y_bot = y_top + h + gap
    pt = kit["sizes"]["caption"]
    src = {c_: [p["left"] for p in pairs if c_ in p["right"]] for c_ in concepts}

    if len(concepts) == len(pairs) and all(len(v) == 1 for v in src.values()):
        # 1:1이면 **공용 트랙**에 두 행을 같이 앉힌다 — 열 폭 = 위·아래 중 넓은 쪽.
        # 각 행을 따로 가운데 정렬하면 칩 폭이 달라 열 중심이 어긋나고, 1:1인데도 선이
        # 사선으로 떨어진다(2026-07-26 사용자 반려 — 수직이어야 대응이 대응으로 읽힌다).
        downs = [p["right"][0] for p in pairs]
        cols = [
            max(
                chip_w(t, pt, L["inset"], L["floor"]),
                chip_w(c_, pt, L["inset"], L["floor"]),
            )
            for t, c_ in zip(lefts, downs, strict=True)
        ]
        need = sum(cols) + L["col_gap"] * (len(cols) - 1)
        if need > box.w:
            raise ValueError(
                f'세로판 {len(cols)}열이 자리에 안 들어간다 — 필요 {need:.2f}" > 자리 {box.w:.2f}". '
                "라벨을 줄이거나 가로 방향(orient='h')으로 바꿀 것."
            )
        x = box.x + (box.w - need) / 2
        up, dn = {}, {}
        for t, c_, cw in zip(lefts, downs, cols, strict=True):
            up[t] = (x + cw / 2, chip_w(t, pt, L["inset"], L["floor"]))
            dn[c_] = (x + cw / 2, chip_w(c_, pt, L["inset"], L["floor"]))
            x += cw + L["col_gap"]
    else:
        # N:M — 아래 행 x는 **자기를 가리키는 위 칩들의 평균**에서. 그다음 겹침만 푼다.
        up = _row(kit, L, lefts, box.x, box.w)
        order = sorted(concepts, key=lambda c_: sum(up[t][0] for t in src[c_]) / len(src[c_]))
        ws_d = [chip_w(c_, pt, L["inset"], L["floor"]) for c_ in order]
        want = [sum(up[t][0] for t in src[c_]) / len(src[c_]) for c_ in order]
        dn = dict(
            zip(order, zip(_spread(want, ws_d, box.x, box.w, L["col_gap"]), ws_d), strict=True)
        )

    fan_in = {}
    for p in pairs:
        for c_ in p["right"]:
            fan_in[c_] = fan_in.get(c_, 0) + 1
    slot = dict.fromkeys(fan_in, 0)
    n_lines = 0
    for p in pairs:
        for c_ in p["right"]:
            n_in = fan_in[c_]
            off = (slot[c_] + 1) / (n_in + 1) if n_in > 1 else 0.5
            slot[c_] += 1
            cx_d, w_d = dn[c_]
            link(
                slide,
                kit,
                up[p["left"]][0],
                y_top + h,
                cx_d - w_d / 2 + w_d * off,  # 도착점을 칩 폭 안에서 분산 — 수렴이 뭉치지 않게
                y_bot,
                dashed=bool(p.get("soft")),
                width=L["line_w"],
            )
            n_lines += 1
    for p in pairs:
        cx_u, w_u = up[p["left"]]
        chip(
            slide,
            cx_u - w_u / 2,
            y_top,
            w_u,
            h,
            p["left"],
            kit,
            style="soft" if p.get("soft") else "outline",
        )
    for c_ in concepts:
        cx_d, w_d = dn[c_]
        chip(
            slide, cx_d - w_d / 2, y_bot, w_d, h, c_, kit, style="plain", font=kit["fonts"]["body"]
        )
    return up, dn, n_lines


def draw(slide, box, data, kit, layout=None):
    """box 안에 이분 매핑을 그린다. 반환: (좌 y맵, 우 y맵, 선 수) — 오딧·자기검증용.

    data = {"pairs": [{"left": str, "right": [str, ...], "soft": bool}, ...]}
    """
    from .. import grid as G

    require(data, META["keys"], META["name"])
    L = merged(LAYOUT, layout)
    box = Box(*box)
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    pt, h = S["caption"], L["chip_h"]

    pairs = data["pairs"]
    lefts = [p["left"] for p in pairs]
    concepts = []  # 등장순 dedup — 우열은 데이터에서 파생한다
    for p in pairs:
        for c_ in p["right"]:
            if c_ not in concepts:
                concepts.append(c_)

    if L["orient"] == "v":
        return _draw_v(slide, box, pairs, lefts, concepts, kit, L)

    # ── 가로: 고유 크기(칩 폭) + 선 길이. 자리가 남으면 선을 gap_max까지만 늘리고 정렬로 처리
    _, _, l_w, r_w = measure(data, kit, layout)
    free = box.w - l_w - r_w
    gap = L["gap"] if L["gap"] is not None else max(L["gap_min"], min(L["gap_max"], free))
    if free < L["gap_min"]:
        raise ValueError(
            f'이분 매핑이 이 자리에 안 들어간다 — 좌열 {l_w:.2f}" + 우열 {r_w:.2f}"만으로 '
            f'{l_w + r_w:.2f}"인데 자리는 {box.w:.2f}"다. 선 자리가 최소 {L["gap_min"]}" 필요하다. '
            "라벨을 줄이거나 자리를 넓힐 것."
        )
    x0 = box.x + (box.w - (l_w + gap + r_w)) / 2 if L["halign"] == "center" else box.x
    x_l = x0 + l_w  # 좌칩 **우변** — 선 출발점이 일정해야 매핑이 기둥으로 읽힌다
    x_r = x_l + gap  # 우칩 **좌변** — 선 도착점 일정

    # ── 세로: 피치는 cap(리듬 상한)까지만. 자리가 남으면 늘리지 않고 정렬한다
    top, bottom = box.y, box.y + box.h - h
    r_pitch = G.pitch(len(concepts), top, bottom, h, what="연결 개념")
    l_pitch = G.pitch(len(lefts), top, bottom, h, what="용어")
    span = max((len(concepts) - 1) * r_pitch, (len(lefts) - 1) * l_pitch) + h
    if L["valign"] == "center":
        top += max(0.0, (box.h - span) / 2)
    r_y = {c_: top + i * r_pitch for i, c_ in enumerate(concepts)}
    # 좌열은 우열 **범위 중앙**에 맞춘다 — 5개를 7개 옆에 위에서부터 쌓으면 아래가 빈다
    l_top = top + ((len(concepts) - 1) * r_pitch - (len(lefts) - 1) * l_pitch) / 2
    l_y = {t: l_top + i * l_pitch for i, t in enumerate(lefts)}

    fan_in = {}
    for p in pairs:
        for c_ in p["right"]:
            fan_in[c_] = fan_in.get(c_, 0) + 1
    slot = dict.fromkeys(fan_in, 0)

    n_lines = 0
    for p in pairs:
        for c_ in p["right"]:
            n_in = fan_in[c_]
            # 도착 y를 칩 높이 안에서 n_in등분 — 1개면 정중앙, 2개면 위·아래로 갈린다
            off = (slot[c_] + 1) / (n_in + 1) if n_in > 1 else 0.5
            slot[c_] += 1
            _line(
                slide,
                kit,
                x_l,
                l_y[p["left"]] + h / 2,
                x_r,
                r_y[c_] + h * off,
                dashed=bool(p.get("soft")),
                width=L["line_w"],
            )
            n_lines += 1

    for p in pairs:
        t, soft = p["left"], bool(p.get("soft"))
        w = chip_w(t, pt, L["inset"], L["floor"])
        b = add_box(
            slide,
            x_l - w,  # 우변 고정
            l_y[t],
            w,
            h,
            fill=C["bg"],
            line=C["accent"] if soft else C["muted"],
            line_w=1.25 if soft else 0.75,
            shape="round",
        )
        set_shape_text(b, t, pt, F["head"], C["primary"], bold=True)
    for c_ in concepts:
        b = add_box(
            slide,
            x_r,  # 좌변 고정
            r_y[c_],
            chip_w(c_, pt, L["inset"], L["floor"]),
            h,
            fill=C["bg_alt"],
            shape="round",
        )
        set_shape_text(b, c_, pt, F["body"], C["primary"])
    return l_y, r_y, n_lines


SAMPLES = [
    (
        "2개일 때 — 1:1 두 줄",
        {
            "pairs": [
                {"left": "리베이트", "right": ["고객에게 지급할 대가"]},
                {"left": "밀어내기", "right": ["위탁약정"]},
            ]
        },
    ),
    (
        "5개가 7개로 — N:M 갈림 + 수렴 2곳 · 점선은 불확실",
        {
            "pairs": [
                {"left": "리베이트", "right": ["고객에게 지급할 대가"]},
                {"left": "밀어내기", "right": ["위탁약정"]},
                {"left": "볼륨디스카운트", "right": ["변동대가", "변동대가 추정치를 제약함"]},
                {
                    "left": "상품권",
                    "right": ["고객에게 지급할 대가", "고객이 행사하지 아니한 권리"],
                    "soft": True,
                },
                {
                    "left": "반품의 회계처리",
                    "right": ["반품권이 있는 판매", "본인 대 대리인의 고려사항"],
                },
            ]
        },
    ),
    (
        "8개가 10개로 — 한 화면 밀도 상한",
        {
            "pairs": [
                {"left": "리베이트", "right": ["고객에게 지급할 대가"]},
                {"left": "밀어내기", "right": ["위탁약정"]},
                {"left": "볼륨디스카운트", "right": ["변동대가"]},
                {"left": "상품권", "right": ["고객이 행사하지 아니한 권리"], "soft": True},
                {"left": "반품", "right": ["반품권이 있는 판매"]},
                {"left": "중개 수수료", "right": ["본인 대 대리인의 고려사항"]},
                {"left": "무상 보증", "right": ["보증", "수행의무 식별"]},
                {"left": "선수금", "right": ["유의적 금융요소", "계약부채"], "soft": True},
            ]
        },
    ),
    (
        "세로판 — 위에서 아래로(항목이 많고 라벨이 짧을 때)",
        {
            "pairs": [
                {"left": "리베이트", "right": ["지급 대가"]},
                {"left": "밀어내기", "right": ["위탁약정"]},
                {"left": "볼륨할인", "right": ["변동대가"]},
                {"left": "상품권", "right": ["미행사 권리"], "soft": True},
                {"left": "반품", "right": ["반품권 판매"]},
            ],
        },
        {"orient": "v"},
    ),
    (
        "세로판 N:1 — 넷이 한 곳으로",
        {
            "pairs": [
                {"left": "리베이트", "right": ["지급 대가"]},
                {"left": "쿠폰", "right": ["지급 대가"]},
                {"left": "마일리지", "right": ["지급 대가"], "soft": True},
                {"left": "판촉물", "right": ["지급 대가"]},
            ],
        },
        {"orient": "v"},
    ),
    (
        "라벨이 길 때 — 폭 한계",
        {
            "pairs": [
                {
                    "left": "수익인식 5단계 적용",
                    "right": ["고객과의 계약에서 생기는 수익의 인식 및 측정 원칙을 정한 기준서"],
                },
                {"left": "계약변경", "right": ["계약변경을 별도 계약으로 회계처리하는 요건"]},
                {
                    "left": "거래가격",
                    "right": ["거래가격을 수행의무에 배분하는 방법"],
                    "soft": True,
                },
            ]
        },
    ),
    (
        "N:1일 때 — 네 갈래가 한 곳으로 모인다",
        {
            "pairs": [
                {"left": "리베이트", "right": ["고객에게 지급할 대가"]},
                {"left": "쿠폰", "right": ["고객에게 지급할 대가"]},
                {"left": "마일리지", "right": ["고객에게 지급할 대가"], "soft": True},
                {"left": "판촉물", "right": ["고객에게 지급할 대가"]},
            ]
        },
    ),
]
