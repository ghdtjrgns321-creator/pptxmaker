"""hub_spoke — 방사 관계. 하나가 중심에 있고 나머지가 그 둘레에 붙는다(1:N 소속 · 동등·동시).

이분 매핑(`bipartite_map`)과 다른 점: 저쪽은 **두 집합의 대응**이고 이쪽은 **한 중심의 소속**이다.
가로 나열(1→N 분기)과 다른 점: 분기는 **순서·경로**가 있고 방사는 **동등·동시**다.

## 이 부품은 **그래픽만** 그린다

설명 패널·결론 배너·열 구성은 **페이지**지 부품이 아니다. 부품은 `box` 안에 그래픽 하나만
놓고, "옆에 설명을 둘까"는 3층(배치)이 정한다.

## 잉크 규율 — MGI 실측 (2026-07-26)

`McKinsey Global Institute: 2025 in charts`(24p) 전 페이지를 렌더해 대조한 결과, 채운 도형으로
잉크를 만들려던 앞선 판들이 전부 틀렸다. 실물이 하는 것:

1. **면을 거의 안 쓴다** — 선과 여백이 그림의 대부분. 회색 채움 도형이 없다.
2. **색은 한 색 + 무채색** — 나머지는 흰 바탕.
3. **큰 수치 타이포가 주인공** — `1.3×` `37×` `27%`가 그림에서 제일 큰 잉크다.
   도형을 칠해서 무게를 만드는 게 아니라 **숫자가 무게를 진다**.
4. **테두리·그림자·둥근모서리 0.**

그래서 어느 모드도 면을 채우지 않는다 — `line`은 헤어라인 방사고, 나머지도 테두리까지다.
강조는 accent **선 하나와 글자**로만 한다(§2 accent 예산).

폐기 이력(전부 사용자 반려): 검정 채움 원 · 두꺼운 컬러 링 · 진한 파이 조각 · 회색 칩 ·
선 끝 도트 · 라벨 밑줄 룰 · 비중 막대 · 폴라 막대 · 버블 · 게이지 링.
**장식을 더하는 방향은 매번 틀렸다** — 이 파일에 잉크를 추가할 때는 반드시 눈검증을 먼저 받는다.

## 모드를 늘린 근거 — 실물 대조 (2026-07-26, part-design ②)

앞선 회차가 6시간·10라운드를 태우고 폐기된 원인은 잉크를 지어낸 것 말고 하나 더 있다:
**부품 하나에 시각 하나**라는 전제로 "번호 하나 고르기"를 시켰다는 것. 실물을 7종 깔아보니
사용자가 고른 것은 하나가 아니라 **셋**이었다 — 방사(`line`)와 나란히 갈 형태가 더 있었다.
1:N 소속을 그리는 방법은 하나가 아니고, 그러니 모드가 여럿인 게 정상이다.

    converge  ← OECD DAF/COMP(2019)14 p7   N개가 위에 나란히 서고 화살이 **한 점으로 모인다**
    scatter   ← Housing LIN 「Models of Hub and Spoke」 p4
                반지름이 **제각각**인 실선 화살이 중심에서 뻗고 끝마다 작은 테두리 노드

`line`과 갈리는 지점은 잉크 양이 아니라 **방향**이다. line은 방향이 없고(소속),
converge는 안쪽(귀속·집약), scatter는 바깥(파급·도달). 물성이 같아도 말하려는 방향이 다르면
다른 모드다. 반대로 3층 계층(RCA 실물)은 물성 자체가 달라서 이 부품에 넣지 않았다.

## 삭제 이력

`tick`(헤어라인 원 + 눈금)은 2026-07-26 눈검증에서 반려·삭제됐다. 실물 대조에서 근거가
나오지 않은 유일한 모드였다 — 내가 "잉크가 더 적다"는 이유로 만든 것이고, 실제로는 큰 타원
하나가 화면을 다 먹고 눈금은 보이지도 않았다. **실물 없이 만든 모드는 남기지 않는다.**
"""

import math

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from ..kit import add_text
from . import Box, merged, require
from .elements import text_w

META = {
    "name": "hub_spoke",
    "어휘": "미등재",  # 카탈로그 G22는 'accent 강조 패널'이다 — 그 번호를 쓰면 안 된다
    "물성": "하나에 무엇이 물려 있나(1:N 소속 · 동등·동시 · 중심이 주장)",
    "arity": "파생",
    # 셀 수 있는 선택 조건 — 컴포저는 재료를 세어 이것과 **비교만** 한다(산문 해석 없음).
    "accepts": {'sets': 1, 'flow': '1:N', 'order': False, 'extra': False, 'n': (3, 8), 'ends': False},
    "keys": ["hub", "spokes"],
    "optional": ["hub_note", "hub_metric", "metric", "soft"],
}

LAYOUT = {
    # line(헤어라인 방사) | converge(위→한 점 수렴) | scatter(불균등 방사)
    "mode": "line",
    "hair": 0.5,  # 헤어라인 굵기(pt) — 굵히면 선이 주인공이 되어 수치를 이긴다
    "lab_w": 2.10,  # 라벨 텍스트 폭(상자 없음 — 정렬만)
    "lab_gap": 0.16,  # 선 끝 ↔ 라벨
    "hub_gap": 0.22,  # 허브 글자 ↔ 선 시작
    "row_h": 0.52,  # 라벨 한 덩이 높이(이름 + 수치)
    "r_max_x": 4.30,  # 배치 반지름 상한 — 넘으면 선만 길어지고 한 덩어리로 안 읽힌다
    "r_max_y": 2.35,
    "start": -90,  # 첫 항목 각도(-90 = 12시), 시계방향
    "max_n": 8,
    # ── converge — 위 가로 나열 → 아래 한 점(OECD 실물)
    "cv_gap": 0.22,  # 위 칩 사이 간격
    "cv_min_w": 1.05,  # 칩 최소 폭 — 이하면 글자가 세로로 쪼개진다(track이 죽인다)
    "cv_chip_h": 0.40,  # 위 칩 높이
    "cv_drop_min": 0.55,  # 칩 아래변 ↔ 허브 윗변, 화살 구간 하한 — 이하면 화살이 아니라 붙은 두 층
    "cv_drop_max": 1.55,  # 상한 — 자리가 커도 화살을 더 늘리지 않는다. 늘리면 화살만 남는다
    #                       (2026-07-26 눈검증: 5.9" 자리에서 화살이 4.5"가 되어 그림이 전부 선)
    "cv_head_gap": 0.07,  # 화살촉 ↔ 허브 윗변. 0이면 촉이 테두리에 물려 뭉갠다
    # ── scatter — 반지름 제각각 방사(Housing LIN 실물)
    "sc_var": 0.22,  # 반지름 불균등 진폭(0=균등, 1=최대). 실물의 "제각각"이 이 부품의 물성이다
    "sc_node_h": 0.30,  # 끝 노드 높이 — 타이포에 묶인 치수라 자리에 비례시키지 않는다
    "sc_inset": 0.30,  # 끝 노드 글자 좌우 여백
    # 반지름 상한이 line보다 **작다**. 노드가 상자를 갖고 있어서, 같은 반지름이면
    # 맨 글자 라벨보다 훨씬 넓게 퍼져 한 덩어리로 안 읽힌다(2026-07-26 눈검증: 12.1" 자리에서
    # 노드 5개가 화면 네 귀퉁이로 흩어졌다). 남는 가로는 3층이 설명을 두는 자리다 — measure()로 알린다.
    "sc_r_max_x": 2.55,
    "sc_r_max_y": 1.85,
    "sc_hub_pt": "section",  # 중심 수치 크기 키(sizes). display(40pt)는 9pt 노드와 위계가 벌어진다
}


def _hair(slide, kit, x1, y1, x2, y2, *, accent=False, w=None, dashed=False):
    """헤어라인 — 화살촉 없음. 방향이 아니라 **소속**을 말하는 선이라 촉이 필요 없다."""
    from pptx.enum.shapes import MSO_CONNECTOR
    from pptx.oxml.ns import qn

    C = kit["rgb"]
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    conn.line.color.rgb = C["accent"] if accent else C["muted"]
    conn.line.width = Pt(w or LAYOUT["hair"])
    if dashed:
        ln = conn.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    return conn


def _hub_text(slide, kit, cx, cy, data, L, *, metric_key="display"):
    """중심 — 상자도 원도 없다. **수치가 제일 큰 잉크**이고 그 아래 이름·보조줄이 붙는다.

    `metric_key`는 수치 크기를 고르는 `sizes` 키다. 기본 display(40pt)는 라벨이 맨 글자인
    line 기준이고, 노드가 상자를 가진 scatter에서는 위계가 너무 벌어져 중심만 보인다.
    """
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    m_pt = S[metric_key]
    paras = []
    if data.get("hub_metric"):
        paras.append([(data["hub_metric"], {"size": m_pt, "bold": True, "color": C["primary"]})])
    paras.append([(data["hub"], {"size": S["head"], "bold": True, "color": C["primary"]})])
    if data.get("hub_note"):
        paras.append([(data["hub_note"], {"size": S["caption"], "color": C["muted"]})])
    h = 0.30 * len(paras) + (m_pt / 118 if data.get("hub_metric") else 0)
    w = (
        max(
            text_w(data.get("hub_metric") or "", m_pt),
            text_w(data["hub"], S["head"]),
            text_w(data.get("hub_note") or "", S["caption"]),
        )
        + 0.20
    )
    add_text(
        slide,
        cx - w / 2,
        cy - h / 2,
        w,
        h,
        paras,
        S["head"],
        F["head"],
        C["primary"],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
        line_spacing=1.06,
    )
    return w, h


def _label(slide, kit, px, py, ang, s, L):
    """항목 — 상자 없음. 방향에 따라 좌/우 정렬만 바꾼다. 수치가 있으면 수치가 더 크다.

    12시·6시에서는 선이 글자를 관통하므로 라벨 덩이를 **선 끝 너머로** 밀어낸다
    (2026-07-26 눈검증 — "용어 정규화" 가운데를 세로선이 뚫었다).
    """
    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    soft = bool(s.get("soft"))
    ink = C["accent"] if soft else C["primary"]
    ca, sa = math.cos(ang), math.sin(ang)
    align = PP_ALIGN.LEFT if ca > 0.20 else (PP_ALIGN.RIGHT if ca < -0.20 else PP_ALIGN.CENTER)
    w = L["lab_w"]
    x = px if align == PP_ALIGN.LEFT else (px - w if align == PP_ALIGN.RIGHT else px - w / 2)
    paras = []
    if s.get("metric"):
        paras.append([(s["metric"], {"size": S["section"] * 0.80, "bold": True, "color": ink})])
    paras.append([(s["label"], {"size": S["body"], "bold": True, "color": ink})])
    h = L["row_h"] + (0.40 if s.get("metric") else 0)
    y = py - h / 2
    if align == PP_ALIGN.CENTER:  # 12시·6시 — 선 끝을 넘겨 세운다
        y = py - h - L["lab_gap"] if sa < 0 else py + L["lab_gap"]
    add_text(
        slide,
        x,
        y,
        w,
        h,
        paras,
        S["body"],
        F["head"],
        C["primary"],
        align=align,
        anchor=MSO_ANCHOR.MIDDLE,
        line_spacing=1.08,
    )
    return x, h, align


def _radii(box, data, L):
    """배치 반지름 — 세로는 **12시·6시 라벨 덩이 높이**를 빼고 잡는다.

    라벨이 선 끝 너머로 서므로 그만큼 자리가 필요하다. 안 빼면 6시 라벨이 자리 밖으로
    나간다(2026-07-26 눈검증).
    """
    lab_h = L["row_h"] + (0.40 if any(x.get("metric") for x in data["spokes"]) else 0)
    pad = lab_h + L["lab_gap"] + 0.10
    return (
        min((box.w - L["lab_w"] * 2) / 2, L["r_max_x"]),
        min(box.h / 2 - pad, L["r_max_y"]),
    )


def _draw_line(slide, box, data, kit, L):
    """헤어라인 방사 — 중심 글자에서 각 라벨로 얇은 선 하나."""
    spokes = data["spokes"]
    n = len(spokes)
    cx, cy = box.x + box.w / 2, box.y + box.h / 2
    hw, hh = _hub_text(slide, kit, cx, cy, data, L)
    rx, ry = _radii(box, data, L)
    if rx < hw / 2 + 0.8 or ry < hh / 2 + 0.5:
        raise ValueError(
            f'방사가 이 자리에 안 들어간다 — 반지름 {rx:.2f}×{ry:.2f}"가 '
            f'중심 글자({hw:.2f}×{hh:.2f}")를 못 감싼다. 자리를 넓힐 것.'
        )
    pos = {}
    for i, s in enumerate(spokes):
        ang = math.radians(L["start"] + i * 360 / n)
        ca, sa = math.cos(ang), math.sin(ang)
        # 선은 중심 **글자 상자**의 변에서 시작한다 — 글자를 뚫지 않게
        t0 = min(
            (hw / 2 + L["hub_gap"]) / max(abs(ca), 1e-6),
            (hh / 2 + L["hub_gap"]) / max(abs(sa), 1e-6),
        )
        x1, y1 = cx + t0 * ca, cy + t0 * sa
        x2, y2 = cx + rx * ca, cy + ry * sa
        _hair(slide, kit, x1, y1, x2, y2, accent=bool(s.get("soft")), dashed=bool(s.get("soft")))
        lx = x2 + (L["lab_gap"] if ca > 0.20 else (-L["lab_gap"] if ca < -0.20 else 0))
        _label(slide, kit, lx, y2, ang, s, L)
        pos[s["label"]] = (x2, y2)
    return pos


def _draw_converge(slide, box, data, kit, L):
    """수렴 — 항목이 위에 나란히 서고 화살이 **아래 한 점(허브)으로 모인다**.

    실물: OECD `DAF/COMP(2019)14` p7. 방사(`line`)와 물성은 같은 1:N인데 **방향이 반대**다.
    line은 "무엇이 물려 있나"를 말하고 converge는 "무엇이 여기로 귀속되나"를 말한다.
    그래서 여기서만 화살촉을 쓴다 — 방향이 곧 주장이라서.

    허브는 상자 하나로 앉힌다(실물도 그렇다). 단 **채우지 않는다** — 검정 채움은 반려됐고,
    이 파일 잉크 규율(면을 거의 안 쓴다)에도 어긋난다. 테두리로 끝점을 만들고 무게는 수치가 진다.

    그리기 **순서가 중요하다**: 허브 상자를 화살보다 **먼저** 놓는다. 나중에 놓으면 상자가
    화살촉을 덮어 촉이 사라진다(2026-07-26 눈검증 — "화살표가 네모칸에 겹쳐지는 버그").
    화살은 상자 윗변에서 `cv_head_gap`만큼 **떨어져 멈춘다** — 닿게 하면 촉이 테두리에 물린다.
    """
    from ..grid import track
    from ..kit import add_box
    from .elements import chip, chip_w, link

    C, S, F = kit["rgb"], kit["sizes"], kit["fonts"]
    spokes = data["spokes"]
    n = len(spokes)
    cw = track(n, box.x, box.x + box.w, L["cv_gap"], L["cv_min_w"], what="수렴 항목")
    # 허브 상자 — 고유 크기(글자에서 파생). 자리에 비례해 늘리지 않는다.
    hub_w = max(chip_w(data["hub"], S["head"], 0.70), cw)
    rows = 1 + bool(data.get("hub_metric")) + bool(data.get("hub_note"))
    hub_h = 0.16 + 0.30 * rows + (0.10 if data.get("hub_metric") else 0)
    # 화살 구간에 **상한**이 있다. 자리가 남으면 화살을 늘리는 게 아니라 덩이를 세로 중앙에 둔다
    # — 안 그러면 5.9" 자리에서 화살만 4.5"가 되어 그림이 전부 선이 된다(2026-07-26 눈검증).
    room = box.h - L["cv_chip_h"] - hub_h
    if room < L["cv_drop_min"]:
        raise ValueError(
            f'수렴이 이 자리에 안 들어간다 — 화살 구간 {room:.2f}" '
            f'< 최소 {L["cv_drop_min"]}". 자리 높이를 늘릴 것.'
        )
    drop = min(room, L["cv_drop_max"])
    top_y = box.y + (box.h - (L["cv_chip_h"] + drop + hub_h)) / 2
    hub_y = top_y + L["cv_chip_h"] + drop
    hub_cx = box.x + box.w / 2
    # 허브 먼저 — z-order가 화살촉을 덮지 않게. 채움 없음, 테두리만.
    add_box(
        slide,
        hub_cx - hub_w / 2,
        hub_y,
        hub_w,
        hub_h,
        fill=C["bg"],
        line=C["primary"],
        line_w=1.25,
        shape="round",
    )
    paras = [[(data["hub"], {"size": S["head"], "bold": True, "color": C["primary"]})]]
    if data.get("hub_metric"):
        paras.insert(
            0,
            [
                (
                    data["hub_metric"],
                    {"size": S["section"] * 0.72, "bold": True, "color": C["primary"]},
                )
            ],
        )
    if data.get("hub_note"):
        paras.append([(data["hub_note"], {"size": S["caption"], "color": C["muted"]})])
    add_text(
        slide,
        hub_cx - hub_w / 2,
        hub_y + 0.05,
        hub_w,
        hub_h - 0.10,
        paras,
        S["head"],
        F["head"],
        C["primary"],
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
        line_spacing=1.05,
    )
    pos = {data["hub"]: (hub_cx, hub_y + hub_h / 2)}
    for i, s in enumerate(spokes):
        x = box.x + i * (cw + L["cv_gap"])
        chip(
            slide,
            x,
            top_y,
            cw,
            L["cv_chip_h"],
            s["label"],
            kit,
            style="soft" if s.get("soft") else "outline",
            pt=S["body"],
        )
        link(
            slide,
            kit,
            x + cw / 2,
            top_y + L["cv_chip_h"],
            hub_cx + (x + cw / 2 - hub_cx) * 0.10,
            hub_y - L["cv_head_gap"],
            dashed=bool(s.get("soft")),
            width=0.9,
            color="accent" if s.get("soft") else "muted",
        )
        pos[s["label"]] = (x + cw / 2, top_y + L["cv_chip_h"])
    return pos


def _draw_scatter(slide, box, data, kit, L):
    """불균등 방사 — 중심에서 **거리가 제각각인** 실선 화살이 뻗고 끝마다 작은 테두리 노드.

    실물: Housing LIN 「Models of Hub and Spoke」 p4. `line`과 갈리는 지점이 둘이다.
    ① 반지름이 고르지 않다 — 균등 방사는 "분류표"로 읽히고 불균등은 "도달 범위"로 읽힌다.
       그래서 `sc_var`를 0으로 두면 이 모드는 존재 이유가 없다.
    ② 끝이 맨 글자가 아니라 테두리 노드다 — 도달한 **대상**이 있다는 표시.

    불균등은 난수가 아니라 **순번에서 파생**한다(계약 ②). 같은 데이터는 같은 그림이어야
    스냅샷 대조가 가능하다.
    """
    from .elements import chip, chip_w, link

    S = kit["sizes"]
    spokes = data["spokes"]
    n = len(spokes)
    cx, cy = box.x + box.w / 2, box.y + box.h / 2
    hw, hh = _hub_text(slide, kit, cx, cy, data, L, metric_key=L["sc_hub_pt"])
    node_ws = [chip_w(s["label"], S["caption"], L["sc_inset"], floor=0.70) for s in spokes]
    # 반지름 = 자리에서 파생하되 끝 노드가 자리 밖으로 안 나가게 노드 크기를 빼고, 상한은
    # 자리가 아니라 **가독성**(sc_r_max_*)에서 온다 — 상자 노드는 맨 글자보다 넓게 퍼진다.
    rx = min((box.w - max(node_ws) * 2) / 2, L["sc_r_max_x"])
    ry = min((box.h - L["sc_node_h"] * 2) / 2, L["sc_r_max_y"])
    if rx < hw / 2 + 0.8 or ry < hh / 2 + 0.5:
        raise ValueError(
            f'방사가 이 자리에 안 들어간다 — 반지름 {rx:.2f}×{ry:.2f}"가 '
            f'중심 글자({hw:.2f}×{hh:.2f}")를 못 감싼다. 자리를 넓힐 것.'
        )
    pos = {}
    for i, s in enumerate(spokes):
        ang = math.radians(L["start"] + i * 360 / n)
        ca, sa = math.cos(ang), math.sin(ang)
        # 불균등 — 순번의 주기적 함수. 3과 n이 서로소가 아니어도 진폭이 죽지 않게 5로 접는다.
        k = ((i * 3) % 5) / 4.0 - 0.5  # -0.5 ~ +0.5
        f = 1.0 - L["sc_var"] * (0.5 - k)
        t0 = min(
            (hw / 2 + L["hub_gap"]) / max(abs(ca), 1e-6),
            (hh / 2 + L["hub_gap"]) / max(abs(sa), 1e-6),
        )
        x1, y1 = cx + t0 * ca, cy + t0 * sa
        x2, y2 = cx + rx * f * ca, cy + ry * f * sa
        link(
            slide,
            kit,
            x1,
            y1,
            x2,
            y2,
            dashed=bool(s.get("soft")),
            width=0.75,
            color="accent" if s.get("soft") else "muted",
        )
        nw = node_ws[i]
        # 노드를 선 끝 **너머로** 밀어낸다. 안 밀면 선이 노드 안까지 들어가 화살촉이 상자에
        # 가려진다(2026-07-26 눈검증: 12시 화살에 촉이 안 보였다).
        side = 1 if ca > 0.20 else (-1 if ca < -0.20 else 0)
        nx = x2 + side * (nw / 2 + 0.03)
        ny = y2 if side else y2 + (-1 if sa < 0 else 1) * (L["sc_node_h"] / 2 + 0.03)
        chip(
            slide,
            nx - nw / 2,
            ny - L["sc_node_h"] / 2,
            nw,
            L["sc_node_h"],
            s["label"],
            kit,
            style="soft" if s.get("soft") else "outline",
            pt=S["caption"],
        )
        pos[s["label"]] = (nx, ny)
    return pos


MODES = {
    "line": _draw_line,
    "converge": _draw_converge,
    "scatter": _draw_scatter,
}


def measure(data, kit, layout=None):
    """(필요 폭, 필요 높이) — 모드마다 다르다. 3층이 자리를 손으로 재지 않게 하는 값이다.

    scatter가 line보다 **작다**는 것이 요점이다. 이 값을 하나로 묶어두면 3층이 scatter에도
    12" 폭을 주고, 부품은 그 자리를 채우려 노드를 흩뿌린다(2026-07-26 눈검증 실패의 경로).
    """
    from .elements import chip_w

    L = merged(LAYOUT, layout)
    S = kit["sizes"]
    n = len(data["spokes"])
    if L["mode"] == "converge":
        w = n * L["cv_min_w"] + (n - 1) * L["cv_gap"]
        rows = 1 + bool(data.get("hub_metric")) + bool(data.get("hub_note"))
        return w, L["cv_chip_h"] + L["cv_drop_min"] + 0.16 + 0.30 * rows
    if L["mode"] == "scatter":
        node_w = max(
            chip_w(s["label"], S["caption"], L["sc_inset"], floor=0.70) for s in data["spokes"]
        )
        return L["sc_r_max_x"] * 2 + node_w * 2, L["sc_r_max_y"] * 2 + L["sc_node_h"] * 2
    return L["r_max_x"] * 2 + L["lab_w"] * 2, L["r_max_y"] * 2 + L["row_h"]


def draw(slide, box, data, kit, layout=None):
    """box 안에 방사 그래픽을 그린다. 반환: {라벨: (x, y)} — 오딧용.

    data = {"hub": str, "hub_note": str, "hub_metric": str,
            "spokes": [{"label": str, "metric": str, "soft": bool}, ...]}
    """
    require(data, META["keys"], META["name"])
    L = merged(LAYOUT, layout)
    box = Box(*box)
    n = len(data["spokes"])
    if n > L["max_n"]:
        raise ValueError(
            f"곁가지 {n}개는 방사로 안 읽힌다(최대 {L['max_n']}) — 각도 간격이 좁아진다."
        )
    if L["mode"] not in MODES:
        raise KeyError(f"알 수 없는 mode: {L['mode']} — {list(MODES)} 중 하나")
    return MODES[L["mode"]](slide, box, data, kit, L)


_D5 = {
    "hub": "온톨로지 그래프",
    "hub_metric": "423",
    "hub_note": "표제어 · 단일 진실 원천",
    "spokes": [
        {"label": "용어 정규화"},
        {"label": "조문 인용"},
        {"label": "사례 매핑"},
        {"label": "판정 로그"},
        {"label": "확장 API", "soft": True},
    ],
}
_D4 = {
    "hub": "판정 엔진",
    "hub_note": "규칙 + 그래프",
    "spokes": [
        {"label": "계약 식별"},
        {"label": "수행의무"},
        {"label": "거래가격"},
        {"label": "인식 시점"},
    ],
}
_DW = {
    "hub": "질의",
    "hub_metric": "288건",
    "hub_note": "2026-07 실측",
    "spokes": [
        {"label": "자동 판정", "metric": "196"},
        {"label": "위임 판단", "metric": "62"},
        {"label": "사람 검토", "metric": "18"},
        {"label": "판정 불가", "metric": "12", "soft": True},
    ],
}

SAMPLES = [
    ("A · 헤어라인 방사 — 수치가 제일 큰 잉크", _D5, {"mode": "line"}),
    ("A · 항목마다 수치", _DW, {"mode": "line"}),
    ("A · 4개 · 수치 없음", _D4, {"mode": "line"}),
    ("B · 수렴 — 위 나란히 → 아래 한 점 (OECD 실물)", _D5, {"mode": "converge"}),
    ("B · 4개 · 수치 있는 허브", _DW, {"mode": "converge"}),
    ("C · 불균등 방사 — 거리가 제각각 (Housing LIN 실물)", _D5, {"mode": "scatter"}),
    ("C · 진폭 0 = 균등(이 모드를 쓸 이유가 없는 상태)", _D5, {"mode": "scatter", "sc_var": 0.0}),
    ("C · 4개", _D4, {"mode": "scatter"}),
]
