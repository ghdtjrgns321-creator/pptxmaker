"""골든 레이아웃 기계 오딧 러너 — 2초. 제3자 채점 **전에** 이걸 통과시킨다.

사용:
    uv run python .claude/skills/pptx-build/scripts/audit_golden.py            # 검사
    uv run python .claude/skills/pptx-build/scripts/audit_golden.py --selftest # 규칙 민감도 검증

규칙 본체는 `goldenfab/audit.py`. 여기는 **레이아웃별 기대값**만 선언한다.
새 레이아웃을 오딧에 넣으려면 SPECS에 한 항목을 추가한다.

왜 이 파일이 있나: S4 한 장에 제3자 채점 2회 = 30분이 들었는데, 두 채점자의 FAIL 9건 중
**7건이 순수 산수**였다. 채점자는 규칙을 **발견**할 뿐 적용은 코드가 한다 — 한 번 캔 규칙을
여기 내려두면 다음 장부터 2초에 잡힌다. 상세: goldenfab/audit.py 모듈 독스트링.
"""

import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from goldenfab import audit as A  # noqa: E402
from goldenfab.kit import load_kit  # noqa: E402
from goldenfab.registry import LAYOUTS  # noqa: E402

K = load_kit()
C = K["rgb"]


def render(name):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    LAYOUTS[name](prs, None)
    return list(prs.slides)[-1]


def air_pairs_problem_grid():
    import goldenfab._variant_k as VK

    return [
        ("밴드1 룰 → 첫 행", VK.B1_HEAD_Y + VK.BAND_RULE_DY + 0.012, VK.ROW_A_Y),
        ("2종 캡션 → 1종 행", VK.ROW_B_Y + VK.ROW_BC_H + VK.CAP_DY + VK.CAP_H, VK.ROW_C_Y),
        ("1종 행 → 복귀 호 라벨", VK.ROW_C_Y + VK.ROW_BC_H, VK.BACK_LABEL_Y),
        ("복귀 호 → 밴드2 라벨", VK.BACK_Y, VK.B2_HEAD_Y),
        ("밴드2 룰 → 한계 첫 칩", VK.B2_HEAD_Y + VK.BAND_RULE_DY + 0.012, VK.LIM_TOP),
        (
            "한계 마지막 칩 → 하한",
            VK.LIM_TOP + VK.LIM_PITCH * 3 + VK.LIM_H,
            A.CONTENT_BOTTOM + A.AIR_MIN,
        ),
    ]


def air_pairs_tech_evidence():
    """S8 이분 매핑 — 소제목 룰 → 첫 칩, 마지막 칩 → 캡션. 좌표는 모듈 상수에서 파생."""
    import goldenfab.s08_variants as S8

    n_concepts = len({c for _t, _g, cs in S8.DEFAULT_C["terms"] for c in cs})
    pitch = S8._pitch(n_concepts, S8.MAP_TOP, S8.MAP_BOTTOM, S8.MAP_H)
    last_bottom = S8.MAP_TOP + (n_concepts - 1) * pitch + S8.MAP_H
    return [
        ("매핑 룰 → 첫 칩", S8.MAP_HEAD_Y + S8.MAP_RULE_DY + 0.012, S8.MAP_TOP),
        ("마지막 칩 → 캡션", last_bottom, 6.09),
    ]


def air_pairs_tech_tree():
    """S9 카탈로그 — 소제목 → 첫 행, 마지막 행 → 하한."""
    import goldenfab.s09_variants as S9

    rows = S9.DEFAULT["cat_rows"]
    pitch = S9._cat_pitch(len(rows))
    last_bottom = S9.CAT_TOP + (len(rows) - 1) * pitch + S9.CAT_H
    return [
        ("카탈로그 소제목 → 첫 행", S9.CAT_HEAD_Y + 0.24, S9.CAT_TOP),
        ("마지막 행 → 하한", last_bottom, A.CONTENT_BOTTOM + A.AIR_MIN),
    ]


# 레이아웃별 기대값. progress=허용 셰브런 수, ink_allow=primary 채움 허용 수(결론 바 1개).
SPECS = {
    "problem_grid": {
        "progress": 0,  # XOR을 셰브런으로 그리면 거짓 인과 (반려 4회)
        "ink_allow": 1,  # primary 채움 = 결론 바 하나뿐
        "node_top_max": 4.6,  # 인과도 구획만 행 노드로 취급
        "air": air_pairs_problem_grid,
    },
    # 그림이 있는 장 — 침범 감시가 목적(2026-07-15 s10 버그: 비율 가정이 틀려 텍스트 1.74" 침범)
    #
    # `known`: 이 오딧을 만들기 **전에** 사용자 승인으로 확정된 장의 기존 위반. 규칙을 끄지 않고
    # 기준선으로 박는다 — **더 나빠지면 잡히고, 지금 상태는 통과**(래칫). 규칙을 빼면 hollow가 되고,
    # 그대로 빨간불이면 게이트가 상시 적색이라 무시당한다(둘 다 이 프로젝트가 이미 겪은 실패다).
    # 해소는 해당 장 재설계 시 — 재설계는 사용자 결정이라 여기서 임의로 하지 않는다.
    "screenshot": {
        "progress": 0,
        "ink_allow": 0,
        "node_top_max": 0,
        "air": lambda: [],
        "known": {
            # 관찰 번호 01~04가 accent 런 4 + 캡션 바 1. "사용처"로는 2(번호 장치·바)지만
            # 계수기는 런 단위라 5로 센다 — 계수 입도 문제이지 설계 결함이 아니다.
            "§2 accent 상한": 5,
            # ("P2③④ 경계": 1 — 2026-07-15 **해소**. 리터럴 피치 1.05를 항목 수에서 파생시켜
            #  마지막 본문 bottom 6.47 → 6.35. 3·4·5·6개 전부 하한 내. 기준선 삭제 = 빚 청산.)
        },
    },
    "tech_capture": {
        "progress": 0,
        "ink_allow": 0,
        "node_top_max": 0,
        "air": lambda: [],
        "known": {"§2 accent 상한": 6, "P4⑩ 검정 충돌": 2},
    },
    # 기술 설명 장 2종 (2026-07-15 표 → 관계 시각화 재설계). **known 없음** — 새로 그렸으므로
    # 빚 없이 시작한다. 여기서 기준선을 깔면 재설계의 의미가 없다.
    # ink_allow=2: 결론 바 + 승인된 다크 면 1개. S4의 "primary 채움 = 결론 바 하나뿐"은 그 장의
    # 규칙이다 — 거기선 검정이 '파국'과 '우리 원칙' 두 뜻을 져서 충돌했다. 여기선 두 번째 검정이
    # §E가 규정한 **다크 JSON 카드**(S8)와 **트리 루트 `기준서 1115`**(S9)로, 결론 바와 뜻이
    # 겹치지 않는다(코드 실물 / 기준서 최상위). 사용자도 코드 카드를 명시 승인했다("코드를
    # 보여주는 시각화는 좋은데").
    "tech_evidence": {  # S8 이분 매핑 — 실무어 → 기준서 개념
        "progress": 0,
        "ink_allow": 2,  # 결론 바 + JSON 카드(§E 다크 라운드 rect)
        "node_top_max": 0,
        "air": air_pairs_tech_evidence,
    },
    "tech_tree": {  # S9 간선 7종 카탈로그 + 위계 트리
        "progress": 0,
        "ink_allow": 2,  # 결론 바 + 트리 루트(기준서 1115)
        "node_top_max": 0,
        # 같은 노드의 재등장이 메시지다 — `변동대가` ×4(트리 개념층 + 카탈로그 hier·cp 출발 +
        # term 도착), `문단 50` ×3(트리 + cp·case 도착). 좌 트리와 우 카탈로그가 **같은 실물을
        # 두 각도로** 본다는 게 이 장의 구성이라, 재등장이 곧 두 구획이 맞물린다는 증거다.
        #
        # 5 → 8은 **악화가 아니라 드러남**이다. 이전 5는 좌 트리(bg_alt)와 우 카탈로그(흰+검정
        # 테두리)의 서명이 **갈려 있어서** 같은 이름이 다른 노드로 세어졌던 것이고, 그 갈림이
        # 바로 제3자가 잡은 클래스 충돌(#10)이었다. NODE_STYLE로 통일하니 진짜 수가 보인다 —
        # 규칙을 느슨하게 한 게 아니라 **측정이 정확해졌다**. 실측 8을 박아 늘면 FAIL.
        "dup_allow": 8,
        "air": air_pairs_tech_tree,
    },
}


def audit(name):
    spec = SPECS[name]
    sl = render(name)
    shapes = list(sl.shapes)
    print(f"── {name}: 도형 {len(shapes)}")
    res = [
        ("§5 진행형 도형", *A.check_progress_shapes(shapes, spec["progress"])),
        ("§2 accent 상한", *A.check_accent(shapes, str(C["accent"]))),
        ("P4⑤ 판정 단어 대비", *A.check_verdict_contrast(shapes, None)),
        ("P4④ 채움률", *A.check_fill_ratio(shapes)),
        ("P4③ 노드 재탕", *A.check_duplicate_nodes(shapes, allow=spec.get("dup_allow", 0))),
        ("P4⑩ 노드 클래스", *A.check_node_class(shapes)),
        ("P4⑩ 검정 충돌", *A.check_ink_collision(shapes, str(C["primary"]), spec["ink_allow"])),
        ("P4⑧ 노드 높이", *A.check_node_heights(shapes, spec["node_top_max"])),
        ("§F 그림 침범", *A.check_picture_overlap(shapes)),
        ("P2③④ 경계", *A.check_bounds(shapes, skip_tops=(6.60, 7.15))),
        ("§6 공기", *A.check_air(spec["air"]())),
    ]
    return A.report(res, known=spec.get("known"))


def selftest():
    """규칙 민감도 — 일부러 어긴 슬라이드를 잡는가. 안 잡으면 죽은 코드다."""
    from pptx.enum.shapes import MSO_SHAPE

    ok = []
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # 셰브런 1개 심기 → §5가 잡아야
    sl.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(1), Inches(1), Inches(1), Inches(0.5))
    hit, msg, n = A.check_progress_shapes(list(sl.shapes), 0)
    ok.append(("§5 셰브런 검출", not hit, msg, n))

    # 흐린 판정 단어 심기 → P4⑤가 잡아야
    tb = sl.shapes.add_textbox(Inches(1), Inches(2), Inches(2), Inches(0.3))
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = "✕ 불가역"
    r.font.color.rgb = C["muted"]  # 3.2:1
    hit, msg, n = A.check_verdict_contrast(list(sl.shapes), None)
    ok.append(("P4⑤ 흐린 판정 검출", not hit, msg, n))

    # 산문은 muted가 정상 — 판정으로 오탐하면 안 된다(2026-07-15 규칙 정밀화의 회귀 방지).
    # 이 케이스가 없으면 "없다"를 뺀 것이 규칙을 죽인 건지 고친 건지 구분할 수 없다.
    prs_p = Presentation()
    prs_p.slide_width, prs_p.slide_height = Inches(13.333), Inches(7.5)
    slp = prs_p.slides.add_slide(prs_p.slide_layouts[6])
    tbp = slp.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(0.3))
    rp = tbp.text_frame.paragraphs[0].add_run()
    rp.text = "문단 뭉치에는 순서도 관계도 없다"  # 17자 서사 본문 — 판정 아님
    rp.font.color.rgb = C["muted"]
    hit_p, msg_p, n_p = A.check_verdict_contrast(list(slp.shapes), None)
    ok.append(("P4⑤ 산문 오탐 0", hit_p, msg_p, n_p))

    # 그림 침범 — 구 s10 방식(height만 지정 → 비율이 폭을 정함)을 재현하면 잡아야
    import goldenfab.s10_screenshot as S10

    prs2 = Presentation()
    prs2.slide_width, prs2.slide_height = Inches(13.333), Inches(7.5)
    sl2 = prs2.slides.add_slide(prs2.slide_layouts[6])
    sl2.shapes.add_picture(S10.SHOT_DEFAULTS["img"], Inches(0.6), Inches(1.8), height=Inches(4.75))
    tb2 = sl2.shapes.add_textbox(Inches(8.6), Inches(1.8), Inches(4.1), Inches(0.3))
    tb2.text_frame.paragraphs[0].add_run().text = "이 화면에서 보이는 것"
    hit, msg, n = A.check_picture_overlap(list(sl2.shapes))
    ok.append(("§F 그림 침범 검출", not hit, msg, n))

    # 정상(fit_picture)은 통과해야 — 오탐 0
    hit3, msg3, n3 = A.check_picture_overlap(list(render("screenshot").shapes))
    ok.append(("§F 정상 오탐 0", hit3, msg3, n3))

    # 노드 클래스 — 같은 이름을 다른 스타일로 그리면 잡아야(제3자가 s9에서 찾은 최대 결함)
    from pptx.enum.shapes import MSO_SHAPE as _MS

    prs_n = Presentation()
    prs_n.slide_width, prs_n.slide_height = Inches(13.333), Inches(7.5)
    sln = prs_n.slides.add_slide(prs_n.slide_layouts[6])
    for i, (fill, line) in enumerate(((C["bg_alt"], None), (C["bg"], C["primary"]))):
        sh = sln.shapes.add_shape(
            _MS.ROUNDED_RECTANGLE, Inches(1 + i * 2), Inches(1), Inches(1.5), Inches(0.3)
        )
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = line
        sh.text_frame.paragraphs[0].add_run().text = "변동대가"  # 같은 이름, 다른 스타일
    hit_n, msg_n, n_n = A.check_node_class(list(sln.shapes))
    ok.append(("P4⑩ 클래스 충돌 검출", not hit_n, msg_n, n_n))

    # 정상(단일 출처 NODE_STYLE)은 통과해야 — 오탐 0
    hit_t, msg_t, n_t = A.check_node_class(list(render("tech_tree").shapes))
    ok.append(("P4⑩ 정상 오탐 0", hit_t, msg_t, n_t))

    # S4의 두 ◇ `근거 충분?`(실선 vs 점선 = 의도된 대조)을 오탐하면 안 된다 — 초판이 테두리를
    # 서명에 넣어 이걸 잡았고, 그럼 duplicate 규칙("구별 장치를 둬라")과 서로 반대를 요구한다.
    hit_k, msg_k, n_k = A.check_node_class(list(render("problem_grid").shapes))
    ok.append(("P4⑩ 의도된 대조 오탐 0", hit_k, msg_k, n_k))

    # 밀도 밴드 — 희소 슬라이드(도형 4개)는 잡고, 골든 본문 장은 전수 통과해야 (2026-07-16
    # 변형 출발점 배선: 스냅샷 회귀가 물러난 자리의 밀집화 게이트. 임계는 스냅샷 파생 — 리터럴 0)
    band = A.density_band()
    prs_d = Presentation()
    prs_d.slide_width, prs_d.slide_height = Inches(13.333), Inches(7.5)
    sld = prs_d.slides.add_slide(prs_d.slide_layouts[6])
    for i in range(4):
        tb = sld.shapes.add_textbox(Inches(1), Inches(1 + i * 0.5), Inches(3), Inches(0.3))
        tb.text_frame.paragraphs[0].add_run().text = f"성긴 항목 {i + 1}"
    hit_d, msg_d, n_d = A.check_density(list(sld.shapes), band)
    ok.append(("§6-D 희소 검출", not hit_d, msg_d, n_d))

    body_keys = [k for k in SPECS]  # 오딧 등록 본문 장 = 골든 본문 표본
    dens_fails = []
    for k in body_keys:
        okk, _m, _n = A.check_density(list(render(k).shapes), band)
        if not okk:
            dens_fails.append(k)
    ok.append(
        (
            "§6-D 골든 오탐 0",
            not dens_fails,
            f"골든 본문 {len(body_keys) - len(dens_fails)}/{len(body_keys)} 통과 {dens_fails}",
            len(dens_fails),
        )
    )

    # 대비 산수 자체
    cr_muted = A.contrast(C["muted"])
    cr_primary = A.contrast(C["primary"])
    ok.append(
        (
            "대비 계산",
            cr_muted < 4.5 < cr_primary,
            f"muted {cr_muted} < 4.5 < primary {cr_primary}",
            0,
        )
    )

    print("── 민감도 자체검증 (규칙이 위반을 실제로 잡는가)")
    return A.report(ok)


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(1 if selftest() else 0)
    fails = []
    for nm in SPECS:
        fails += audit(nm)
    sys.exit(1 if fails else 0)
