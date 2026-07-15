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
        ("P4③ 노드 재탕", *A.check_duplicate_nodes(shapes)),
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
