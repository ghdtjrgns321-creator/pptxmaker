"""content 필수 키 계약 — 골든 기본값이 공장 문으로 새는 것을 막는 단일 출처.

왜 필요한가(2026-07-15 실증): 계약(01.5_outline.md)에 8장 전부 "README §번호"가 확정으로
적혔는데 반영 0장이었다. 그런데 전 게이트가 PASS를 냈다. 원인은 침묵 폴백 —
`c = {**DEFAULT, **(c or {})}`는 content가 없으면 골든 글로 조용히 채운다. 에러도 경고도 없고,
**빈 장이 아니라 그럴듯하게 채워진 장**이 나오므로 아무도 못 본다.

증거: cover·toc·part는 `c["kicker"]`를 직접 접근해 content 없으면 KeyError로 죽는다.
그 3종은 항상 제대로 주입됐다. 조용한 12종은 전부 골든 글이 나갔다. 시끄러운 게 옳다.

역할 분담:
    골든 기본값 = 회귀 게이트(compare_golden) 전용 기준점. 골든이 안 변했는지 비교하는 잣대.
    공장 문(build_pptx._render_golden) = 계약 content만 통과. 골든 글은 한 글자도 못 나간다.

compare_golden은 레이아웃 함수를 직접 호출(`FAB[key](prs, None)`)하므로 이 계약을 안 거친다
— 우회 플래그 불필요. 공장 문에서만 강제된다.

**두 번째 문(2026-07-25 신설): `adapted.`/`novel` 장 스크립트.**
위 문은 `golden.<layout>` 경로(`_render_golden`)에만 있었다. `_render_scripted`는 `assert_content`를
부르지 않았으므로 **실전 덱을 `adapted.`로 만들면 골든(K-IFRS) 글이 그대로 남아도 전 게이트가
green**이었다 — 이 파일이 막으려던 사고와 정확히 같은 클래스가 옆문으로 열려 있었다.
그 문은 `assert_scripted_content`가 진다. 임계는 **"그 스크립트가 실제로 읽는 키"**다(아래 참조).
"""

import re
import sys

# DEFAULT dict 없이 c[...]를 직접 접근하는 레이아웃 — 명시 선언(실측: layouts.py 소스)
#
# part의 `total`은 코드상 `c.get("total", 6)`로 **선택**이지만 여기서는 **필수**로 둔다:
# 안 주면 5부 덱에 점이 6개 찍힌다 — 에러 없이 조용히 틀리는 값이라 이 계약이 막으려는
# 바로 그 클래스다. 현 deck-spec은 이미 4키 전부 주고 있어 과엄격으로 깨지는 곳은 없다.
EXPLICIT_KEYS = {
    "cover": {"brand_name", "kicker", "title", "value_prop", "year"},
    "toc": {"deck_sub", "deck_title", "items", "title"},
    "part": {"lead", "no", "title", "total"},
}


def golden_defaults(fn):
    """레이아웃 함수 → 그 모듈의 골든 기본값 dict. 없으면 None.

    모듈마다 이름이 다르다(DEFAULT · DEFAULT_C · SHOT_DEFAULTS · VARIANT_C_DEFAULTS …).
    이름 대신 "모듈 최상위에서 이름에 DEFAULT를 포함한 dict가 정확히 1개"로 찾는다.
    """
    mod = sys.modules[fn.__module__]
    found = [v for n, v in vars(mod).items() if "DEFAULT" in n.upper() and isinstance(v, dict)]
    if len(found) != 1:
        return None
    return found[0]


def required_keys(name, fn):
    """레이아웃이 요구하는 content 키 집합. 계약이 이 키를 다 채워야 골든 글이 안 샌다."""
    if name in EXPLICIT_KEYS:
        return set(EXPLICIT_KEYS[name])
    d = golden_defaults(fn)
    if d is None:
        raise ValueError(
            f"golden.{name}: 골든 기본값 dict를 특정할 수 없다(모듈 {fn.__module__}에 "
            "이름에 DEFAULT를 포함한 최상위 dict가 0개 또는 2개 이상). "
            "content_contract.EXPLICIT_KEYS에 필수 키를 명시하거나 모듈의 DEFAULT를 하나로 정리할 것."
        )
    return set(d)


def content_keys_read(src):
    """소스가 실제로 읽는 content 키 — `c["k"]` · `c.get("k")`. 계약 선언의 진위 판정 기준.

    단일 출처: `test_wiring` 통합E·통합F도 이 함수를 쓴다(각자 정규식을 두면 갈라진다).
    """
    return set(re.findall(r"""c\[["']([a-z0-9_]+)["']\]""", src)) | set(
        re.findall(r"""c\.get\(["']([a-z0-9_]+)["']""", src)
    )


def scripted_required_keys(module, src):
    """장 스크립트가 **주입을 요구하는** 키 = (모듈 DEFAULT 키) ∩ (실제로 읽는 키).

    왜 교집합인가(오탐 0 설계, 2026-07-25 실측): dense 기준작은 DEFAULT에 키를 두고도 본문은
    **모듈 상수를 그리는** 구멍이 28건 있다(`test_wiring` 통합F 래칫). 그 키를 필수로 요구하면
    골든 출발 스크립트가 즉시 28건 적색이 되고, 상시 적색 게이트는 무시당한다 — 이 프로젝트가
    이미 겪은 경로다. 그래서 **주입해도 안 읽히는 키는 요구하지 않는다**. 대신 구멍 자체를
    통합F가 감소 전용 래칫으로 고정하므로 "안 읽히는 키"가 늘어나 검사를 우회할 수는 없다.

    DEFAULT dict이 아예 없는 스크립트(실전용으로 새로 쓴 adapted·novel)는 필수 키 0 —
    골든 글이 들어있지 않으므로 유출 위험 자체가 없다. 억지로 요구하면 novel이 막힌다.
    """
    defaults = set()
    for n, v in vars(module).items():
        if "DEFAULT" in n.upper() and isinstance(v, dict):
            defaults |= set(v)
    return defaults & content_keys_read(src)


def assert_scripted_content(name, module, src, content, slide_no):
    """`adapted.<layout>`/`novel` 장 스크립트의 content 계약. 누락 시 ValueError(빌드 중단)."""
    req = scripted_required_keys(module, src)
    if not req:
        return True
    content = content or {}
    given = {k for k in content if not _is_empty(content[k])}
    missing = req - given
    if not missing:
        return True
    lines = [
        f"슬라이드 {slide_no} ({name}): 장 스크립트가 읽는 content 키 {len(missing)}/{len(req)}개 누락 "
        f"→ 그 자리에 스크립트의 골든 기본값(다른 프로젝트 글)이 그대로 나간다.",
        f"  누락: {sorted(missing)}",
        "  필수 키 = 스크립트 DEFAULT 키 ∩ 스크립트가 실제 읽는 키 "
        "(안 읽히는 키는 요구하지 않는다 — goldenfab/content_contract.scripted_required_keys).",
    ]
    empty = {k for k in req & set(content) if _is_empty(content[k])}
    if empty:
        lines.append(
            f"  이 중 {sorted(empty)}는 키만 있고 **값이 비었다** — 키 존재는 주입이 아니다."
        )
    raise ValueError("\n".join(lines))


def _is_empty(v):
    """빈 값 = 사실상 미주입. 키만 있고 값이 None/""/[]이면 폴백이 돌거나 빈 자리가 난다.

    2026-07-15 제3자 반증: `"points": None`을 넘기면 키가 있으니 통과했고, 함수 내부
    `c["points"] or [골든 리터럴]`이 골든 원문을 주입했다. 키 존재만 보면 못 막는다.
    """
    if v is None:
        return True
    if isinstance(v, (str, list, tuple, dict)) and len(v) == 0:
        return True
    return False


def assert_content(name, fn, content, slide_no):
    """계약 content가 골든 기본값을 전부 덮는지 검사. 누락 시 ValueError(빌드 중단).

    부분 주입도, 빈 값 주입도 실패시킨다 — 누락 키가 골든 글로 조용히 메워지는 게 사고의 본질이다.
    """
    req = required_keys(name, fn)
    content = content or {}
    given = {k for k in content if not _is_empty(content[k])}
    missing = req - given
    if not missing:
        return True

    empty = {k for k in req & set(content) if _is_empty(content[k])}
    extra = set(content) - req
    lines = [
        f"슬라이드 {slide_no} (golden.{name}): content 누락 {len(missing)}/{len(req)}키 "
        f"→ 이 키들은 골든 기본값(다른 프로젝트 글)으로 채워진다.",
        f"  누락: {sorted(missing)}",
    ]
    if empty:
        lines.append(
            f"  이 중 {sorted(empty)}는 키는 있으나 **값이 비었다**(None/빈 문자열/빈 리스트) "
            "— 키 존재는 주입이 아니다. 실제 내용을 넣을 것."
        )
    if not given:
        lines.append(
            "  content가 통째로 없다. 골든은 서식 견본이지 콘텐츠 원천이 아니다 — "
            "계약(01.5_outline.md)이 지정한 출처에서 본문을 채울 것."
        )
    if extra:
        lines.append(f"  계약에 없는 키(오타 의심): {sorted(extra)}")
    lines.append(
        "  키 목록: uv run python _workspace/dump_contract.py → golden-content-contract.md"
    )
    raise ValueError("\n".join(lines))
