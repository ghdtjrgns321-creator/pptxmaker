# pptmaker 지도

> **자동 생성 — 손으로 고치지 않는다.** `uv run python scripts/make_map.py`
> 모든 줄이 코드·SKILL.md에서 파생된다. 어긋나면 생성기가 exit 1로 죽는다.

## 1. 덱 하나가 만들어지는 순서

원문 단일 출처: `.claude/skills/pptmaker/SKILL.md`

```
① 재료 읽기    사용자가 지정한 재료(README 또는 FINAL-REPORT) + _workspace/01.5_outline.md
               계약이 없으면 deck-outline-grill로 먼저 확정 — 없이 진행 금지
               계약 확정 직후 `deck_state.py init <프로젝트>` — 이후 단계마다 STATE.md 갱신
② 형태 후보    **내용은 ①에서 이미 정해졌다.** 계약의 채록을 초점으로 삼고, 그 초점을
               담는 **배치·인코딩 4~6가지**를 낸다 — 고르는 건 "무엇을"이 아니라 "어떻게"다
③ 갤러리       후보를 **실제 pptx로 그려** 한 벌 만든다. 한 장에 한 주제, 후보를 나란히.
               장 이름·그 장의 초점(채록)·후보별 형태 설명을 같이 적는다(설명 없으면 못 고른다)
               **후보는 스케치다** — 구획 비율만 실제로, 라벨은 대표 2~3개 + "… N개"로 접는다
               갤러리 스크립트에 `CORE = {"S4": [핵심 실체 낱말...]}`를 선언한다
               `deck_state.py gallery <프로젝트>` — 본문 장 수 미달, 또는 후보가 핵심 실체를
               절반 미만 담으면 exit 1 (2026-08-03: 후보 ④에서 5축을 통째로 빼 반려당함)
④ 선택         사용자가 번호로 회신 — "s04=1 · s07=2 · s09=전탈락"
               전탈락이면 그 장은 ②로 되돌아간다
⑤ 본편         고른 형태로 장을 만든다. 파이썬 직접 조판(deckkit — 토큰·골격·블록 8종·커넥터)
⑥ 눈검증       render_deck.ps1 → PNG 전장 확인 → 겹침·잘림·넘침 수정. **이게 진짜 게이트다**
               수정 라운드는 규율이 다르다 — 아래 §수정 라운드
⑦ 기계 채점    score_deck.py — 튀는 곳만 본다(판정자 아님)
               끝내기 전 `deck_state.py check <프로젝트>` — 반려가 `하네스: 미정`이면 exit 1
               (미정으로 덮으면 다음 덱에서 같은 반려가 다시 나온다 — 2026-08-02 실측)
```

절차가 이름을 부르는 것들이 실재하는지:

| 이름 | 실재 | 위치 |
| --- | --- | --- |
| `deck-outline-grill` | O | .claude/skills/deck-outline-grill |
| `deck_state.py` | O | scripts/deck_state.py |
| `render_deck.ps1` | O | .claude/skills/pptx-build/scripts/render_deck.ps1 |
| `score_deck.py` | O | .claude/skills/pptx-build/scripts/score_deck.py |

## 2. 살아있는 코드

| 파일 | 줄 | 공개 함수 | 주요 함수 |
| --- | ---: | ---: | --- |
| `.claude/skills/pptx-build/scripts/check_deck.py` | 252 | 6 | coverage, headlines, formal, nav, check, main |
| `.claude/skills/pptx-build/scripts/deckkit/__init__.py` | 5 | 0 |  |
| `.claude/skills/pptx-build/scripts/deckkit/blocks.py` | 350 | 8 | kpi_strip, slot_cards, numbered_steps, branch, problem_solution, compare_table |
| `.claude/skills/pptx-build/scripts/deckkit/kit.py` | 648 | 24 | set_context, C, In, T, R, ARROW |
| `.claude/skills/pptx-build/scripts/deckkit/outline.py` | 88 | 3 | tables, parts, part_of |
| `.claude/skills/pptx-build/scripts/goldenfab/__init__.py` | 8 | 0 |  |
| `.claude/skills/pptx-build/scripts/goldenfab/audit.py` | 829 | 23 | contrast, shape_kind, fill_hex, line_hex, runs_of, box |
| `.claude/skills/pptx-build/scripts/goldenfab/kit.py` | 157 | 7 | load_kit, new_presentation, mix, fit_picture, add_box, set_shape_text |
| `.claude/skills/pptx-build/scripts/scan_material.py` | 204 | 2 | scan, main |
| `.claude/skills/pptx-build/scripts/score_deck.py` | 89 | 1 | main |

합계 **2630줄**

## 3. 기계 검사 배선

| 검사 | 누가 부르나 |
| --- | --- |
| `check_accent` | generic_checks |
| `check_verdict_contrast` | generic_checks |
| `check_fill_ratio` | generic_checks |
| `check_adhoc_card` | generic_checks |
| `check_duplicate_nodes` | generic_checks |
| `check_node_class` | generic_checks |
| `check_text_overflow` | generic_checks |
| `check_picture_overlap` | generic_checks |
| `check_text_collision` | generic_checks |
| `check_bounds` | generic_checks |
| `check_density` | generic_checks |
| `check_icon_vocab` | 스크립트에서 직접 |

## 4. 문서

살아있는 문서 = 하네스가 실제로 읽는 것. 역사 기록은 그때의 사실이라 죽은 참조가 정상이다.

| 파일 | 줄 | 최종 커밋 | 종류 | 없는 참조 |
| --- | ---: | --- | --- | --- |
| `.claude/skills/deck-outline-grill/SKILL.md` | 108 | 2026-08-03 | **살아있음** | — |
| `.claude/skills/pptmaker/SKILL.md` | 226 | 2026-08-04 | **살아있음** | — |
| `.claude/skills/pptx-build/SKILL.md` | 90 | 2026-08-04 | **살아있음** | — |
| `.claude/skills/pptx-build/references/design-rules.md` | 608 | 2026-08-04 | **살아있음** | — |
| `.claude/skills/pptx-build/references/reference-metrics.md` | 38 | 2026-07-06 | **살아있음** | — |
| `.claude/skills/pptx-visuals/SKILL.md` | 82 | 2026-07-29 | **살아있음** | — |
| `.claude/skills/pptx-visuals/references/archetype-catalog.md` | 298 | 2026-08-03 | **살아있음** | — |
| `.claude/skills/pptx-visuals/references/spec-fields.md` | 84 | 2026-07-29 | **살아있음** | — |
| `.claude/skills/pptx-visuals/references/visual-selection.md` | 113 | 2026-08-03 | **살아있음** | — |
| `docs/2026-07-20-dense-failure-analysis.md` | 93 | 2026-07-21 | 역사 | 8개 |
| `docs/2026-07-24-deck-smith-inheritance-experiment.md` | 69 | 2026-08-02 | 역사 | — |
| `docs/2026-07-25-골든-도해-어휘-인벤토리.md` | 262 | 2026-07-27 | 역사 | 12개 |
| `docs/2026-07-25-어휘-정본-이전-계획.md` | 248 | 2026-07-27 | 역사 | 12개 |
| `docs/2026-07-26-부품-도구함-재설계.md` | 419 | 2026-07-27 | 역사 | 20개 |
| `docs/2026-07-26-부품-선택-모델-핸드오프.md` | 194 | 2026-07-27 | 역사 | 8개 |
| `docs/2026-07-26-부품-선택-모델-확정.md` | 177 | 2026-07-27 | 역사 | 3개 |
| `docs/2026-07-27-선택규칙-실전재료-시험.md` | 254 | 2026-07-27 | 역사 | 4개 |
| `docs/2026-07-27-핸드오프-덱-완주.md` | 95 | 2026-07-27 | 역사 | 5개 |
| `docs/2026-07-29-핸드오프-구성층-공백.md` | 202 | 2026-07-29 | 역사 | 10개 |
| `docs/2026-07-29-핸드오프-조립-밀도.md` | 141 | 2026-07-29 | 역사 | 1개 |
| `docs/CHANGELOG.md` | 30 | 2026-08-02 | 역사 | — |
| `docs/handoff-2026-07-20-decision-table-catalog.md` | 54 | 미커밋 | 역사 | — |
| `docs/handoff-2026-07-20-golden-dense.md` | 68 | 미커밋 | 역사 | 19개 |
| `docs/handoff-2026-07-24-audit-and-inheritance.md` | 41 | 미커밋 | 역사 | 2개 |
| `docs/handoff-2026-08-03-다음-작업.md` | 148 | 미커밋 | 역사 | — |
| `docs/superpowers/specs/2026-07-08-visual-diversity-design.md` | 40 | 2026-07-09 | 역사 | 3개 |
| `docs/superpowers/specs/2026-07-15-gate-repair-design.md` | 362 | 2026-07-15 | 역사 | 8개 |
| `docs/superpowers/specs/2026-07-16-golden-adaptive-wiring-design.md` | 68 | 2026-07-16 | 역사 | 7개 |

## 5. 경고

없음.
