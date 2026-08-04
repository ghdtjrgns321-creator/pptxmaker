# 10. 전수 커버리지 — N = 58

## 분모를 두 번 뽑은 이유

> `census.py`는 이 저장소의 파일이 아니라 전역 스킬(`~/.claude/skills/final-report/scripts/census.py`)이다.
> 아래 명령은 그 경로를 줄여 적은 것이다.

`census.py`는 `.claude`를 **하드코딩으로 제외**한다(일반적으로 하네스 디렉터리이므로). 그런데
이 저장소는 `.claude/skills/`가 곧 **제품**이다 — 조판 도구함·게이트·스킬 문서가 전부 그 안에
있다. 그대로 돌리면 분모가 30이 되고, 그 30에는 `deckkit`도 `check_deck.py`도 없다.

그래서 **루트와 `.claude`를 각각 루트로 두고 두 번** 돌려 분모를 합쳤다. 조용한 제외를
만들지 않기 위해서다.

```bash
uv run python census.py inventory .        --out inventory.json          # N = 30
uv run python census.py inventory .claude  --out inventory-claude.json   # N = 28
```

| 분모      | 정독 N | 목록만 | 제외 디렉터리로 걸러진 파일                                                                                  |
| --------- | -----: | -----: | ------------------------------------------------------------------------------------------------------------ |
| 루트      | **30** |     31 | 5,455 (`.claude` 139 · `.git` 1,828 · `.ruff_cache` 82 · `.venv` 3,092 · `__pycache__` 2 · `_workspace` 312) |
| `.claude` | **28** |     96 | 14 (`__pycache__`)                                                                                           |
| **합계**  | **58** |    127 |                                                                                                              |

### 제외분 (조용한 제외 금지)

| 디렉터리                                         |  건수 | 사유                                                                                                                                                              |
| ------------------------------------------------ | ----: | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `.venv` · `.git` · `.ruff_cache` · `__pycache__` | 5,004 | 빌드·캐시·VCS 내부                                                                                                                                                |
| `_workspace`                                     |   312 | 덱 작업 산출물(계약·조판 스크립트·pptx·렌더). `.gitignore` 대상이며 하네스가 아니다. 단 이 보고서는 그 안의 `deck.py`·계약·변동성 실험을 **실증 예시로 인용**한다 |
| `.claude`(루트 census 기준)                      |   139 | 두 번째 census가 분모로 삼았다 — 실제 제외 아님                                                                                                                   |

---

## 파일 → 장 매핑

### 루트 (N = 30)

| 파일                                                                 | kind   | 반영 장 | 비고                                                        |
| -------------------------------------------------------------------- | ------ | ------- | ----------------------------------------------------------- |
| `.gitignore`                                                         | config | 8       | 산출물·핸드오프 제외 규칙. `golden/render-부품` 주석은 낡음 |
| `CLAUDE.md`                                                          | doc    | 1, 6, 8 | 규칙 단일 출처. 스테일 2건 기록                             |
| `README.md`                                                          | doc    | 8       | **스테일 19건** — 폐기된 4단계 구조를 설명                  |
| `pyproject.toml`                                                     | config | 1       | 스택 표의 출처                                              |
| `docs/MAP.md`                                                        | doc    | 6       | 자동 생성물. 동기성 검증 완료                               |
| `docs/CHANGELOG.md`                                                  | doc    | 9       | 사고·전환점 원천                                            |
| `scripts/make_map.py`                                                | code   | 6       | 지도 생성기                                                 |
| `scripts/deck_state.py`                                              | code   | 5, 6    | 상태·획일화·반려                                            |
| `scripts/check_outline.py`                                           | code   | 3       | 계약↔재료. `deck_state init`이 부른다                       |
| `_plan/verify_no_archive_refs.py`                                    | code   | 8       | 미배선 + 현재 exit 1                                        |
| `_plan/tracked.txt`                                                  | text   | 9       | 2026-07-29 정리 직전 스냅샷(371행). 지금 안 쓰임            |
| `_plan/untracked.txt`                                                | text   | 9       | 같음(20행)                                                  |
| `docs/2026-07-20-dense-failure-analysis.md`                          | doc    | 9       | 사고 3                                                      |
| `docs/2026-07-24-deck-smith-inheritance-experiment.md`               | doc    | 9       | 결론 3건 무효 표기                                          |
| `docs/2026-07-25-골든-도해-어휘-인벤토리.md`                         | doc    | 9       | 걷어낸 구조(시각 어휘)                                      |
| `docs/2026-07-25-어휘-정본-이전-계획.md`                             | doc    | 9       | 사고 2 원천                                                 |
| `docs/2026-07-26-부품-도구함-재설계.md`                              | doc    | 9       | 걷어낸 구조(장 변형 모듈)                                   |
| `docs/2026-07-26-부품-선택-모델-핸드오프.md`                         | doc    | 9       | 걷어낸 구조                                                 |
| `docs/2026-07-26-부품-선택-모델-확정.md`                             | doc    | 9       | 걷어낸 구조                                                 |
| `docs/2026-07-27-선택규칙-실전재료-시험.md`                          | doc    | 9       | 추측값 10종 중 9종 오류                                     |
| `docs/2026-07-27-핸드오프-덱-완주.md`                                | doc    | 9       | 걷어낸 구조                                                 |
| `docs/2026-07-29-핸드오프-구성층-공백.md`                            | doc    | 9       | 사고 4                                                      |
| `docs/2026-07-29-핸드오프-조립-밀도.md`                              | doc    | 9       | 사고 4                                                      |
| `docs/handoff-2026-07-20-decision-table-catalog.md`                  | doc    | 9       | 사고 3                                                      |
| `docs/handoff-2026-07-20-golden-dense.md`                            | doc    | 9       | 사고 3 (상단 폐기 상자)                                     |
| `docs/handoff-2026-07-24-audit-and-inheritance.md`                   | doc    | 9       | 걷어낸 구조(훅)                                             |
| `docs/handoff-2026-08-03-다음-작업.md`                               | doc    | 5, 8    | 게이트 실행 명령의 유일한 근거                              |
| `docs/superpowers/specs/2026-07-08-visual-diversity-design.md`       | doc    | 9       | 어휘 3종 수렴                                               |
| `docs/superpowers/specs/2026-07-15-gate-repair-design.md`            | doc    | 9       | 사고 1 원천                                                 |
| `docs/superpowers/specs/2026-07-16-golden-adaptive-wiring-design.md` | doc    | 9       | 걷어낸 구조                                                 |

### `.claude` (N = 28)

| 파일                                                             | kind   | 반영 장 | 비고                                                                          |
| ---------------------------------------------------------------- | ------ | ------- | ----------------------------------------------------------------------------- |
| `.claude/settings.json`                                          | config | 5, 8    | `{"hooks":{}}` — 프로젝트 훅 0개                                              |
| `.claude/skills/pptmaker/SKILL.md`                               | doc    | 2       | 파이프라인 ①~⑦ 단일 출처                                                      |
| `.claude/skills/deck-outline-grill/SKILL.md`                     | doc    | 3       | 계약 서식·인터뷰 절차                                                         |
| `.claude/skills/pptx-build/SKILL.md`                             | doc    | 4, 5    | 도구함 인덱스                                                                 |
| `.claude/skills/pptx-visuals/SKILL.md`                           | doc    | 8       | 미배선 모듈의 사용법                                                          |
| `.claude/skills/pptx-build/scripts/deckkit/__init__.py`          | code   | 4       | 패키지 마커, 코드 0줄                                                         |
| `.claude/skills/pptx-build/scripts/deckkit/brand.yaml`           | config | 4       | 토큰 33개 전수 전사                                                           |
| `.claude/skills/pptx-build/scripts/deckkit/kit.py`               | code   | 4       | 원시도구·골격·부 내비                                                         |
| `.claude/skills/pptx-build/scripts/deckkit/blocks.py`            | code   | 4       | 블록 8종                                                                      |
| `.claude/skills/pptx-build/scripts/deckkit/outline.py`           | code   | 3       | 계약 판독                                                                     |
| `.claude/skills/pptx-build/scripts/check_deck.py`                | code   | 5       | 하드페일 9종 + 보고 1종                                                       |
| `.claude/skills/pptx-build/scripts/score_deck.py`                | code   | 5       | 오딧 러너                                                                     |
| `.claude/skills/pptx-build/scripts/render_deck.ps1`              | code   | 5       | COM 렌더                                                                      |
| `.claude/skills/pptx-build/scripts/scan_material.py`             | code   | 2       | 재료 계수(선택 도구)                                                          |
| `.claude/skills/pptx-build/scripts/goldenfab/__init__.py`        | code   | 5       | 패키지 마커                                                                   |
| `.claude/skills/pptx-build/scripts/goldenfab/audit.py`           | code   | 5       | 검사 12종                                                                     |
| `.claude/skills/pptx-build/scripts/goldenfab/kit.py`             | code   | 5, 8    | 7함수 중 1개만 배선                                                           |
| `.claude/skills/pptx-build/references/design-rules.md`           | doc    | 8       | 608줄. 강제 태그 0건 확인                                                     |
| `.claude/skills/pptx-build/references/reference-metrics.md`      | doc    | 8       | 인용값 6개 중 4개 불일치                                                      |
| `.claude/skills/pptx-build/assets/brand-kit.yaml`                | config | 8       | 두 번째 토큰 파일                                                             |
| `.claude/skills/pptx-build/assets/golden-snapshot.json`          | config | 5       | 밀도 밴드의 유일한 파생 원천. **large — 구조 정독**(17 slides 배열 + `names`) |
| `.claude/skills/pptx-build/assets/sample-material-pptmaker.json` | config | 8       | 미배선. 시험 입력 데이터 잔존                                                 |
| `.claude/skills/pptx-visuals/scripts/visuals.py`                 | code   | 8       | **미배선 809줄**                                                              |
| `.claude/skills/pptx-visuals/scripts/mpl_exhibits.py`            | code   | 8       | **미배선 395줄**                                                              |
| `.claude/skills/pptx-visuals/scripts/make_icons.py`              | code   | 1, 8    | 아이콘 생성기(일회성)                                                         |
| `.claude/skills/pptx-visuals/references/archetype-catalog.md`    | doc    | 6, 8    | hollow 행 5개                                                                 |
| `.claude/skills/pptx-visuals/references/spec-fields.md`          | doc    | 8       | 미배선 모듈의 필드 규격                                                       |
| `.claude/skills/pptx-visuals/references/visual-selection.md`     | doc    | 8       | 폐기 어휘를 교정책으로 권함                                                   |

### 목록만 (정독 면제)

| 묶음                                                            |            건수 | 사유 · 반영 장                                         |
| --------------------------------------------------------------- | --------------: | ------------------------------------------------------ |
| `golden/render/s01.png` ~ `s17.png`                             |              17 | 이미지 자산. 골든 17장 — 5장에서 개수·역할로 다룸      |
| `docs/images/*.png`                                             |               5 | README 삽화                                            |
| `.claude/skills/pptx-build/assets/icons/**`                     | 72 + `_svg/` 24 | 균일 자산(24종 × 3색). 4·8장에서 개수·생성 경로로 다룸 |
| `golden/golden-deck.pptx` · `results/local-ai-assist-소개.pptx` |               2 | 바이너리 산출물. 5장에서 도형 수 실측으로 다룸         |
| `_plan/move-*.tsv`                                              |               4 | 이동 대장. 9장에서 성격만                              |
| `uv.lock` · `.python-version` · `.gitattributes`                |               3 | 잠금·환경 파일. 1장 스택 표                            |

---

### 부록 — 두 번째 census 기준 경로

`.claude`를 루트로 돌린 census는 경로를 그 디렉터리 기준으로 낸다. 대조가 성립하도록 같은
28개 파일을 그 기준으로 다시 적는다. **위 표와 같은 파일이며 중복이 아니라 좌표계 차이다.**

| 파일                                                     | kind   | 비고              |
| -------------------------------------------------------- | ------ | ----------------- |
| `settings.json`                                          | config | 위 표와 동일 항목 |
| `skills/deck-outline-grill/SKILL.md`                     | doc    | 위 표와 동일 항목 |
| `skills/pptmaker/SKILL.md`                               | doc    | 위 표와 동일 항목 |
| `skills/pptx-build/assets/brand-kit.yaml`                | config | 위 표와 동일 항목 |
| `skills/pptx-build/assets/golden-snapshot.json`          | config | large — 구조 정독 |
| `skills/pptx-build/assets/sample-material-pptmaker.json` | config | 위 표와 동일 항목 |
| `skills/pptx-build/references/design-rules.md`           | doc    | 위 표와 동일 항목 |
| `skills/pptx-build/references/reference-metrics.md`      | doc    | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/check_deck.py`                | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/deckkit/__init__.py`          | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/deckkit/blocks.py`            | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/deckkit/brand.yaml`           | config | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/deckkit/kit.py`               | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/deckkit/outline.py`           | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/goldenfab/__init__.py`        | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/goldenfab/audit.py`           | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/goldenfab/kit.py`             | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/render_deck.ps1`              | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/scan_material.py`             | code   | 위 표와 동일 항목 |
| `skills/pptx-build/scripts/score_deck.py`                | code   | 위 표와 동일 항목 |
| `skills/pptx-build/SKILL.md`                             | doc    | 위 표와 동일 항목 |
| `skills/pptx-visuals/references/archetype-catalog.md`    | doc    | 위 표와 동일 항목 |
| `skills/pptx-visuals/references/spec-fields.md`          | doc    | 위 표와 동일 항목 |
| `skills/pptx-visuals/references/visual-selection.md`     | doc    | 위 표와 동일 항목 |
| `skills/pptx-visuals/scripts/make_icons.py`              | code   | 위 표와 동일 항목 |
| `skills/pptx-visuals/scripts/mpl_exhibits.py`            | code   | 위 표와 동일 항목 |
| `skills/pptx-visuals/scripts/visuals.py`                 | code   | 위 표와 동일 항목 |
| `skills/pptx-visuals/SKILL.md`                           | doc    | 위 표와 동일 항목 |

## verify 실행

```
$ uv run python census.py verify . FINAL-REPORT/10_COVERAGE.md
인벤토리 N = 30, 커버리지 표 경로 = 110 + 글롭 3
PASS — 전수 커버리지 30/30
EXIT=0

$ uv run python census.py verify .claude FINAL-REPORT/10_COVERAGE.md
인벤토리 N = 28, 커버리지 표 경로 = 110 + 글롭 3
PASS — 전수 커버리지 28/28
EXIT=0
```

## 이 커버리지가 증명하는 것과 못 하는 것

**증명하는 것**: 58개 파일을 전부 열어 읽었고, 각각이 어느 장에 반영됐는지 대응된다.

**증명하지 못하는 것**: 맞게 썼는지. 그래서 이 보고서는 커버리지 외에 **배선 실측**을 별도로
걸었다 — 본문의 현재형 동작 서술은 `grep -rn`으로 호출자를 확인한 것만 쓰고, 호출자가 0이면
"미배선"으로 적었다. `pptx-visuals` 1,204줄이 그렇게 본문에서 [8장 한계](8_TEST-DECISIONS.md)로
내려갔다.
