# 13. 전수 커버리지

## 분모

이 프로젝트는 **하네스 자체가 제품**이라 `.claude/`가 본체다. 그런데 census 스크립트의
`EXCLUDE_DIRS`가 `.claude`를 기본 제외하므로 **두 번 돌려 합쳤다.**

```bash
# ① 프로젝트 루트 (.claude는 스크립트가 자동 제외)
python census.py inventory . --exclude _archive --exclude _lab --out inventory.json

# ② 하네스 본체
python census.py inventory .claude --out inventory-claude.json
```

| 인벤토리  | 정독 N | 목록만  |
| --------- | ------ | ------- |
| 루트      | 27     | 145     |
| `.claude` | 26     | 96      |
| **합계**  | **53** | **241** |

### 제외 표

| 디렉토리                                   | 건수  | 사유                                                                                                                 |
| ------------------------------------------ | ----- | -------------------------------------------------------------------------------------------------------------------- |
| `_archive`                                 | 1,415 | 2026-07-29 정리분 1,400 + 구판 FINAL-REPORT 15. 살아있는 코드가 참조하지 않음이 기계로 검증됨([10](10_TEST-DECISIONS.md) 검증 1·2) |
| `_lab`                                     | 61    | 실험판 작업 디렉토리(gitignore). 덱 한 벌이 끝나면 `_workspace/`로 정리 예정                                         |
| `_workspace`                               | 40    | 스킬 작업공간 — 아웃라인 계약·감사 스크래치(스크립트 기본 제외)                                                      |
| `FINAL-REPORT`                             | 15    | **이 보고서 자신**(구판은 `_archive/FINAL-REPORT-구판/`으로 이동)                                                    |
| `.venv`·`.git`·`.ruff_cache`·`__pycache__` | 6,838 | 빌드·캐시                                                                                                            |

## 파일 → 장 매핑

### 하네스 — 경로는 `.claude/` 기준 상대경로

> `.claude`를 root로 verify하기 때문에 이 표만 `.claude/`를 뗀 상대경로로 적는다.

| 파일                                                             | kind   | 반영 장   | 비고                                       |
| ---------------------------------------------------------------- | ------ | --------- | ------------------------------------------ |
| `skills/pptmaker/SKILL.md`                               | doc    | 2 · 4 · 8 | 파이프라인 ①~⑦ · 확정 규칙 5 · 반려 목록 8 |
| `skills/deck-outline-grill/SKILL.md`                     | doc    | 3         | 아웃라인 계약 인터뷰                       |
| `skills/pptx-build/SKILL.md`                             | doc    | 5         | 조판 도구함 진입                           |
| `skills/pptx-visuals/SKILL.md`                           | doc    | 6         | 시각 어휘 진입                             |
| `agents/deck-smith.md`                                   | doc    | 8         | 배치 사다리 4층                            |
| `settings.json`                                          | config | 11        | 훅 배선 제거(v7)                           |
| `skills/pptx-build/assets/brand-kit.yaml`                | config | 5         | 색 6·폰트 3·크기 8                         |
| `skills/pptx-build/assets/golden-snapshot.json`          | config | 5 · 7     | large — 구조 정독(17장·도형 850)           |
| `skills/pptx-build/assets/sample-material-pptmaker.json` | config | 3         | `scan_material` 산출 샘플                  |
| `skills/pptx-build/references/design-rules.md`           | doc    | 5 · 9     | 디자인 규율 639줄                          |
| `skills/pptx-build/references/reference-metrics.md`      | doc    | 5         | 실물 레퍼런스 실측                         |
| `skills/pptx-build/scripts/goldenfab/audit.py`           | code   | 7         | `generic_checks` 12 · `density_band`       |
| `skills/pptx-build/scripts/goldenfab/dense.py`           | code   | 5         | 밀도 부품                                  |
| `skills/pptx-build/scripts/goldenfab/grid.py`            | code   | 5         | 좌표계 · `pitch`/`track`                   |
| `skills/pptx-build/scripts/goldenfab/kit.py`             | code   | 5         | 원시도구 · `new_presentation`              |
| `skills/pptx-build/scripts/goldenfab/layouts.py`         | code   | 5         | 정형 3종                                   |
| `skills/pptx-build/scripts/goldenfab/__init__.py`        | code   | 5         | 패키지 서술                                |
| `skills/pptx-build/scripts/render_deck.ps1`              | code   | 7         | COM 렌더 · 뮤텍스 직렬화                   |
| `skills/pptx-build/scripts/score_deck.py`                | code   | 7         | 채점 러너                                  |
| `skills/pptx-build/scripts/scan_material.py`             | code   | 3         | 재료 계수                                  |
| `skills/pptx-visuals/references/spec-fields.md`          | doc    | 6         | 렌더러 필드 단일 출처                      |
| `skills/pptx-visuals/references/visual-selection.md`     | doc    | 6         | 형식 결정표                                |
| `skills/pptx-visuals/references/archetype-catalog.md`    | doc    | 6         | 아키타입 · (A)는 그림 사전                 |
| `skills/pptx-visuals/scripts/visuals.py`                 | code   | 6         | 네이티브 6 + 도형 9 + 아이콘               |
| `skills/pptx-visuals/scripts/mpl_exhibits.py`            | code   | 6         | mpl 9종                                    |
| `skills/pptx-visuals/scripts/make_icons.py`              | code   | 6         | 아이콘 생성(pymupdf 별도)                  |

### 루트

| 파일                                                                 | kind   | 반영 장 | 비고                                            |
| -------------------------------------------------------------------- | ------ | ------- | ----------------------------------------------- |
| `CLAUDE.md`                                                          | doc    | 1 · 8   | 작업 모드 · 반려 목록 안내                      |
| `README.md`                                                          | doc    | 9       | **한계 9** — 아카이브된 구조를 서술 중(미갱신)  |
| `pyproject.toml`                                                     | config | 1       | 기술 스택                                       |
| `.gitignore`                                                         | config | 5 · 11  | 골든 기준선을 커밋 대상으로 전환한 근거         |
| `ref/catalog.md`                                                     | doc    | 5       | 레퍼런스 크롭 목록                              |
| `_plan/verify_no_archive_refs.py`                                    | code   | 9 · 10  | 아카이브 참조 검증기 1부·2부                    |
| `_plan/tracked.txt`                                                  | text   | 13      | 정리 시점 조사 스냅샷(일회성·gitignore)         |
| `_plan/untracked.txt`                                                | text   | 13      | 같음                                            |
| `docs/CHANGELOG.md`                                                  | doc    | 11      | 하네스 변경 전체 이력                           |
| `docs/PIPELINE.md`                                                   | doc    | 9       | **한계 9** — v6 구조를 "현재"로 서술 중(미갱신) |
| `docs/2026-07-20-dense-failure-analysis.md`                          | doc    | 11      | dense 개편 전환점                               |
| `docs/2026-07-24-deck-smith-inheritance-experiment.md`               | doc    | 8 · 11  | 상속 루프 실험                                  |
| `docs/2026-07-25-골든-도해-어휘-인벤토리.md`                         | doc    | 6 · 11  | (A)/(B) 두 축 재수립                            |
| `docs/2026-07-25-어휘-정본-이전-계획.md`                             | doc    | 6 · 11  | 어휘 이전 계획                                  |
| `docs/2026-07-26-부품-도구함-재설계.md`                              | doc    | 11      | v6 3층 분리                                     |
| `docs/2026-07-26-부품-선택-모델-핸드오프.md`                         | doc    | 11      | 선택 모델                                       |
| `docs/2026-07-26-부품-선택-모델-확정.md`                             | doc    | 11      | 선택 모델 확정                                  |
| `docs/2026-07-27-선택규칙-실전재료-시험.md`                          | doc    | 3 · 12  | 화살표 오판(사고 4) 실측                        |
| `docs/2026-07-27-핸드오프-덱-완주.md`                                | doc    | 11      | 완주 시도 기록                                  |
| `docs/2026-07-29-핸드오프-구성층-공백.md`                            | doc    | 11      | v7 직전 진단                                    |
| `docs/2026-07-29-핸드오프-조립-밀도.md`                              | doc    | 11      | v7 직전 진단                                    |
| `docs/handoff-2026-07-20-decision-table-catalog.md`                  | doc    | 11      | 결정표 카탈로그                                 |
| `docs/handoff-2026-07-20-golden-dense.md`                            | doc    | 11      | 골든 dense                                      |
| `docs/handoff-2026-07-24-audit-and-inheritance.md`                   | doc    | 8 · 11  | 오딧·상속                                       |
| `docs/superpowers/specs/2026-07-08-visual-diversity-design.md`       | doc    | 11      | v4 다양성 설계                                  |
| `docs/superpowers/specs/2026-07-15-gate-repair-design.md`            | doc    | 7 · 12  | 게이트 수리(사고 1·6)                           |
| `docs/superpowers/specs/2026-07-16-golden-adaptive-wiring-design.md` | doc    | 11      | 골든 적응 배선                                  |

## 목록만 (정독 면제)

| 글롭                                        | 건수 | 사유 · 대표 표본을 다룬 장                                                                 |
| ------------------------------------------- | ---- | ------------------------------------------------------------------------------------------ |
| `ref/**`                                    | 75   | 균일 데이터 — 레퍼런스 크롭 PNG. 목록은 `ref/catalog.md`(정독), 실측치는 [5](5_COMPOSE.md) |
| `skills/pptx-build/assets/icons/**` | 96   | 균일 데이터 — Lucide 아이콘 3색 변형. 생성기 `make_icons.py` 정독, [6](6_VISUALS.md)       |
| `golden/render/**`                          | 17   | 균일 데이터 — 골든 렌더 PNG. 자산 설명 [5](5_COMPOSE.md)                                   |
| `golden/render-부품/**`                     | 27   | 균일 데이터 — 부품 렌더 PNG. 자산 설명 [5](5_COMPOSE.md)                                   |
| `data/**`                                   | 13   | 대용량 저작물 PDF(gitignore) — 밀도 실측 원천. [5](5_COMPOSE.md)                           |
| `docs/images/**`                            | 5    | 균일 데이터 — 문서 삽화                                                                    |
| `_plan/**`                                  | 4    | 이동 대장 TSV — [11](11_JOURNEY.md) v7                                                     |
| `golden/golden-deck.pptx`                   | 1    | 바이너리 — 도형 서명은 스냅샷으로 정독([5](5_COMPOSE.md)·[7](7_GATES.md))                  |
| 루트 바이너리·잠금                          | 3    | `uv.lock` 등                                                                               |

## verify 출력

```
$ python census.py verify . FINAL-REPORT/13_COVERAGE.md --exclude _archive --exclude _lab
인벤토리 N = 27, 커버리지 표 경로 = 75 + 글롭 7
PASS — 전수 커버리지 27/27

$ python census.py verify .claude FINAL-REPORT/13_COVERAGE.md
인벤토리 N = 26, 커버리지 표 경로 = 75 + 글롭 7
PASS — 전수 커버리지 26/26
```

합계 **53/53**.
