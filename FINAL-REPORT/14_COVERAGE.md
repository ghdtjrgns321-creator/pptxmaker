# 14. 전수 커버리지 부록 (COVERAGE)

> 이 장은 "다 읽었다"의 기계 증명이다. 분모 N은 손이 아니라 `census.py inventory`가 고정했고,
> 아래 매핑 표가 N개 파일 전부를 반영 장에 대응시킨다. 말미의 verify PASS 출력이 차집합 0의 증거다.

## 분모 N = 108 (루트 48 + .claude 서브트리 60)

census 기본 제외 상수(EXCLUDE_DIRS)에 `.claude`가 포함되어 있으나, 이 프로젝트는 본체 코드가
`.claude/skills/`에 있다. 따라서 인벤토리를 2회 실행해 분모를 합산했다 — 스크립트 무수정, 재현 가능.

실행 커맨드(재현용):

```
python <final-report 스킬>/scripts/census.py inventory <프로젝트 루트> --exclude _workspace_prev
python <final-report 스킬>/scripts/census.py inventory <프로젝트 루트>/.claude --exclude _workspace_prev --exclude state
```

| 실행    | 정독 분모 | kind 내역                           | 목록만(정독 면제) |
| ------- | --------- | ----------------------------------- | ----------------- |
| 루트    | 48        | code 4 · config 8 · doc 34 · text 2 | 376               |
| .claude | 60        | code 41 · config 3 · doc 16         | 99                |

## 제외 내역 (조용한 제외 금지 — 전량 기록)

| 디렉토리               | 걸러진 파일 | 구분                                 | 사유                                                                       |
| ---------------------- | ----------- | ------------------------------------ | -------------------------------------------------------------------------- |
| `.git`                 | 926         | 기본 제외                            | VCS 내부                                                                   |
| `.venv`                | 3,090       | 기본 제외                            | uv 가상환경(재현 가능 산출물)                                              |
| `.ruff_cache`          | 34          | 기본 제외                            | 린터 캐시                                                                  |
| `__pycache__`          | 20 + 32     | 기본 제외                            | 바이트코드 캐시                                                            |
| `.claude`              | 242         | 기본 제외 → **별도 인벤토리로 회수** | 본체 코드 소재 — 2차 실행(N=60)으로 분모 편입                              |
| `_workspace_prev`      | 983 + 4     | `--exclude` 명시                     | 폐기된 과거 스킬 작업 스크래치(v4 이전) — 여정 사료는 docs·git 이력이 대체 |
| `state` (.claude 하위) | 46          | `--exclude` 명시                     | 하네스 런타임 상태(계약·검증 픽스처) — gitignore 대상, 코드가 아님         |

## 파일 → 장 매핑 (108/108)

| 파일                                                                 | kind   | 반영 장             | 비고                                                            |
| -------------------------------------------------------------------- | ------ | ------------------- | --------------------------------------------------------------- |
| `.claude/agents/consistency-qa.md`                                   | doc    | 2                   | 파이프라인 에이전트 정의                                        |
| `.claude/agents/deck-composer.md`                                    | doc    | 2                   | 파이프라인 에이전트 정의                                        |
| `.claude/agents/pptx-builder.md`                                     | doc    | 2                   | 파이프라인 에이전트 정의                                        |
| `.claude/skills/consistency-qa/SKILL.md`                             | doc    | 8                   |                                                                 |
| `.claude/skills/consistency-qa/scripts/audit_pptx.py`                | code   | 8                   |                                                                 |
| `.claude/skills/consistency-qa/scripts/check_contract.py`            | code   | 8                   |                                                                 |
| `.claude/skills/deck-compose/SKILL.md`                               | doc    | 4                   |                                                                 |
| `.claude/skills/deck-compose/assets/standard-deck-spec.json`         | config | 4                   |                                                                 |
| `.claude/skills/deck-compose/references/golden-content-contract.md`  | doc    | 4                   |                                                                 |
| `.claude/skills/deck-compose/references/layout-matching.md`          | doc    | 4                   |                                                                 |
| `.claude/skills/deck-outline-grill/SKILL.md`                         | doc    | 3                   |                                                                 |
| `.claude/skills/pptmaker/SKILL.md`                                   | doc    | 2                   | 오케스트레이터                                                  |
| `.claude/skills/pptx-build/SKILL.md`                                 | doc    | 5                   |                                                                 |
| `.claude/skills/pptx-build/assets/brand-kit.yaml`                    | config | 5                   |                                                                 |
| `.claude/skills/pptx-build/assets/golden-snapshot.json`              | config | 7·8                 | 구조 정독(146KB — 최상위 키·17장·519도형 실측)                  |
| `.claude/skills/pptx-build/references/deck-spec-schema.md`           | doc    | 5                   |                                                                 |
| `.claude/skills/pptx-build/references/design-rules.md`               | doc    | 5                   |                                                                 |
| `.claude/skills/pptx-build/references/reference-metrics.md`          | doc    | 5                   |                                                                 |
| `.claude/skills/pptx-build/scripts/audit_deck.py`                    | code   | 8                   |                                                                 |
| `.claude/skills/pptx-build/scripts/audit_golden.py`                  | code   | 8                   |                                                                 |
| `.claude/skills/pptx-build/scripts/build_pptx.py`                    | code   | 5                   |                                                                 |
| `.claude/skills/pptx-build/scripts/compare_golden.py`                | code   | 8                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/__init__.py`            | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/_variant_h.py`          | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/_variant_k.py`          | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/audit.py`               | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/common.py`              | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/content_contract.py`    | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/grid.py`                | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/kit.py`                 | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/layouts.py`             | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/reference.py`           | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/registry.py`            | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s06_pilot.py`           | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s06_proto_f.py`         | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s06_variants.py`        | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s08_variants.py`        | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s09_variants.py`        | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s10_screenshot.py`      | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s11_branch_snap.py`     | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s11_variants.py`        | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s12_variants.py`        | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s14_variants.py`        | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s15_variants.py`        | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s17_variants.py`        | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s18_variants.py`        | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/s21_closing.py`         | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/goldenfab/shape_kind.py`          | code   | 7                   |                                                                 |
| `.claude/skills/pptx-build/scripts/render_deck.ps1`                  | code   | 5·9                 | 미추적 신규(수동 렌더 보조)                                     |
| `.claude/skills/pptx-build/scripts/verify_content_contract.py`       | code   | 8                   |                                                                 |
| `.claude/skills/pptx-visuals/SKILL.md`                               | doc    | 6                   |                                                                 |
| `.claude/skills/pptx-visuals/references/archetype-catalog.md`        | doc    | 6                   |                                                                 |
| `.claude/skills/pptx-visuals/references/visual-selection.md`         | doc    | 6                   |                                                                 |
| `.claude/skills/pptx-visuals/scripts/check_candidates.py`            | code   | 6                   |                                                                 |
| `.claude/skills/pptx-visuals/scripts/make_icons.py`                  | code   | 6                   |                                                                 |
| `.claude/skills/pptx-visuals/scripts/make_mockups.py`                | code   | 6                   |                                                                 |
| `.claude/skills/pptx-visuals/scripts/mpl_exhibits.py`                | code   | 6                   |                                                                 |
| `.claude/skills/pptx-visuals/scripts/recommend_archetypes.py`        | code   | 6                   |                                                                 |
| `.claude/skills/pptx-visuals/scripts/render_real_mockups.py`         | code   | 6                   |                                                                 |
| `.claude/skills/pptx-visuals/scripts/visuals.py`                     | code   | 6                   |                                                                 |
| `.gitignore`                                                         | config | 10 (증거 소실 근거) |                                                                 |
| `CLAUDE.md`                                                          | doc    | 1·12                | 변경 이력 표가 여정 사료                                        |
| `_workspace_kifrs/01.5_outline.md`                                   | doc    | 9                   |                                                                 |
| `_workspace_kifrs/02_deck-spec.json`                                 | config | 9                   |                                                                 |
| `_workspace_kifrs/03_exhibit-candidates.json`                        | config | 9                   |                                                                 |
| `_workspace_kifrs/03_qa-report.md`                                   | doc    | 8·9                 | QA FAIL 실측 기록                                               |
| `_workspace_kifrs/04_contract-check.md`                              | doc    | 8·9                 | QA FAIL 실측 기록                                               |
| `_workspace_kifrs/05_ripple-golden.md`                               | doc    | 9                   |                                                                 |
| `_workspace_kifrs/06_s4-color-inventory.md`                          | doc    | 9                   |                                                                 |
| `_workspace_kifrs/07-compare-before.txt`                             | text   | 9                   |                                                                 |
| `_workspace_kifrs/07-s8-s9-color.md`                                 | doc    | 9                   |                                                                 |
| `_workspace_kifrs/07-s8-s9-probe.txt`                                | text   | 9                   |                                                                 |
| `_workspace_kifrs/07-s8-s9-selftest.md`                              | doc    | 9                   |                                                                 |
| `_workspace_kifrs/07-s8-s9-verify.md`                                | doc    | 9                   |                                                                 |
| `_workspace_kifrs/HANDOFF.md`                                        | doc    | 9                   |                                                                 |
| `_workspace_kifrs/dump_contract.py`                                  | code   | 9                   | 검증 스크립트(미배선)                                           |
| `_workspace_kifrs/golden_defaults.json`                              | config | 9                   |                                                                 |
| `_workspace_kifrs/mockups/gallery.html`                              | config | 9                   |                                                                 |
| `_workspace_kifrs/param-audit.md`                                    | doc    | 9                   |                                                                 |
| `_workspace_kifrs/verify_s8s9_data.py`                               | code   | 9                   | 검증 스크립트(미배선)                                           |
| `docs/PIPELINE.md`                                                   | doc    | 2                   |                                                                 |
| `docs/handoff-2026-07-18-golden-densify.md`                          | doc    | 7·9                 | 2026-07-18 밀도 개선 핸드오프                                   |
| `docs/superpowers/specs/2026-07-08-visual-diversity-design.md`       | doc    | 11·12               | ADR·전환 근거                                                   |
| `docs/superpowers/specs/2026-07-15-gate-repair-design.md`            | doc    | 11·12               | ADR·전환 근거                                                   |
| `docs/superpowers/specs/2026-07-16-golden-adaptive-wiring-design.md` | doc    | 11·12               | ADR·전환 근거                                                   |
| `docs/user/00-INDEX.md`                                              | doc    | 12                  | 19장 시절 동결 — 스테일 자체를 기록                             |
| `docs/user/01-journey.md`                                            | doc    | 12                  | 19장 시절 동결 — 스테일 자체를 기록                             |
| `docs/user/02-failure-analysis.md`                                   | doc    | 12                  | 19장 시절 동결 — 스테일 자체를 기록                             |
| `docs/user/03-golden-deck.md`                                        | doc    | 12                  | 19장 시절 동결 — 스테일 자체를 기록                             |
| `docs/user/04-pipeline.md`                                           | doc    | 12                  | 19장 시절 동결 — 스테일 자체를 기록                             |
| `docs/user/05-golden-deck-plan.md`                                   | doc    | 12                  | 19장 시절 동결 — 스테일 자체를 기록                             |
| `docs/user/06-outline-grill.md`                                      | doc    | 12                  | 19장 시절 동결 — 스테일 자체를 기록                             |
| `docs/user/07-factory-port.md`                                       | doc    | 12                  | 19장 시절 동결 — 스테일 자체를 기록                             |
| `docs/user/08-system-overview.md`                                    | doc    | 12                  | 19장 시절 동결 — 스테일 자체를 기록                             |
| `golden/00_factsheet.md`                                             | doc    | 7                   |                                                                 |
| `golden/00_phase0-crosscheck.md`                                     | doc    | 7                   |                                                                 |
| `golden/_pilot_s06/ref_anatomy.md`                                   | doc    | 7·9                 | s06 파일럿 해부·단위                                            |
| `golden/_pilot_s06/units.md`                                         | doc    | 7·9                 | s06 파일럿 해부·단위                                            |
| `golden/backup/s06_pilot_v3.py`                                      | code   | 7·12                | 백업(과거형)                                                    |
| `golden/backup/v5-monochrome-2026-07-18/brand-kit.yaml`              | config | 7·12                | 백업(과거형)                                                    |
| `golden/build_golden.py`                                             | code   | 7                   |                                                                 |
| `golden/variants/compare_full.md`                                    | doc    | 7·8                 | 스냅샷 회귀 최근 판정                                           |
| `golden/variants/desktop.ini`                                        | config | 14                  | 탐색기 폴더 설정 파일 — 내용 무의미(비정독 사유 아님·전문 확인) |
| `input/k-ifrs-1115/claude-report-v2.md`                              | doc    | 9                   | 파일럿 입력 재료                                                |
| `input/k-ifrs-1115/notebooklm-report-1.md`                           | doc    | 9                   | 파일럿 입력 재료                                                |
| `input/k-ifrs-1115/notebooklm-report-v2.md`                          | doc    | 9                   | 파일럿 입력 재료                                                |
| `pyproject.toml`                                                     | config | 1                   | 기술 스택 실측                                                  |
| `ref/catalog.md`                                                     | doc    | 6                   | 레퍼런스 카탈로그                                               |

## .claude 서브트리 verify 대조 경로 (60건)

verify 2차 실행은 `.claude`를 루트로 삼아 상대경로로 대조한다. 본표의 `.claude/...` 표기와 동일 파일이다.

| verify 기준 경로                                            | 본표 경로                                                           |
| ----------------------------------------------------------- | ------------------------------------------------------------------- |
| `agents/consistency-qa.md`                                  | `.claude/agents/consistency-qa.md`                                  |
| `agents/deck-composer.md`                                   | `.claude/agents/deck-composer.md`                                   |
| `agents/pptx-builder.md`                                    | `.claude/agents/pptx-builder.md`                                    |
| `skills/consistency-qa/SKILL.md`                            | `.claude/skills/consistency-qa/SKILL.md`                            |
| `skills/consistency-qa/scripts/audit_pptx.py`               | `.claude/skills/consistency-qa/scripts/audit_pptx.py`               |
| `skills/consistency-qa/scripts/check_contract.py`           | `.claude/skills/consistency-qa/scripts/check_contract.py`           |
| `skills/deck-compose/SKILL.md`                              | `.claude/skills/deck-compose/SKILL.md`                              |
| `skills/deck-compose/assets/standard-deck-spec.json`        | `.claude/skills/deck-compose/assets/standard-deck-spec.json`        |
| `skills/deck-compose/references/golden-content-contract.md` | `.claude/skills/deck-compose/references/golden-content-contract.md` |
| `skills/deck-compose/references/layout-matching.md`         | `.claude/skills/deck-compose/references/layout-matching.md`         |
| `skills/deck-outline-grill/SKILL.md`                        | `.claude/skills/deck-outline-grill/SKILL.md`                        |
| `skills/pptmaker/SKILL.md`                                  | `.claude/skills/pptmaker/SKILL.md`                                  |
| `skills/pptx-build/SKILL.md`                                | `.claude/skills/pptx-build/SKILL.md`                                |
| `skills/pptx-build/assets/brand-kit.yaml`                   | `.claude/skills/pptx-build/assets/brand-kit.yaml`                   |
| `skills/pptx-build/assets/golden-snapshot.json`             | `.claude/skills/pptx-build/assets/golden-snapshot.json`             |
| `skills/pptx-build/references/deck-spec-schema.md`          | `.claude/skills/pptx-build/references/deck-spec-schema.md`          |
| `skills/pptx-build/references/design-rules.md`              | `.claude/skills/pptx-build/references/design-rules.md`              |
| `skills/pptx-build/references/reference-metrics.md`         | `.claude/skills/pptx-build/references/reference-metrics.md`         |
| `skills/pptx-build/scripts/audit_deck.py`                   | `.claude/skills/pptx-build/scripts/audit_deck.py`                   |
| `skills/pptx-build/scripts/audit_golden.py`                 | `.claude/skills/pptx-build/scripts/audit_golden.py`                 |
| `skills/pptx-build/scripts/build_pptx.py`                   | `.claude/skills/pptx-build/scripts/build_pptx.py`                   |
| `skills/pptx-build/scripts/compare_golden.py`               | `.claude/skills/pptx-build/scripts/compare_golden.py`               |
| `skills/pptx-build/scripts/goldenfab/__init__.py`           | `.claude/skills/pptx-build/scripts/goldenfab/__init__.py`           |
| `skills/pptx-build/scripts/goldenfab/_variant_h.py`         | `.claude/skills/pptx-build/scripts/goldenfab/_variant_h.py`         |
| `skills/pptx-build/scripts/goldenfab/_variant_k.py`         | `.claude/skills/pptx-build/scripts/goldenfab/_variant_k.py`         |
| `skills/pptx-build/scripts/goldenfab/audit.py`              | `.claude/skills/pptx-build/scripts/goldenfab/audit.py`              |
| `skills/pptx-build/scripts/goldenfab/common.py`             | `.claude/skills/pptx-build/scripts/goldenfab/common.py`             |
| `skills/pptx-build/scripts/goldenfab/content_contract.py`   | `.claude/skills/pptx-build/scripts/goldenfab/content_contract.py`   |
| `skills/pptx-build/scripts/goldenfab/grid.py`               | `.claude/skills/pptx-build/scripts/goldenfab/grid.py`               |
| `skills/pptx-build/scripts/goldenfab/kit.py`                | `.claude/skills/pptx-build/scripts/goldenfab/kit.py`                |
| `skills/pptx-build/scripts/goldenfab/layouts.py`            | `.claude/skills/pptx-build/scripts/goldenfab/layouts.py`            |
| `skills/pptx-build/scripts/goldenfab/reference.py`          | `.claude/skills/pptx-build/scripts/goldenfab/reference.py`          |
| `skills/pptx-build/scripts/goldenfab/registry.py`           | `.claude/skills/pptx-build/scripts/goldenfab/registry.py`           |
| `skills/pptx-build/scripts/goldenfab/s06_pilot.py`          | `.claude/skills/pptx-build/scripts/goldenfab/s06_pilot.py`          |
| `skills/pptx-build/scripts/goldenfab/s06_proto_f.py`        | `.claude/skills/pptx-build/scripts/goldenfab/s06_proto_f.py`        |
| `skills/pptx-build/scripts/goldenfab/s06_variants.py`       | `.claude/skills/pptx-build/scripts/goldenfab/s06_variants.py`       |
| `skills/pptx-build/scripts/goldenfab/s08_variants.py`       | `.claude/skills/pptx-build/scripts/goldenfab/s08_variants.py`       |
| `skills/pptx-build/scripts/goldenfab/s09_variants.py`       | `.claude/skills/pptx-build/scripts/goldenfab/s09_variants.py`       |
| `skills/pptx-build/scripts/goldenfab/s10_screenshot.py`     | `.claude/skills/pptx-build/scripts/goldenfab/s10_screenshot.py`     |
| `skills/pptx-build/scripts/goldenfab/s11_branch_snap.py`    | `.claude/skills/pptx-build/scripts/goldenfab/s11_branch_snap.py`    |
| `skills/pptx-build/scripts/goldenfab/s11_variants.py`       | `.claude/skills/pptx-build/scripts/goldenfab/s11_variants.py`       |
| `skills/pptx-build/scripts/goldenfab/s12_variants.py`       | `.claude/skills/pptx-build/scripts/goldenfab/s12_variants.py`       |
| `skills/pptx-build/scripts/goldenfab/s14_variants.py`       | `.claude/skills/pptx-build/scripts/goldenfab/s14_variants.py`       |
| `skills/pptx-build/scripts/goldenfab/s15_variants.py`       | `.claude/skills/pptx-build/scripts/goldenfab/s15_variants.py`       |
| `skills/pptx-build/scripts/goldenfab/s17_variants.py`       | `.claude/skills/pptx-build/scripts/goldenfab/s17_variants.py`       |
| `skills/pptx-build/scripts/goldenfab/s18_variants.py`       | `.claude/skills/pptx-build/scripts/goldenfab/s18_variants.py`       |
| `skills/pptx-build/scripts/goldenfab/s21_closing.py`        | `.claude/skills/pptx-build/scripts/goldenfab/s21_closing.py`        |
| `skills/pptx-build/scripts/goldenfab/shape_kind.py`         | `.claude/skills/pptx-build/scripts/goldenfab/shape_kind.py`         |
| `skills/pptx-build/scripts/render_deck.ps1`                 | `.claude/skills/pptx-build/scripts/render_deck.ps1`                 |
| `skills/pptx-build/scripts/verify_content_contract.py`      | `.claude/skills/pptx-build/scripts/verify_content_contract.py`      |
| `skills/pptx-visuals/SKILL.md`                              | `.claude/skills/pptx-visuals/SKILL.md`                              |
| `skills/pptx-visuals/references/archetype-catalog.md`       | `.claude/skills/pptx-visuals/references/archetype-catalog.md`       |
| `skills/pptx-visuals/references/visual-selection.md`        | `.claude/skills/pptx-visuals/references/visual-selection.md`        |
| `skills/pptx-visuals/scripts/check_candidates.py`           | `.claude/skills/pptx-visuals/scripts/check_candidates.py`           |
| `skills/pptx-visuals/scripts/make_icons.py`                 | `.claude/skills/pptx-visuals/scripts/make_icons.py`                 |
| `skills/pptx-visuals/scripts/make_mockups.py`               | `.claude/skills/pptx-visuals/scripts/make_mockups.py`               |
| `skills/pptx-visuals/scripts/mpl_exhibits.py`               | `.claude/skills/pptx-visuals/scripts/mpl_exhibits.py`               |
| `skills/pptx-visuals/scripts/recommend_archetypes.py`       | `.claude/skills/pptx-visuals/scripts/recommend_archetypes.py`       |
| `skills/pptx-visuals/scripts/render_real_mockups.py`        | `.claude/skills/pptx-visuals/scripts/render_real_mockups.py`        |
| `skills/pptx-visuals/scripts/visuals.py`                    | `.claude/skills/pptx-visuals/scripts/visuals.py`                    |

## 목록만(정독 면제) 파일 — 475건

바이너리·렌더 산출물·잠금 파일은 정독 대상이 아니다(census 분류 기준). 구조·대표 표본은 해당 장에서 다뤘다.

| 위치                                            | 건수 | 내용                                              | 대표 표본 다룬 장    |
| ----------------------------------------------- | ---- | ------------------------------------------------- | -------------------- |
| `golden/variants/`                              | 122  | 시안 렌더 PNG·pptx                                | 7 (변형 채택사)      |
| `_workspace_kifrs/`                             | 100  | 목업 갤러리 PNG 등                                | 9 (파일럿)           |
| `.claude/skills/` 하위                          | 99   | Lucide 아이콘 SVG 24종×3색·레퍼런스 PNG·견본 pptx | 6 (시각 어휘)        |
| `golden/ref·render·render_v6·_pilot_s06·backup` | 52   | 레퍼런스·렌더 스냅샷                              | 7                    |
| `data/`                                         | 10   | BCG·Dallas 등 벤치마크 PDF·pptx 원본              | 6·12                 |
| `results/`                                      | 4    | 빌드 실물 pptx(25·20·16·16장) — 게이트 미통과분   | 9 (서두에 전건 계측) |
| 기타(uv.lock·.python-version 등)                | 잔여 | 잠금·버전 고정                                    | 1                    |

확장자 분포(실측): PNG 400 · pptx 39 · SVG 24 · PDF 10 · 기타.

## verify 실행 결과

2026-07-19 실행 실측. 두 실행 모두 exit 0 — 차집합 0건.

```
$ python census.py verify . FINAL-REPORT/14_COVERAGE.md --exclude _workspace_prev
인벤토리 N = 48, 커버리지 표 경로 = 185 + 글롭 0
PASS — 전수 커버리지 48/48
exit 0

$ python census.py verify .claude FINAL-REPORT/14_COVERAGE.md --exclude _workspace_prev --exclude state
인벤토리 N = 60, 커버리지 표 경로 = 185 + 글롭 0
PASS — 전수 커버리지 60/60
exit 0
```

합산 **108/108**. 글롭 0건 — 디렉토리 단위 뭉뚱그림 없이 파일별로 매핑했다(전부 매칭 글롭에 의한 hollow-PASS 여지 없음).

## 이 부록이 증명하지 않는 것

커버리지 PASS는 "N개 파일을 전부 읽고 장에 대응시켰다"의 증명일 뿐, **"쓴 내용이 맞다"의 증명은 아니다**.
후자는 별도의 주장 감사 패스(집필에 관여하지 않은 감사자가 각 장의 현재형 동작 서술·다이어그램 엣지·수치를
코드와 대조)가 담당하며, 그 결과는 `11_TEST-DECISIONS.md`의 검증 로그에 기록된다.
