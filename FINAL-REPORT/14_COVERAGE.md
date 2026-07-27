# 14. 전수 커버리지 부록 (COVERAGE)

> 이 장은 "다 읽었다"의 기계 증명이다. 분모 N은 손이 아니라 `census.py inventory`가 고정했고,
> 아래 매핑 표가 N개 파일 전부를 반영 장에 대응시킨다. 말미의 verify PASS 출력이 차집합 0의 증거다.

## 분모 N = 165 (루트 82 + .claude 서브트리 83)

census 기본 제외 상수(EXCLUDE_DIRS)에 `.claude`가 포함되어 있으나, 이 프로젝트는 본체 코드가
`.claude/skills/`에 있다. 따라서 인벤토리를 2회 실행해 분모를 합산했다 — 스크립트 무수정, 재현 가능.

실행 커맨드(재현용):

```
uv run python ~/.claude/skills/final-report/scripts/census.py inventory . --exclude _workspace_prev
uv run python ~/.claude/skills/final-report/scripts/census.py inventory .claude --exclude state --exclude _workspace_prev
```

| 실행    | 정독 분모 | kind 내역                            | 목록만(정독 면제) |
| ------- | --------- | ------------------------------------ | ----------------- |
| 루트    | 82        | code 32 · config 9 · doc 39 · text 2 | 442               |
| .claude | 83        | code 60 · config 5 · doc 18          | 99                |

## 제외 내역 (조용한 제외 금지 — 전량 기록)

| 디렉토리               | 걸러진 파일 | 구분                                 | 사유                                                              |
| ---------------------- | ----------- | ------------------------------------ | ----------------------------------------------------------------- |
| `.git`                 | 1,492       | 기본 제외                            | VCS 내부                                                          |
| `.venv`                | 3,092       | 기본 제외                            | uv 가상환경(재현 가능 산출물)                                     |
| `.ruff_cache`          | 46          | 기본 제외                            | 린터 캐시                                                         |
| `__pycache__`          | 25 + 61     | 기본 제외                            | 바이트코드 캐시                                                   |
| `FINAL-REPORT`         | 15          | 기본 제외                            | 이 보고서 자신                                                    |
| `.claude`              | 296         | 기본 제외 → **별도 인벤토리로 회수** | 본체 코드 소재 — 2차 실행(N=83)으로 분모 편입                     |
| `_workspace_prev`      | 875 + 3     | `--exclude` 명시                     | 폐기된 과거 스킬 작업 스크래치 — 여정 사료는 docs·git 이력이 대체 |
| `state` (.claude 하위) | 48          | `--exclude` 명시                     | 하네스 런타임 상태 — gitignore 대상, 코드가 아님                  |
| `_workspace`           | 1           | 기본 제외                            | 구 워크스페이스 잔재                                              |

## 파일 → 장 매핑 (165/165)

| 파일                                                                       | kind   | 반영 장 |
| -------------------------------------------------------------------------- | ------ | ------- |
| `.gitignore`                                                               | config | 10      |
| `_workspace_kifrs/01.5_outline.md`                                         | doc    | 9       |
| `_workspace_kifrs/02_deck-spec.json`                                       | config | 9       |
| `_workspace_kifrs/03_exhibit-candidates.json`                              | config | 9       |
| `_workspace_kifrs/03_qa-report.md`                                         | doc    | 9       |
| `_workspace_kifrs/04_contract-check.md`                                    | doc    | 9       |
| `_workspace_kifrs/05_ripple-golden.md`                                     | doc    | 9       |
| `_workspace_kifrs/06_s4-color-inventory.md`                                | doc    | 9       |
| `_workspace_kifrs/07-compare-before.txt`                                   | text   | 9       |
| `_workspace_kifrs/07-s8-s9-color.md`                                       | doc    | 9       |
| `_workspace_kifrs/07-s8-s9-probe.txt`                                      | text   | 9       |
| `_workspace_kifrs/07-s8-s9-selftest.md`                                    | doc    | 9       |
| `_workspace_kifrs/07-s8-s9-verify.md`                                      | doc    | 9       |
| `_workspace_kifrs/dump_contract.py`                                        | code   | 9       |
| `_workspace_kifrs/golden_defaults.json`                                    | config | 9       |
| `_workspace_kifrs/HANDOFF.md`                                              | doc    | 9       |
| `_workspace_kifrs/mockups/gallery.html`                                    | config | 9       |
| `_workspace_kifrs/param-audit.md`                                          | doc    | 9       |
| `_workspace_kifrs/verify_s8s9_data.py`                                     | code   | 9       |
| `CLAUDE.md`                                                                | doc    | 1·12    |
| `docs/2026-07-20-dense-failure-analysis.md`                                | doc    | 11·12   |
| `docs/2026-07-24-deck-smith-inheritance-experiment.md`                     | doc    | 11·12   |
| `docs/2026-07-25-골든-도해-어휘-인벤토리.md`                               | doc    | 11·12   |
| `docs/2026-07-25-어휘-정본-이전-계획.md`                                   | doc    | 11·12   |
| `docs/2026-07-26-부품-도구함-재설계.md`                                    | doc    | 11·12   |
| `docs/2026-07-26-부품-선택-모델-핸드오프.md`                               | doc    | 11·12   |
| `docs/2026-07-26-부품-선택-모델-확정.md`                                   | doc    | 11·12   |
| `docs/2026-07-27-선택규칙-실전재료-시험.md`                                | doc    | 11·12   |
| `docs/2026-07-27-핸드오프-덱-완주.md`                                      | doc    | 11·12   |
| `docs/CHANGELOG.md`                                                        | doc    | 12      |
| `docs/handoff-2026-07-20-decision-table-catalog.md`                        | doc    | 11·12   |
| `docs/handoff-2026-07-20-golden-dense.md`                                  | doc    | 11·12   |
| `docs/handoff-2026-07-24-audit-and-inheritance.md`                         | doc    | 11·12   |
| `docs/PIPELINE.md`                                                         | doc    | 2       |
| `docs/superpowers/specs/2026-07-08-visual-diversity-design.md`             | doc    | 11·12   |
| `docs/superpowers/specs/2026-07-15-gate-repair-design.md`                  | doc    | 11·12   |
| `docs/superpowers/specs/2026-07-16-golden-adaptive-wiring-design.md`       | doc    | 11·12   |
| `golden/_pilot_s06/ref_anatomy.md`                                         | doc    | 7·13    |
| `golden/_pilot_s06/units.md`                                               | doc    | 7·13    |
| `golden/build_golden.py`                                                   | code   | 7       |
| `golden/LEGACY/00_factsheet.md`                                            | doc    | 12      |
| `golden/LEGACY/00_phase0-crosscheck.md`                                    | doc    | 12      |
| `golden/LEGACY/build_golden.py`                                            | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/golden-snapshot.json`              | config | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/__init__.py`         | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/_variant_h.py`       | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/_variant_k.py`       | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/audit.py`            | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/common.py`           | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/content_contract.py` | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/grid.py`             | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/kit.py`              | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/layouts.py`          | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/reference.py`        | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/registry.py`         | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s06_mid.py`          | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s06_pilot.py`        | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s06_proto_f.py`      | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s06_variants.py`     | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s08_variants.py`     | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s09_variants.py`     | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s10_screenshot.py`   | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s11_branch_snap.py`  | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s11_variants.py`     | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s12_variants.py`     | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s14_variants.py`     | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s15_variants.py`     | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s17_variants.py`     | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s18_variants.py`     | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/s21_closing.py`      | code   | 12      |
| `golden/LEGACY/g/golden-deck-2026-07-20/goldenfab-src/shape_kind.py`       | code   | 12      |
| `golden/LEGACY/g/s06_pilot_v3.py`                                          | code   | 12      |
| `golden/LEGACY/g/v5-monochrome-2026-07-18/brand-kit.yaml`                  | config | 12      |
| `golden/LEGACY/variants/compare_full.md`                                   | doc    | 12      |
| `golden/LEGACY/variants/desktop.ini`                                       | config | 12      |
| `golden/variants/compare_full.md`                                          | doc    | 8       |
| `input/k-ifrs-1115/claude-report-v2.md`                                    | doc    | 9       |
| `input/k-ifrs-1115/notebooklm-report-1.md`                                 | doc    | 9       |
| `input/k-ifrs-1115/notebooklm-report-v2.md`                                | doc    | 9       |
| `pyproject.toml`                                                           | config | 1       |
| `README.md`                                                                | doc    | —       |
| `ref/catalog.md`                                                           | doc    | 6       |
| `.claude/agents/consistency-qa.md`                                         | doc    | 2       |
| `.claude/agents/deck-composer.md`                                          | doc    | 2       |
| `.claude/agents/deck-smith.md`                                             | doc    | 2       |
| `.claude/agents/pptx-builder.md`                                           | doc    | 2       |
| `.claude/hooks/pptx_arm.py`                                                | code   | 8       |
| `.claude/hooks/pptx_render_gate.sh`                                        | code   | 8       |
| `.claude/settings.json`                                                    | config | 8       |
| `.claude/skills/consistency-qa/scripts/audit_pptx.py`                      | code   | 8       |
| `.claude/skills/consistency-qa/scripts/check_contract.py`                  | code   | 8       |
| `.claude/skills/consistency-qa/SKILL.md`                                   | doc    | 8       |
| `.claude/skills/deck-compose/references/golden-content-contract.md`        | doc    | 4       |
| `.claude/skills/deck-compose/references/layout-matching.md`                | doc    | 4       |
| `.claude/skills/deck-compose/SKILL.md`                                     | doc    | 4       |
| `.claude/skills/deck-outline-grill/SKILL.md`                               | doc    | 3       |
| `.claude/skills/part-design/SKILL.md`                                      | doc    | 6       |
| `.claude/skills/pptmaker/SKILL.md`                                         | doc    | 2       |
| `.claude/skills/pptx-build/assets/brand-kit.yaml`                          | config | 5       |
| `.claude/skills/pptx-build/assets/golden-snapshot.json`                    | config | 7·8     |
| `.claude/skills/pptx-build/assets/legacy-refs.json`                        | config | 8       |
| `.claude/skills/pptx-build/assets/sample-material-pptmaker.json`           | config | 4       |
| `.claude/skills/pptx-build/references/deck-spec-schema.md`                 | doc    | 5       |
| `.claude/skills/pptx-build/references/design-rules.md`                     | doc    | 5       |
| `.claude/skills/pptx-build/references/reference-metrics.md`                | doc    | 5       |
| `.claude/skills/pptx-build/scripts/audit_deck.py`                          | code   | 8       |
| `.claude/skills/pptx-build/scripts/audit_golden.py`                        | code   | 8       |
| `.claude/skills/pptx-build/scripts/build_pptx.py`                          | code   | 5       |
| `.claude/skills/pptx-build/scripts/compare_golden.py`                      | code   | 8       |
| `.claude/skills/pptx-build/scripts/fig_mockup.py`                          | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/__init__.py`                  | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/audit.py`                     | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/common.py`                    | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/content_contract.py`          | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/dense.py`                     | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/__init__.py`          | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/bipartite_map.py`     | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/card_row.py`          | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/elements.py`          | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/fan_in.py`            | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/gate_branch.py`       | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/hub_spoke.py`         | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/layer_stack.py`       | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/n_branch.py`          | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/numbered_steps.py`    | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/relation_catalog.py`  | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/figures/routing_lane.py`      | code   | 6       |
| `.claude/skills/pptx-build/scripts/goldenfab/frames.py`                    | code   | 4       |
| `.claude/skills/pptx-build/scripts/goldenfab/grid.py`                      | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/kit.py`                       | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/layouts.py`                   | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/reference.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/registry.py`                  | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s04_dense.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s06_mid.py`                   | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s06_variants.py`              | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s08_dense.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s09_dense.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s10_dense.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s11_dense.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s12_dense.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s14_dense.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s15_dense.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s16_dense.py`                 | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s17_variants.py`              | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/s21_closing.py`               | code   | 7       |
| `.claude/skills/pptx-build/scripts/goldenfab/select.py`                    | code   | 4       |
| `.claude/skills/pptx-build/scripts/pick_parts.py`                          | code   | 4       |
| `.claude/skills/pptx-build/scripts/preflight_dense.py`                     | code   | 8       |
| `.claude/skills/pptx-build/scripts/render_deck.ps1`                        | code   | 5       |
| `.claude/skills/pptx-build/scripts/scan_material.py`                       | code   | 3       |
| `.claude/skills/pptx-build/scripts/test_wiring.py`                         | code   | 8       |
| `.claude/skills/pptx-build/scripts/verify_content_contract.py`             | code   | 8       |
| `.claude/skills/pptx-build/scripts/verify_selection.py`                    | code   | 4       |
| `.claude/skills/pptx-build/SKILL.md`                                       | doc    | 5       |
| `.claude/skills/pptx-visuals/references/archetype-catalog.md`              | doc    | 6       |
| `.claude/skills/pptx-visuals/references/visual-selection.md`               | doc    | 6       |
| `.claude/skills/pptx-visuals/scripts/check_candidates.py`                  | code   | 6       |
| `.claude/skills/pptx-visuals/scripts/make_icons.py`                        | code   | 6       |
| `.claude/skills/pptx-visuals/scripts/make_mockups.py`                      | code   | 6       |
| `.claude/skills/pptx-visuals/scripts/mpl_exhibits.py`                      | code   | 6       |
| `.claude/skills/pptx-visuals/scripts/recommend_archetypes.py`              | code   | 6       |
| `.claude/skills/pptx-visuals/scripts/render_real_mockups.py`               | code   | 6       |
| `.claude/skills/pptx-visuals/scripts/visuals.py`                           | code   | 6       |
| `.claude/skills/pptx-visuals/SKILL.md`                                     | doc    | 6       |

## .claude 서브트리 verify 대조 경로 (83건)

verify 2차 실행은 `.claude`를 루트로 삼아 상대경로로 대조한다. 본표의 `.claude/...` 표기와 동일 파일이다.

| verify 기준 경로                                                  | 본표 경로                                                                 |
| ----------------------------------------------------------------- | ------------------------------------------------------------------------- |
| `agents/consistency-qa.md`                                        | `.claude/agents/consistency-qa.md`                                        |
| `agents/deck-composer.md`                                         | `.claude/agents/deck-composer.md`                                         |
| `agents/deck-smith.md`                                            | `.claude/agents/deck-smith.md`                                            |
| `agents/pptx-builder.md`                                          | `.claude/agents/pptx-builder.md`                                          |
| `hooks/pptx_arm.py`                                               | `.claude/hooks/pptx_arm.py`                                               |
| `hooks/pptx_render_gate.sh`                                       | `.claude/hooks/pptx_render_gate.sh`                                       |
| `settings.json`                                                   | `.claude/settings.json`                                                   |
| `skills/consistency-qa/scripts/audit_pptx.py`                     | `.claude/skills/consistency-qa/scripts/audit_pptx.py`                     |
| `skills/consistency-qa/scripts/check_contract.py`                 | `.claude/skills/consistency-qa/scripts/check_contract.py`                 |
| `skills/consistency-qa/SKILL.md`                                  | `.claude/skills/consistency-qa/SKILL.md`                                  |
| `skills/deck-compose/references/golden-content-contract.md`       | `.claude/skills/deck-compose/references/golden-content-contract.md`       |
| `skills/deck-compose/references/layout-matching.md`               | `.claude/skills/deck-compose/references/layout-matching.md`               |
| `skills/deck-compose/SKILL.md`                                    | `.claude/skills/deck-compose/SKILL.md`                                    |
| `skills/deck-outline-grill/SKILL.md`                              | `.claude/skills/deck-outline-grill/SKILL.md`                              |
| `skills/part-design/SKILL.md`                                     | `.claude/skills/part-design/SKILL.md`                                     |
| `skills/pptmaker/SKILL.md`                                        | `.claude/skills/pptmaker/SKILL.md`                                        |
| `skills/pptx-build/assets/brand-kit.yaml`                         | `.claude/skills/pptx-build/assets/brand-kit.yaml`                         |
| `skills/pptx-build/assets/golden-snapshot.json`                   | `.claude/skills/pptx-build/assets/golden-snapshot.json`                   |
| `skills/pptx-build/assets/legacy-refs.json`                       | `.claude/skills/pptx-build/assets/legacy-refs.json`                       |
| `skills/pptx-build/assets/sample-material-pptmaker.json`          | `.claude/skills/pptx-build/assets/sample-material-pptmaker.json`          |
| `skills/pptx-build/references/deck-spec-schema.md`                | `.claude/skills/pptx-build/references/deck-spec-schema.md`                |
| `skills/pptx-build/references/design-rules.md`                    | `.claude/skills/pptx-build/references/design-rules.md`                    |
| `skills/pptx-build/references/reference-metrics.md`               | `.claude/skills/pptx-build/references/reference-metrics.md`               |
| `skills/pptx-build/scripts/audit_deck.py`                         | `.claude/skills/pptx-build/scripts/audit_deck.py`                         |
| `skills/pptx-build/scripts/audit_golden.py`                       | `.claude/skills/pptx-build/scripts/audit_golden.py`                       |
| `skills/pptx-build/scripts/build_pptx.py`                         | `.claude/skills/pptx-build/scripts/build_pptx.py`                         |
| `skills/pptx-build/scripts/compare_golden.py`                     | `.claude/skills/pptx-build/scripts/compare_golden.py`                     |
| `skills/pptx-build/scripts/fig_mockup.py`                         | `.claude/skills/pptx-build/scripts/fig_mockup.py`                         |
| `skills/pptx-build/scripts/goldenfab/__init__.py`                 | `.claude/skills/pptx-build/scripts/goldenfab/__init__.py`                 |
| `skills/pptx-build/scripts/goldenfab/audit.py`                    | `.claude/skills/pptx-build/scripts/goldenfab/audit.py`                    |
| `skills/pptx-build/scripts/goldenfab/common.py`                   | `.claude/skills/pptx-build/scripts/goldenfab/common.py`                   |
| `skills/pptx-build/scripts/goldenfab/content_contract.py`         | `.claude/skills/pptx-build/scripts/goldenfab/content_contract.py`         |
| `skills/pptx-build/scripts/goldenfab/dense.py`                    | `.claude/skills/pptx-build/scripts/goldenfab/dense.py`                    |
| `skills/pptx-build/scripts/goldenfab/figures/__init__.py`         | `.claude/skills/pptx-build/scripts/goldenfab/figures/__init__.py`         |
| `skills/pptx-build/scripts/goldenfab/figures/bipartite_map.py`    | `.claude/skills/pptx-build/scripts/goldenfab/figures/bipartite_map.py`    |
| `skills/pptx-build/scripts/goldenfab/figures/card_row.py`         | `.claude/skills/pptx-build/scripts/goldenfab/figures/card_row.py`         |
| `skills/pptx-build/scripts/goldenfab/figures/elements.py`         | `.claude/skills/pptx-build/scripts/goldenfab/figures/elements.py`         |
| `skills/pptx-build/scripts/goldenfab/figures/fan_in.py`           | `.claude/skills/pptx-build/scripts/goldenfab/figures/fan_in.py`           |
| `skills/pptx-build/scripts/goldenfab/figures/gate_branch.py`      | `.claude/skills/pptx-build/scripts/goldenfab/figures/gate_branch.py`      |
| `skills/pptx-build/scripts/goldenfab/figures/hub_spoke.py`        | `.claude/skills/pptx-build/scripts/goldenfab/figures/hub_spoke.py`        |
| `skills/pptx-build/scripts/goldenfab/figures/layer_stack.py`      | `.claude/skills/pptx-build/scripts/goldenfab/figures/layer_stack.py`      |
| `skills/pptx-build/scripts/goldenfab/figures/n_branch.py`         | `.claude/skills/pptx-build/scripts/goldenfab/figures/n_branch.py`         |
| `skills/pptx-build/scripts/goldenfab/figures/numbered_steps.py`   | `.claude/skills/pptx-build/scripts/goldenfab/figures/numbered_steps.py`   |
| `skills/pptx-build/scripts/goldenfab/figures/relation_catalog.py` | `.claude/skills/pptx-build/scripts/goldenfab/figures/relation_catalog.py` |
| `skills/pptx-build/scripts/goldenfab/figures/routing_lane.py`     | `.claude/skills/pptx-build/scripts/goldenfab/figures/routing_lane.py`     |
| `skills/pptx-build/scripts/goldenfab/frames.py`                   | `.claude/skills/pptx-build/scripts/goldenfab/frames.py`                   |
| `skills/pptx-build/scripts/goldenfab/grid.py`                     | `.claude/skills/pptx-build/scripts/goldenfab/grid.py`                     |
| `skills/pptx-build/scripts/goldenfab/kit.py`                      | `.claude/skills/pptx-build/scripts/goldenfab/kit.py`                      |
| `skills/pptx-build/scripts/goldenfab/layouts.py`                  | `.claude/skills/pptx-build/scripts/goldenfab/layouts.py`                  |
| `skills/pptx-build/scripts/goldenfab/reference.py`                | `.claude/skills/pptx-build/scripts/goldenfab/reference.py`                |
| `skills/pptx-build/scripts/goldenfab/registry.py`                 | `.claude/skills/pptx-build/scripts/goldenfab/registry.py`                 |
| `skills/pptx-build/scripts/goldenfab/s04_dense.py`                | `.claude/skills/pptx-build/scripts/goldenfab/s04_dense.py`                |
| `skills/pptx-build/scripts/goldenfab/s06_mid.py`                  | `.claude/skills/pptx-build/scripts/goldenfab/s06_mid.py`                  |
| `skills/pptx-build/scripts/goldenfab/s06_variants.py`             | `.claude/skills/pptx-build/scripts/goldenfab/s06_variants.py`             |
| `skills/pptx-build/scripts/goldenfab/s08_dense.py`                | `.claude/skills/pptx-build/scripts/goldenfab/s08_dense.py`                |
| `skills/pptx-build/scripts/goldenfab/s09_dense.py`                | `.claude/skills/pptx-build/scripts/goldenfab/s09_dense.py`                |
| `skills/pptx-build/scripts/goldenfab/s10_dense.py`                | `.claude/skills/pptx-build/scripts/goldenfab/s10_dense.py`                |
| `skills/pptx-build/scripts/goldenfab/s11_dense.py`                | `.claude/skills/pptx-build/scripts/goldenfab/s11_dense.py`                |
| `skills/pptx-build/scripts/goldenfab/s12_dense.py`                | `.claude/skills/pptx-build/scripts/goldenfab/s12_dense.py`                |
| `skills/pptx-build/scripts/goldenfab/s14_dense.py`                | `.claude/skills/pptx-build/scripts/goldenfab/s14_dense.py`                |
| `skills/pptx-build/scripts/goldenfab/s15_dense.py`                | `.claude/skills/pptx-build/scripts/goldenfab/s15_dense.py`                |
| `skills/pptx-build/scripts/goldenfab/s16_dense.py`                | `.claude/skills/pptx-build/scripts/goldenfab/s16_dense.py`                |
| `skills/pptx-build/scripts/goldenfab/s17_variants.py`             | `.claude/skills/pptx-build/scripts/goldenfab/s17_variants.py`             |
| `skills/pptx-build/scripts/goldenfab/s21_closing.py`              | `.claude/skills/pptx-build/scripts/goldenfab/s21_closing.py`              |
| `skills/pptx-build/scripts/goldenfab/select.py`                   | `.claude/skills/pptx-build/scripts/goldenfab/select.py`                   |
| `skills/pptx-build/scripts/pick_parts.py`                         | `.claude/skills/pptx-build/scripts/pick_parts.py`                         |
| `skills/pptx-build/scripts/preflight_dense.py`                    | `.claude/skills/pptx-build/scripts/preflight_dense.py`                    |
| `skills/pptx-build/scripts/render_deck.ps1`                       | `.claude/skills/pptx-build/scripts/render_deck.ps1`                       |
| `skills/pptx-build/scripts/scan_material.py`                      | `.claude/skills/pptx-build/scripts/scan_material.py`                      |
| `skills/pptx-build/scripts/test_wiring.py`                        | `.claude/skills/pptx-build/scripts/test_wiring.py`                        |
| `skills/pptx-build/scripts/verify_content_contract.py`            | `.claude/skills/pptx-build/scripts/verify_content_contract.py`            |
| `skills/pptx-build/scripts/verify_selection.py`                   | `.claude/skills/pptx-build/scripts/verify_selection.py`                   |
| `skills/pptx-build/SKILL.md`                                      | `.claude/skills/pptx-build/SKILL.md`                                      |
| `skills/pptx-visuals/references/archetype-catalog.md`             | `.claude/skills/pptx-visuals/references/archetype-catalog.md`             |
| `skills/pptx-visuals/references/visual-selection.md`              | `.claude/skills/pptx-visuals/references/visual-selection.md`              |
| `skills/pptx-visuals/scripts/check_candidates.py`                 | `.claude/skills/pptx-visuals/scripts/check_candidates.py`                 |
| `skills/pptx-visuals/scripts/make_icons.py`                       | `.claude/skills/pptx-visuals/scripts/make_icons.py`                       |
| `skills/pptx-visuals/scripts/make_mockups.py`                     | `.claude/skills/pptx-visuals/scripts/make_mockups.py`                     |
| `skills/pptx-visuals/scripts/mpl_exhibits.py`                     | `.claude/skills/pptx-visuals/scripts/mpl_exhibits.py`                     |
| `skills/pptx-visuals/scripts/recommend_archetypes.py`             | `.claude/skills/pptx-visuals/scripts/recommend_archetypes.py`             |
| `skills/pptx-visuals/scripts/render_real_mockups.py`              | `.claude/skills/pptx-visuals/scripts/render_real_mockups.py`              |
| `skills/pptx-visuals/scripts/visuals.py`                          | `.claude/skills/pptx-visuals/scripts/visuals.py`                          |
| `skills/pptx-visuals/SKILL.md`                                    | `.claude/skills/pptx-visuals/SKILL.md`                                    |
## 목록만(정독 면제) 파일 — 541건

바이너리·렌더 산출물·잠금 파일은 정독 대상이 아니다(census 분류 기준). 구조·대표 표본은 해당 장에서 다뤘다.

| 위치                               | 건수 | 내용                                    | 대표 표본 다룬 장 |
| ---------------------------------- | ---- | --------------------------------------- | ----------------- |
| `golden/`                          | 190  | 시안·LEGACY 렌더 PNG·pptx               | 7 · 12            |
| `.claude/skills/pptx-build/assets` | 96   | Lucide 아이콘 SVG 24종×3색 · 아이콘 PNG | 6 (시각 어휘)     |
| `_workspace_kifrs/`                | 100  | 목업 갤러리 PNG 등                      | 9 (파일럿)        |
| `ref/`                             | 75   | 레퍼런스 장표 PNG(BCG·McKinsey 등)      | 6 · 1             |
| `_render/`                         | 44   | 눈검증용 렌더 PNG(pptx에서 재생성)      | 8                 |
| `data/`                            | 13   | 벤치마크 PDF·pptx 원본                  | 1 · 6             |
| `results/`                         | 12   | 빌드 실물 pptx 3건 + 부품 목업 9건      | 9 · 6             |
| 기타(uv.lock·.python-version 등)   | 11   | 잠금·버전 고정·문서 이미지              | 1                 |

확장자 분포(실측): PNG 441 · pptx 63 · SVG 24 · PDF 10 · 기타 3.

## verify 실행 결과

2026-07-27 실행 실측. 두 실행 모두 exit 0 — 차집합 0건.

```
$ uv run python census.py verify . FINAL-REPORT/14_COVERAGE.md --exclude _workspace_prev
인벤토리 N = 82, 커버리지 표 경로 = 268 + 글롭 0
PASS — 전수 커버리지 82/82

$ uv run python census.py verify .claude FINAL-REPORT/14_COVERAGE.md --exclude state --exclude _workspace_prev
인벤토리 N = 83, 커버리지 표 경로 = 268 + 글롭 0
PASS — 전수 커버리지 83/83
```

합산 **165/165**. 글롭 0건 — 디렉토리 단위 뭉뚱그림 없이 파일별로 매핑했으므로
전부 매칭 글롭에 의한 hollow-PASS 여지가 없다.

## 이 부록이 증명하지 않는 것

커버리지 PASS는 "N개 파일을 전부 읽고 장에 대응시켰다"의 증명일 뿐, **"쓴 내용이 맞다"의 증명은 아니다**.
후자는 별도의 주장 감사 패스가 담당하며, 그 결과는 [11_TEST-DECISIONS.md](11_TEST-DECISIONS.md)의
검증 로그에 기록된다. 이 프로젝트가 시스템에 대해 말하는 것과 같은 구조다 —
**분모를 기계로 고정해도, 내용의 진위는 별도 게이트가 필요하다.**
