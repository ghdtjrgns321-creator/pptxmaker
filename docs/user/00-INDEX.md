# pptmaker 사용자 문서 — 목차

pptmaker는 내 프로젝트를 소개하는 B2B 셀링 PPT를 일관 서식으로 찍어내는 서식 공장이다.
이 폴더는 "지금까지 무슨 일이 있었고, 왜 방향을 바꾸며, 지금 무엇이 구현돼 어떻게 도는가"를
사용자 관점에서 정리한 문서 묶음이다. 01→08이 여정 순서다 — 과거(01~03) → 현재 파이프라인(04)
→ 골든 덱 실행(05~07) → 시스템 전체 정리(08).

## 문서 목록 (8건)

| 순서 | 문서                                        | 내용                                                                        | 상태             |
| ---- | ------------------------------------------- | --------------------------------------------------------------------------- | ---------------- |
| 01   | [여정](01-journey.md)                       | v1 → v4까지 4번의 개편이 각각 무엇을 고치려 했고 무엇이 남았는가            | 기록(과거)       |
| 02   | [실패 분석](02-failure-analysis.md)         | 왜 결과물이 "없어 보이는가" — McKinsey 실측 대조 갭 7개와 근본 원인         | 기록(진단)       |
| 03   | [골든 덱 전환](03-golden-deck.md)           | 규칙 쌓기를 멈추고 골든 덱 기준으로 가기로 한 결정과 진행 방식              | 결정 확정        |
| 04   | [전체 파이프라인](04-pipeline.md)           | 현재 파이프라인의 데이터 흐름과 단계별 역할 (현재형)                        | 현행             |
| 05   | [골든 덱 실행 계획](05-golden-deck-plan.md) | grilling으로 확정한 결정 9건, 견본 덱 구성, Phase 0~4 실행 단계             | Phase 2까지 완료 |
| 06   | [목차 공동설계 게이트](06-outline-grill.md) | deck-outline-grill — 빌드 전 목차·장별 역할을 사용자와 공동 확정하는 ①.5    | 구현 완료        |
| 07   | [공장 이식 진행표](07-factory-port.md)      | Phase 2 — 골든 19장 → goldenfab 15타입 이식, 도형 501/501 일치 검증         | 완료             |
| 08   | [시스템 전체 정리](08-system-overview.md)   | 무엇을 만들었고 지금 어떻게 도는가 — 정본 팩트시트·구성요소·품질 게이트 4겹 | 현행(정본)       |

## 현재 상태 요약 (2026-07-11)

- **Phase 0~2 완료:** 골든 덱 19장(501도형) 사용자 승인 확정 → 공장(goldenfab 15타입) 이식 →
  compare_golden 회귀 게이트로 골든↔공장 도형 501/501 일치 검증.
- **다음:** Phase 3(K-IFRS 덱을 공장으로 end-to-end 재빌드, 갭 7개 재점검) → Phase 4(v4 잔재 정리).
- 파이프라인 골격(추출 → 목차 인터뷰 → 구성 → 빌드 → 검수)은 유지. 시각 품질의 수렴점이
  "게이트 통과"에서 "골든 덱과의 일치"로 바뀌었다.

## 구현 자산 인벤토리 (현재 저장소 실측)

| 분류            | 자산                                                                                                                        | 규모              | 상세 문서                    |
| --------------- | --------------------------------------------------------------------------------------------------------------------------- | ----------------- | ---------------------------- |
| 스킬 7종        | pptmaker(오케스트레이터) · content-extract · deck-outline-grill · deck-compose · pptx-build · pptx-visuals · consistency-qa | `.claude/skills/` | [04](04-pipeline.md)         |
| 에이전트 4종    | content-extractor · deck-composer · pptx-builder · consistency-qa                                                           | `.claude/agents/` | [04](04-pipeline.md)         |
| 골든 덱(정답지) | `golden/` 스크립트 22파일 → golden-deck.pptx 19장·501도형, 시안 렌더 누적 106장                                             | 동결(기준 원본)   | [05](05-golden-deck-plan.md) |
| 디자인 규칙     | `pptx-build/references/design-rules.md` — Ⅰ전역·Ⅱ장유형 A~J·Ⅲ프로세스 P1~P4                                                 | 96항목            | [08](08-system-overview.md)  |
| 공장 엔진       | `pptx-build/scripts/goldenfab/` 20파일, 레이아웃 타입 15종 registry + build_pptx.py `golden.*` 디스패치                     | 15타입            | [07](07-factory-port.md)     |
| 품질 게이트     | compare_golden.py(도형 전수 회귀) · audit_pptx.py(마감 오딧) · 물성 매칭 결정표(layout-matching.md 15행)                    | 4겹 게이트        | [08](08-system-overview.md)  |

## 더 깊은 문서 (개발자용)

- 파이프라인 설계 원본: [../PIPELINE.md](../PIPELINE.md)
- 프로젝트 운영 규칙: [../../CLAUDE.md](../../CLAUDE.md)
- 골든 콘텐츠 팩트시트: [../../golden/00_factsheet.md](../../golden/00_factsheet.md)
