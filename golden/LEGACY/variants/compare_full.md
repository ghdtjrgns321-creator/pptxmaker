# compare_golden — 골든 레퍼런스 덱 vs 기준선 스냅샷 대조

허용오차 ±0.005" · 비교 속성: 수/type/text/fill/run(pt·bold)/run_colors/cells/l·t·w·h
기준선: `.claude\skills\pptx-build\assets\golden-snapshot.json` · 소스: `goldenfab/reference.py`

| 슬라이드 | 대상 | 내용 | 판정 |
|---|---|---|---|
| S1 표지 | 전체 | 8/8 일치 | PASS |
| S2 목차 | 전체 | 31/31 일치 | PASS |
| S3 간지1 | 전체 | 11/11 일치 | PASS |
| S4 문제정의 | 전체 | 48/48 일치 | PASS |
| S5 간지2 | 전체 | 11/11 일치 | PASS |
| S6 실행그래프 | 전체 | 43/43 일치 | PASS |
| S7 간지3 | 전체 | 11/11 일치 | PASS |
| S8 용어사전 | 전체 | 42/42 일치 | PASS |
| S9 지식그래프 | 전체 | 80/80 일치 | PASS |
| S10 스크린샷 | 전체 | 21/21 일치 | PASS |
| S11 판단트리 | 전체 | 46/46 일치 | PASS |
| S12 구조화출력 | 전체 | 33/33 일치 | PASS |
| S13 간지4 | 전체 | 11/11 일치 | PASS |
| S14 트러블슈팅 | 전체 | 37/37 일치 | PASS |
| S15 골든테스트 | 전체 | 46/46 일치 | PASS |
| S16 경계 | 전체 | 26/26 일치 | PASS |
| S17 클로징 | 전체 | 14/14 일치 | PASS |

총계: 슬라이드 17/17, 도형 일치 519/519, 불일치 0건

## 파라미터 유효성 (상이 입력 반영)
- closing 커스텀: PASS
- screenshot 커스텀: PASS
- problem_grid override: PASS
- exec_graph override: PASS
- tech_evidence override: PASS
- tech_tree override: PASS
- tech_mechanism override: PASS
- tech_capture override: PASS
- ab_simulation override: PASS
- validation override: PASS
- mirror_matrix override: PASS
- boundary override: PASS
