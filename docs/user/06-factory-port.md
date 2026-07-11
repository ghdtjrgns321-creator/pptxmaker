# Phase 2 — 공장 역설계 진행표

골든 덱(19장, 동결 기준 원본)을 공장 레이아웃 엔진(`pptx-build/scripts/goldenfab/`)으로 이식한다.
검증은 `scripts/compare_golden.py` — 골든 원본과 도형 전수 비교(±0.005", 텍스트·채움색·폰트 포함),
불일치 0이어야 통과. 이식 후 수정은 goldenfab에서만 한다(05-plan 결정 #9).

## 아키타입 인벤토리 — 골든 19장 → 공장 레이아웃 타입 매핑

| 장  | 골든 슬라이드          | 골든 소스(동결)              | 공장 레이아웃 타입 | 이식 상태             |
| --- | ---------------------- | ---------------------------- | ------------------ | --------------------- |
| 1   | 표지                   | build_golden.build_s01_cover | `cover`            | 완료 (8/8 일치)       |
| 2   | 목차                   | build_golden.build_s02_toc   | `toc`              | 완료 (36/36 일치)     |
| 3   | 간지 Ⅰ                 | build_golden.build_part      | `part`             | 완료 (12/12 ×2케이스) |
| 4   | 문제 정의(다구획)      | _variant_k.variant_k         | `problem_grid`     | 완료 (42/42) |
| 5   | 간지 Ⅱ                 | (part 반복)                  | `part`             | 완료                  |
| 6   | 파이프라인 실행 그래프 | s06_variants.variant_c       | `exec_graph`       | 완료 (43/43) |
| 7   | 간지 Ⅲ                 | (part 반복)                  | `part`             | 완료                  |
| 8   | 기술1 용어사전         | s08_variants.variant_c       | `tech_evidence`    | 완료 (20/20) |
| 9   | 기술2 지식그래프       | s09_variants.variant_a       | `tech_tree`        | 완료 (39/39) |
| 10  | 그래프 실물 화면       | s10_screenshot.variant_a     | `screenshot`       | 완료 (21/21·파라미터) |
| 11  | 기술3 판단트리         | s11_variants.variant_d       | `tech_mechanism`   | 완료 (46/46) |
| 12  | 기술4 구조화 출력      | s12_variants.variant_b       | `tech_capture`     | 완료 (33/33) |
| 13  | 간지 Ⅳ                 | (part 반복)                  | `part`             | 완료                  |
| 14  | 트러블슈팅 A/B         | s14_variants.variant_c       | `ab_simulation`    | 완료 (37/37) |
| 15  | 골든테스트 결과        | s15_variants.variant_c       | `validation`       | 완료 (46/46) |
| 16  | 간지 Ⅴ                 | (part 반복)                  | `part`             | 완료                  |
| 17  | 7축 미러 매트릭스      | s17_variants.variant_c       | `mirror_matrix`    | 완료 (29/29) |
| 18  | 정직한 한계(경계)      | s18_variants.variant_b       | `boundary`         | 완료 (26/26) |
| 19  | 클로징                 | s21_closing.variant_a        | `closing`          | 완료 (15/15·파라미터) |

고유 레이아웃 타입 **15종**(part 1종이 6장 커버 — 초판의 13종은 오산, registry 실측 15).
슬라이스 2 완료 — **15/15 타입 이식**, 전 19장 도형 501/501 일치·불일치 0(compare_slice2.md).
콘텐츠 파라미터화 수준: cover·toc·part·closing·screenshot = dict 인자 / 나머지 10종 = 골든 콘텐츠
내장(심층 파라미터화는 Phase 3 실전 수요 발생 시 pull — YAGNI, 결정 #8).

## 이식 규약

- 좌표·서식 상수는 골든 소스에서 그대로 복사(임의 변경 금지) — 콘텐츠만 dict 인자로.
- 콘텐츠 dict의 스키마는 이식 시 타입별로 layouts.py 독스트링에 기록한다.
- 골든에 하드코딩된 개수 조건(예: 목차 헤어라인 `i < 5`)은 개수 파라미터로 일반화하되,
  골든과 같은 입력에서 같은 출력이 나옴을 compare_golden으로 증명한다.
- compare_golden.py는 이식 타입이 늘 때마다 케이스를 추가하고, 콘텐츠 파라미터가 반영되는지
  상이 입력 2케이스 이상으로 확인한다(ripple).

## 대조 리포트

- 슬라이스 1: 도형 68/68 일치, 불일치 0 (S1 8 · S2 36 · 간지 12×2) — compare_slice1.md
- 슬라이스 2: 전 19장 도형 501/501 일치, 불일치 0 + 파라미터 유효성 2/2(closing·screenshot 상이 입력 반영) — compare_slice2.md
