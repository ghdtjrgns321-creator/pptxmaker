# pptmaker — 골든 덱 기반 B2B 셀링 PPT 서식 공장

![Python](https://img.shields.io/badge/Python-3.13-3776AB)
![python-pptx](https://img.shields.io/badge/python--pptx-%E2%89%A51.0.2-15171B)
![PyYAML](https://img.shields.io/badge/PyYAML-%E2%89%A56.0.3-15171B)
![matplotlib](https://img.shields.io/badge/matplotlib-%E2%89%A53.11.0-15171B)
![uv](https://img.shields.io/badge/env-uv%20%C2%B7%20.venv%20%C2%B7%20uv.lock-5A29E4)
![golden deck](https://img.shields.io/badge/golden_deck-5_parts_%C2%B7_17_slides-D66E3A)
![snapshot](https://img.shields.io/badge/snapshot-519_shapes-D66E3A)
![regression](https://img.shields.io/badge/compare__golden-519%2F519_PASS-2E7D32)
![pilot QA](https://img.shields.io/badge/K--IFRS_pilot_QA-FAIL_17-C62828)
![status](https://img.shields.io/badge/status-WIP_%C2%B7_no_final_pptx-C62828)

> **현재 상태 배너 (2026-07-19)** — 이 프로젝트는 **미완성**이다.
> - **현행 계약·게이트를 통과한 최종 산출물이 없다.** 산출 목적지 `results/`에는 과거 빌드 4건이 남아 있으나(2026-07-08 25장·20장, 07-14·07-15 각 16장, 전부 git 미추적), 전부 현행 골든 17장 체계 이전이거나 QA FAIL 판정을 받은 산출물이다 — 완주 PASS 0건.
> - 유일한 실전 시도인 **K-IFRS 1115 파일럿은 2026-07-15 QA FAIL(계약 위반 17건) 판정 후 동결** 상태다. 이후 재빌드 기록이 없다. 상세: [9_KIFRS-PILOT.md](9_KIFRS-PILOT.md)
> - 골든 덱 고밀도 개편(2026-07-18 착수)이 **미커밋 워킹트리**로 진행 중이다(`brand-kit.yaml` 수정, s06 파일럿 실험 등).
>
> 위 배지의 초록(회귀 게이트 PASS)은 골든 덱 보호 게이트의 상태이고, 빨강(파일럿 FAIL·산출물 부재)은 실전 파이프라인의 상태다. 둘을 섞지 않는다.

## 한 줄 요약

사용자가 미리 정리해 둔 프로젝트별 `FINAL-REPORT/`(보고서형 md 묶음)를 재료로, **손으로 깎아 스냅샷(519도형)으로 박제한 골든 덱 17장의 서식·밀도를 기계 게이트(계약 대조·오딧·회귀)로 강제**하면서, 일관 브랜드의 **네이티브 .pptx** B2B 셀링 덱을 찍어내는 서식 공장 — 템플릿에 텍스트를 채우는 방식이 아니라, 게이트를 통과하지 못하면 빌드가 중단되는 방식이다.

## 시스템 한눈

```
 <프로젝트>/FINAL-REPORT/  (보고서형 md 묶음 — 사용자가 사전 정리)
        │
        ▼
 ① deck-outline-grill ────▶ 01.5_outline.md        ← 목차·채록 계약 (사용자 "확정")
        │
        ▼
 ② deck-composer ─────────▶ 03_exhibit-candidates.json (본문 장당 3안)
        │
        ▼
 ②.5 목업 승인(사용자) ───▶ 02_deck-spec.json 확정
        │
        ▼
 ③ pptx-builder ──────────▶ deck.pptx              ← build_pptx.py + brand-kit.yaml + goldenfab
        │                                             (골든 글 유출은 content_contract가 빌드 중단)
        ▼
 ④ consistency-qa ──┬─ PASS ─▶ results/<프로젝트명>-소개.pptx  (완주 PASS 0건 — 잔존 4건은 게이트 미통과분)
                    └─ FAIL ─▶ 원인 단계 되돌림(1회)          (K-IFRS 파일럿: FAIL 17건 후 동결)
```

각 단계의 입력·처리·출력과 분기·게이트 전체 그림은 [2_PIPELINE.md](2_PIPELINE.md).

## 읽는 순서

| #   | 파일                                           | 무엇을 담나                                                                         |
| --- | ---------------------------------------------- | ----------------------------------------------------------------------------------- |
| 0   | 0_README.md                                    | 표지: 배지·한줄요약·한눈 다이어그램·읽는 순서·핵심 수치                             |
| 1   | [1_OVERVIEW.md](1_OVERVIEW.md)                 | 문제 정의→왜 이 접근→아키텍처 원칙→기술 스택                                        |
| 2   | [2_PIPELINE.md](2_PIPELINE.md)                 | 전체 파이프라인 end-to-end + 실증 예시 1건                                          |
| 3   | [3_OUTLINE-GRILL.md](3_OUTLINE-GRILL.md)       | ① 목차·콘텐츠 인터뷰 게이트                                                         |
| 4   | [4_DECK-COMPOSE.md](4_DECK-COMPOSE.md)         | ② deck-spec 구성(선별·배치·레이아웃 매칭)                                           |
| 5   | [5_PPTX-BUILD.md](5_PPTX-BUILD.md)             | ③ 네이티브 빌더(build_pptx·brand-kit·design-rules·마스터 스탬프)                    |
| 6   | [6_VISUALS.md](6_VISUALS.md)                   | 시각 어휘층(visuals 13종+·mpl 9종·아이콘·목업 갤러리·아키타입 추천)                 |
| 7   | [7_GOLDEN-DECK.md](7_GOLDEN-DECK.md)           | 골든 덱 5부 17장·goldenfab·스냅샷 회귀·adapted/novel                                |
| 8   | [8_QA-GATES.md](8_QA-GATES.md)                 | 오딧·게이트 전 층(audit_golden/audit_deck/audit_pptx/check_contract/compare_golden) |
| 9   | [9_KIFRS-PILOT.md](9_KIFRS-PILOT.md)           | K-IFRS 1115 파일럿 현재 상태 — FAIL 17건·동결·미완 목록                             |
| 10  | [10_DIFFERENTIATION.md](10_DIFFERENTIATION.md) | 차별점·특장점(근거 필수) + 한계·트레이드오프                                        |
| 11  | [11_TEST-DECISIONS.md](11_TEST-DECISIONS.md)   | 검증 로그(M/N) + ADR                                                                |
| 12  | [12_JOURNEY.md](12_JOURNEY.md)                 | v1→v5·골든 전환 여정(과거형 전용)                                                   |
| 13  | [13_TROUBLESHOOTING.md](13_TROUBLESHOOTING.md) | 사고·근본원인·복구(과거형 전용)                                                     |
| 14  | [14_COVERAGE.md](14_COVERAGE.md)               | 전수 커버리지 부록                                                                  |

## 핵심 수치

| 수치                       | 값(실측 출처)                                                                                                                                           | 상술 장                              |
| -------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------ |
| 골든 덱 규모               | **5부 17장** (`goldenfab/reference.py` — 17장은 SLIDE_ORDER 17엔트리, 5부는 같은 파일 TOC 5엔트리 파생. SLIDE_ORDER의 `part` 슬라이드는 4장)            | [7_GOLDEN-DECK.md](7_GOLDEN-DECK.md) |
| 골든 스냅샷                | **519도형/17장** (`golden-snapshot.json`)                                                                                                               | [7_GOLDEN-DECK.md](7_GOLDEN-DECK.md) |
| compare_golden 최근 판정   | **17/17장 · 519/519도형 · 불일치 0 PASS** (`golden/variants/compare_full.md`)                                                                           | [8_QA-GATES.md](8_QA-GATES.md)       |
| goldenfab 레이아웃 타입    | **15종** (`goldenfab/registry.py`)                                                                                                                      | [7_GOLDEN-DECK.md](7_GOLDEN-DECK.md) |
| 슬라이드 타입 체계         | 레거시 **11종** + 골든 계열 **3종**(golden.\*/adapted.\*/novel)                                                                                         | [5_PPTX-BUILD.md](5_PPTX-BUILD.md)   |
| 시각 어휘                  | 네이티브 차트 6종 · mpl 9종 · 도형 다이어그램 18종 · Lucide 아이콘 24종(72 PNG)                                                                         | [6_VISUALS.md](6_VISUALS.md)         |
| 슬라이드 밀도 목표         | 120~200단어/장 (BCG 111슬라이드 실측 median 197 근거)                                                                                                   | [5_PPTX-BUILD.md](5_PPTX-BUILD.md)   |
| design-rules 규모          | **126항목**(불릿 110 + 번호 16, 2026-07-19 재계수) — 문서·이력의 "96항목" 표기는 재편 시점(커밋 04bd0ac, 2026-07-11) 값이고 이후 §H·§I·P5 추가로 늘었다 | [5_PPTX-BUILD.md](5_PPTX-BUILD.md)   |
| 오딧·게이트                | audit_golden 규칙 11종·SPECS 7종 / audit_deck 8항목 / 다양성 게이트 4종 / check_contract 6종                                                            | [8_QA-GATES.md](8_QA-GATES.md)       |
| K-IFRS 파일럿 QA 최종 판정 | **FAIL — 계약 위반 17건**(2026-07-15), 이후 동결                                                                                                        | [9_KIFRS-PILOT.md](9_KIFRS-PILOT.md) |
