# pptmaker — 부품 창고 기반 B2B 셀링 PPT 서식 공장

![Python](https://img.shields.io/badge/Python-3.13-3776AB)
![python-pptx](https://img.shields.io/badge/python--pptx-%E2%89%A51.0.2-15171B)
![PyYAML](https://img.shields.io/badge/PyYAML-%E2%89%A56.0.3-15171B)
![matplotlib](https://img.shields.io/badge/matplotlib-%E2%89%A53.11.0-15171B)
![uv](https://img.shields.io/badge/env-uv%20%C2%B7%20.venv%20%C2%B7%20uv.lock-5A29E4)
![golden deck](https://img.shields.io/badge/golden_deck-5_parts_%C2%B7_17_slides-D66E3A)
![snapshot](https://img.shields.io/badge/snapshot-797_shapes-D66E3A)
![parts](https://img.shields.io/badge/parts-10_figures_%C2%B7_6_frames-D66E3A)
![regression](https://img.shields.io/badge/compare__golden-797%2F797_PASS-2E7D32)
![selection](https://img.shields.io/badge/verify__selection-8%2F8_PASS-2E7D32)
![wiring](https://img.shields.io/badge/test__wiring-PASS-2E7D32)
![status](https://img.shields.io/badge/status-WIP_%C2%B7_no_completed_deck-C62828)

> **현재 상태 배너 (2026-07-27)** — 이 프로젝트는 **미완성**이다.
>
> - **라인은 다 깔았는데 덱을 한 벌도 끝까지 만들어본 적이 없다.** ① 인터뷰 → ② 부품·틀 선택 → ③ 조립 빌드 → ④ 검수의 각 구간은 명령 한 줄로 재현되지만, 넷을 꿴 완주 PASS는 **0건**이다.
> - 산출 목적지 `results/k-ifrs1115/`에는 과거 빌드 3건이 남아 있으나(전부 git 미추적) 현행 체계 이전이거나 QA FAIL 판정분이다.
> - 유일한 실전 시도인 **K-IFRS 1115 파일럿은 2026-07-15 QA FAIL(계약 위반 17건) 후 동결**이다. 그 spec은 이제 현행 공장에서 빌드조차 되지 않는다 — 그 사이 공장이 네 번 바뀌었다. 상세: [9_KIFRS-PILOT.md](9_KIFRS-PILOT.md)
> - **완주를 막을 지점은 실측으로 확인돼 있다**(추측 아님): `composed` 장 타입을 ④ 검수층이 아직 모른다.
>
> 초록 배지(회귀·선택·배선 PASS)는 **기준선과 규칙**의 상태이고, 빨강(완주 산출물 부재)은 **실전 파이프라인**의 상태다. 둘을 섞지 않는다.

## 한 줄 요약

사용자가 미리 정리해 둔 프로젝트별 `FINAL-REPORT/`를 재료로, **골든 덱에서 꺼낸 도해 부품 10종과 배치 틀 6종을 재료를 세어 고르고 조립해** 일관 브랜드의 **네이티브 .pptx** B2B 셀링 덱을 찍어내는 서식 공장 — 템플릿에 텍스트를 채우는 방식이 아니라, **맞는 부품이 없으면 멈추고 게이트를 통과하지 못하면 빌드가 중단되는** 방식이다.

## 시스템 한눈

```
  [생산 축] 재료 → 덱                        [제작 축] 실물 → 부품
  ─────────────────────────────             ─────────────────────
  FINAL-REPORT ─ scan_material              실물 레퍼런스
       │ ① GRILL 인터뷰                          │ part-design
       ▼                                        ▼
  01.5_outline.md (계약)                    1층 원소 elements.py
       │ ② composer                              │
       ▼                                        ▼
  deck-spec.json ◀── 부품을 **고른다** ── 2층 도해 figures/ (10종)
       │  · 축 5개로 재료를 센다                 │
       │  · 후보 0이면 **멈춘다**                ▼
       │ ③ builder                          3층 배치 frames.py (6종)
       ▼                                          · 골든덱 = 조합 사례
  deck.pptx  ── composed / golden.* / adapted.* / novel / 레거시 8종
       │ ④ QA — 계약 3자 대조 → 오딧 → 눈검증
       ▼
  results/<프로젝트명>-소개.pptx     ← 완주 PASS 0건
```

각 단계의 입력·처리·출력과 게이트 전체 그림은 [2_PIPELINE.md](2_PIPELINE.md).

## 읽는 순서

| #   | 파일                                           | 무엇을 담나                                                                      |
| --- | ---------------------------------------------- | -------------------------------------------------------------------------------- |
| 0   | 0_README.md                                    | 표지: 배지·한줄요약·한눈 다이어그램·읽는 순서·핵심 수치                          |
| 1   | [1_OVERVIEW.md](1_OVERVIEW.md)                 | 정체 → 문제 정의 → 왜 이 접근 → 아키텍처 원칙 11 → 기술 스택                     |
| 2   | [2_PIPELINE.md](2_PIPELINE.md)                 | 두 축(생산·제작) + 4단계 end-to-end + 장 타입 4계열                              |
| 3   | [3_OUTLINE-GRILL.md](3_OUTLINE-GRILL.md)       | ① 목차·콘텐츠 인터뷰 게이트 + 재료 스캐너                                        |
| 4   | [4_DECK-COMPOSE.md](4_DECK-COMPOSE.md)         | ② **부품·틀 선택 규칙**(축 5개·2단 규칙) + 매칭 결정표 + 콘텐츠 계약             |
| 5   | [5_PPTX-BUILD.md](5_PPTX-BUILD.md)             | ③ 네이티브 빌더 — 4계열 디스패치·GRID·brand-kit·design-rules·빌드 중단 게이트    |
| 6   | [6_VISUALS.md](6_VISUALS.md)                   | 시각 어휘 두 축 — **(A) 부품 도구함 10종**(3층·계약 3조건) · (B) 차트·다이어그램 |
| 7   | [7_GOLDEN-DECK.md](7_GOLDEN-DECK.md)           | 골든 덱 5부 17장 — 조합 사례이자 회귀 기준, dense 승격, 부품화 무손실            |
| 8   | [8_QA-GATES.md](8_QA-GATES.md)                 | 게이트 12종 전 층 + 커버리지 갭(정직 기재)                                       |
| 9   | [9_KIFRS-PILOT.md](9_KIFRS-PILOT.md)           | K-IFRS 파일럿 — FAIL 17건·동결·왜 되살리지 않는가                                |
| 10  | [10_DIFFERENTIATION.md](10_DIFFERENTIATION.md) | 차별점·특장점(근거 필수) + 한계·트레이드오프                                     |
| 11  | [11_TEST-DECISIONS.md](11_TEST-DECISIONS.md)   | 검증 로그(M/N) + ADR                                                             |
| 12  | [12_JOURNEY.md](12_JOURNEY.md)                 | v1→부품 도구함까지 구조 전환 17번(과거형 전용)                                   |
| 13  | [13_TROUBLESHOOTING.md](13_TROUBLESHOOTING.md) | 사고 23건·근본원인·복구(과거형 전용)                                             |
| 14  | [14_COVERAGE.md](14_COVERAGE.md)               | 전수 커버리지 부록 — 165/165 verify PASS                                         |

## 핵심 수치

| 수치                  | 값(실측 출처)                                                                  | 상술 장                    |
| --------------------- | ------------------------------------------------------------------------------ | -------------------------- |
| 골든 덱 규모          | **5부 17장** (`goldenfab/reference.py` SLIDE_ORDER 17엔트리 · TOC 5엔트리)     | [7](7_GOLDEN-DECK.md)      |
| 골든 스냅샷           | **797도형/17장** (AUTO_SHAPE 330·TEXT_BOX 318·PICTURE 80·LINE 68·CHART 1)      | [7](7_GOLDEN-DECK.md)      |
| compare_golden 판정   | **17/17장 · 797/797도형 · 불일치 0 PASS** + param 12/12                        | [8](8_QA-GATES.md)         |
| **도해 부품**         | **10종** (`figures/`, 3,103줄) — 계약 3조건: 자리 독립·개수 파생·글 무소유     | [6](6_VISUALS.md)          |
| **배치 틀**           | **6종** (`frames.py`) — 전부 골든 실측 추출                                    | [4](4_DECK-COMPOSE.md)     |
| **선택 규칙**         | 축 5개(`sets·flow·order·extra·ends`) + 2단 규칙 · 부품 10종이 **9칸**에 갈린다 | [4](4_DECK-COMPOSE.md)     |
| **선택 규칙 재현**    | **골든 8장 8/8 일치** — 사람의 도해 선택을 재료만 세어 재현                    | [11](11_TEST-DECISIONS.md) |
| **채움률 하한**       | `FILL_MIN = 0.96` — 골든 본문 실측(96~114%), 선택·빌드 두 층에 배선            | [5](5_PPTX-BUILD.md)       |
| 골든에 부품이 물린 장 | **6장** (S4·S6·S8·S9·S11 ×2) — 교체 전 회차 픽셀 동일                          | [7](7_GOLDEN-DECK.md)      |
| goldenfab 규모        | 최상위 10,640 → **6,507줄**(−39%, 시안 8파일 삭제) + `figures/` 3,103줄        | [7](7_GOLDEN-DECK.md)      |
| 장 타입 계열          | **4계열** — `composed` / `golden.*` / `adapted.*`·`novel` / 레거시 8종         | [5](5_PPTX-BUILD.md)       |
| goldenfab 레이아웃    | **15종** (`registry.py`)                                                       | [7](7_GOLDEN-DECK.md)      |
| 시각 어휘 (B)         | 네이티브 차트 6종 · mpl 9종 · 도형 다이어그램 9종 · Lucide 아이콘 24종(72 PNG) | [6](6_VISUALS.md)          |
| 슬라이드 밀도 목표    | 120~200단어/장 (BCG 111슬라이드 실측 median 197 근거)                          | [1](1_OVERVIEW.md)         |
| design-rules 규모     | **148항목**(불릿 124 + 번호 24) — P0 + 3부                                     | [5](5_PPTX-BUILD.md)       |
| 게이트                | **12종** — 회귀·오딧·계약·선택·배선·훅. `generic_checks`가 규칙 단일 출처      | [8](8_QA-GATES.md)         |
| 배선 검사             | `test_wiring` 가드 2 + 통합 A~I 9종 · 래칫 3종(감소 전용)                      | [8](8_QA-GATES.md)         |
| 완주 PASS             | **0건** — 이 파이프라인으로 ④까지 통과한 실전 .pptx 없음                       | [9](9_KIFRS-PILOT.md)      |
| 커버리지              | **165/165** verify PASS (루트 82 + .claude 83)                                 | [14](14_COVERAGE.md)       |

## 재현 명령

```
uv run python golden/build_golden.py                                    # 골든 17장 빌드
uv run python .claude/skills/pptx-build/scripts/compare_golden.py       # 회귀 17/17·797/797
uv run python .claude/skills/pptx-build/scripts/audit_golden.py         # 골든 오딧
uv run python .claude/skills/pptx-build/scripts/verify_selection.py     # 선택 규칙 재현·분리·멈춤
uv run python .claude/skills/pptx-build/scripts/test_wiring.py          # 하네스 배선 전수
uv run python .claude/skills/pptx-build/scripts/scan_material.py FINAL-REPORT/   # 재료 세기
```

**단, 골든 빌드는 저자 머신에서만 재현된다** — `ref/`가 gitignore이고 S10 이미지가 타 프로젝트 절대경로다. 코드·스냅샷을 근거로 든 서술은 클론에서도 재현되고, 로컬 산출물을 근거로 든 서술은 그렇지 않다([10_DIFFERENTIATION.md](10_DIFFERENTIATION.md) §3-5).
