# pptmaker — FINAL REPORT

![Python](https://img.shields.io/badge/Python-3.13-3776AB?logo=python&logoColor=white)
![python-pptx](https://img.shields.io/badge/python--pptx-네이티브_개체-2B579A?logo=microsoftpowerpoint&logoColor=white)
![uv](https://img.shields.io/badge/uv-lockfile_재현-DE5FE9?logo=uv&logoColor=white)
![하네스](https://img.shields.io/badge/하네스-26파일_3185줄-555555)
![시각어휘](https://img.shields.io/badge/시각_어휘-24종-D66E3A)
![게이트](https://img.shields.io/badge/generic__checks-12종-1F6FEB)
![밀도밴드](https://img.shields.io/badge/밀도_하한-535자_실측파생-8A9099)
![정리](https://img.shields.io/badge/2026--07--29_정리-13710→3185줄-CC0000)
![실전완주](https://img.shields.io/badge/실전_완주-0건-999999)

**장마다 다른 형식으로 조판하는 B2B 셀링 PPT 공장** — 형식을 미리 배정하지 않고, 장마다
"무엇을 말할까"를 후보로 **그려서** 사용자가 고르게 한 뒤 파이썬으로 직접 조판한다.

## 시스템 한눈

```
 <프로젝트>/FINAL-REPORT/*.md          brand-kit.yaml (색 6·폰트 3·크기 8)
        │  보고서형 md 묶음                    │  단일 출처
        ▼                                     │
 ┌──────────────────────────────────────────┼──────────────────────────────┐
 │  ① 재료 읽기 ── 아웃라인 계약 없으면 인터뷰(deck-outline-grill)          │
 │        ▼        scan_material.py 가 표·목록·화살표를 기계로 센다         │
 │  ② 질문 후보 ── 장마다 "무엇을 말할까" 4~6개                             │
 │        ▼                                                                │
 │  ③ 갤러리   ── 후보를 실제 pptx로 그린다 ─────────┐                      │
 │        ▼                                          │ 왕복이 게이트다      │
 │  ④ 선택     ── 사용자 "s04=1 · s07=2 · s09=전탈락"┘  (전탈락 → ②로)     │
 │        ▼                                                                │
 │  ⑤ 조판     ── goldenfab kit·grid·dense·layouts + pptx-visuals 24종     │
 │        ▼                                                                │
 │  ⑥ 눈검증   ── render_deck.ps1 → PNG 전장 Read  ★ 진짜 게이트            │
 │        ▼                                                                │
 │  ⑦ 기계채점 ── score_deck.py (generic_checks 12 + 밀도 밴드 535자)       │
 └─────────────────────────────────────────────────────────────────────────┘
        ▼
 results/<프로젝트명>-소개.pptx
```

메인 대화가 직접 수행한다. 서브에이전트 파이프라인이 아니다 — 덱 만들기는 사용자와 주고받는
일이라 에이전트로 쪼개면 그 왕복이 끊긴다.

## 읽는 순서

| #   | 파일                                        | 무엇을 담나                                         |
| --- | ------------------------------------------- | --------------------------------------------------- |
| 1   | [1_OVERVIEW](1_OVERVIEW.md)                 | 무엇을 막는 시스템인가 · 왜 이 접근인가 · 기술 스택 |
| 2   | [2_PIPELINE](2_PIPELINE.md)                 | ①~⑦ 전체 흐름 + end-to-end 실증 예시                |
| 3   | [3_MATERIAL](3_MATERIAL.md)                 | 재료와 아웃라인 계약 · `scan_material.py`           |
| 4   | [4_GALLERY](4_GALLERY.md)                   | 질문 후보와 갤러리 선택 — 이 시스템의 중심          |
| 5   | [5_COMPOSE](5_COMPOSE.md)                   | 조판 도구함 `goldenfab` 1,573줄                     |
| 6   | [6_VISUALS](6_VISUALS.md)                   | 시각 어휘 24종 · 폐기 9종                           |
| 7   | [7_GATES](7_GATES.md)                       | 눈검증과 기계 채점 · 밀도 밴드 파생                 |
| 8   | [8_INHERITANCE](8_INHERITANCE.md)           | 반려를 어떻게 축적하나 (배치 사다리 4층)            |
| 9   | [9_DIFFERENTIATION](9_DIFFERENTIATION.md)   | 통상 접근 대비 차별점 · 한계                        |
| 10  | [10_TEST-DECISIONS](10_TEST-DECISIONS.md)   | 검증 M/N + 의사결정 ADR                             |
| 11  | [11_JOURNEY](11_JOURNEY.md)                 | 아키텍처 전환 — 무엇을 버렸나 (과거형)              |
| 12  | [12_TROUBLESHOOTING](12_TROUBLESHOOTING.md) | 실제 사고와 근본 원인 (과거형)                      |
| 13  | [13_COVERAGE](13_COVERAGE.md)               | 전수 커버리지 N/N                                   |

## 핵심 수치 (2026-07-31 실측)

| 항목                    | 값                                                                | 상술                      |
| ----------------------- | ----------------------------------------------------------------- | ------------------------- |
| 하네스                  | 26파일 · 코드 3,185줄 · 문서 1,628줄                              | [5](5_COMPOSE.md)         |
| 조판 도구함 `goldenfab` | 1,573줄 (audit 828 · dense 248 · layouts 250 · kit 157 · grid 84) | [5](5_COMPOSE.md)         |
| 시각 어휘               | 24종 (네이티브 6 · mpl 9 · 도형 9)                                | [6](6_VISUALS.md)         |
| 기계 검사               | `generic_checks` 12종                                             | [7](7_GATES.md)           |
| 밀도 하한               | 글자 535자 · 도형 바닥 20 — 골든 스냅샷에서 **파생**(리터럴 아님) | [7](7_GATES.md)           |
| 브랜드 단일 출처        | 색 6 · 폰트 3 · 크기 8단계                                        | [5](5_COMPOSE.md)         |
| 고정 자산               | 골든 17장(도형 850) · 골든 렌더 17 · 부품 렌더 27 · 아이콘 73     | [5](5_COMPOSE.md)         |
| 반려 목록               | 8건                                                               | [8](8_INHERITANCE.md)     |
| 2026-07-29 정리         | 13,710줄 → 3,185줄 · 아카이브 1,400개                             | [11](11_JOURNEY.md)       |
| **실전 완주**           | **0건** — ③⑤⑥⑦은 실측으로 돌지만 한 벌을 끝까지 낸 적 없음        | [9](9_DIFFERENTIATION.md) |
