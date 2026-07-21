# 물성 → 부품 매칭 결정표 (단일 출처)

deck-spec의 각 장에 시각 부품을 배정할 때 이 표 하나로 판정한다.
판정 기준은 **주제가 아니라 물성**(내용이 실제로 어떤 형태의 것인가) — design-rules Ⅲ부 P1.

> **왜 물성 하나로 통일하나 (2026-07-20).** 예전엔 같은 축을 세 이름으로 불렀다 — 물성(이 표)·
> 형상(archetype-catalog·recommend_archetypes)·주장형태(visual-selection). 축이 갈라져 배선이
> 안 맞았다. **물성을 정본 축으로 고정**하고 나머지 둘은 이 표의 근거(왜)·기계화(코드)로
> 종속시킨다. 물성 목록은 즉흥 발명이 아니라 **검증된 범용 taxonomy에 정합**한다:
> FT Visual Vocabulary 9범주(편차·상관·랭킹·분포·시계열·크기·부분전체·공간·흐름 — 정량 절반)
> + 데이터과학 다이어그램 4종(계층·네트워크·공정·매트릭스 — 구조 절반) + Abela 4목적.
> 그래서 이 표는 K-IFRS 골든 덱에 국한되지 않는다 — **다음 프로젝트의 다른 도메인도 같은 축**.

## 부품 두 계열

- **golden.\<layout\>** — goldenfab 15타입(고밀도·검증된 골든 부품). 구조·정성·실물 물성의 ✅정본.
  content dict 필수(텍스트 전량 교체, 좌표·색·도형 고정).
- **chart/diagram L-ID** — archetype-catalog L01~L38(범용 차트·도형). 정량 물성과, golden에 없는
  물성의 부품. 골든이 못 채운 칸을 메우는 나머지 절반이다.

**둘은 경쟁이 아니라 한 축의 양 절반이다.** golden = 구조·실물 물성(✅), chart L-ID = 정량 물성.
한 물성에 golden 타입이 있으면 그게 1순위, 없으면 L-ID로 내려간다. 어디에도 없으면 §검색.

## A. 구조·정성·실물 물성 → golden.\<layout\> (15타입 전수)

| 물성 (내용의 실제 형태)                         | 판정 질문                                 | golden 타입             | 없을 때 L-ID 폴백    |
| ----------------------------------------------- | ----------------------------------------- | ----------------------- | -------------------- |
| 덱의 첫 인상 — 제목·가치 제안                   | 이 장이 덱의 표지인가                     | `golden.cover`          | —                    |
| 부 목록과 이정표                                | 전체 구성을 안내하는 장인가               | `golden.toc`            | —                    |
| 부의 시작 선언                                  | 새 부(챕터)를 여는 간지인가               | `golden.part`           | —                    |
| 결격·문제의 다구획 해부(원인 4~5개 병렬)        | "왜 안 되는가"를 여러 구획으로 쪼개는가   | `golden.problem_grid`   | L18 cards            |
| 입력→출력의 분기 있는 실행 흐름                 | 무언가가 단계·분기를 거쳐 흐르는가        | `golden.exec_graph`     | L16 branch·L20       |
| 주장 + 실물 데이터 증거(표·JSON 카드)           | 실제 데이터 엔트리를 펼쳐 보여야 하는가   | `golden.tech_evidence`  | L26 table            |
| 위계 구조(트리·fan-out) + 구축 방법             | 계층 관계 자체가 내용인가                 | `golden.tech_tree`      | L15 layers           |
| 실물 화면(캡처)이 주인공                        | 스크린샷을 크게 보여주는 장인가           | `golden.screenshot`     | —                    |
| 메커니즘(걸림·주입 규칙) + 실물 적용 한 줄 서사 | "언제·어떻게 발동하나"를 설명하는가       | `golden.tech_mechanism` | L20 process_band     |
| 실물 결과 화면 + 스키마·검증 보조               | 산출물 화면과 그 형식 계약을 함께 보이나  | `golden.tech_capture`   | —                    |
| 구 방식 vs 신 방식 — 같은 입력의 동작 대비      | 전/후·경쟁 방식을 시뮬레이션으로 대비하나 | `golden.ab_simulation`  | L29 two_column·L19   |
| 시험 구성(재료 분해) + 최종 기록 + 실패 해부    | 검증·테스트 결과를 보고하는 장인가        | `golden.validation`     | L28 metrics          |
| 축별 좌우 대결(자사 vs 경쟁)                    | 여러 축에서 둘을 나란히 비교하나          | `golden.mirror_matrix`  | L27 table:matrix·L24 |
| 경계 — 안(보장)과 밖(거절·유보)                 | 못 하는 것·범위 한계를 선언하는 장인가    | `golden.boundary`       | L34 pro_con          |
| 덱을 닫는 마침 문장(수미상관)                   | 마지막 장인가                             | `golden.closing`        | —                    |
| 순서 없는 독립 항목 N개(병렬)                   | 서열 없는 요소 3~6을 나열하나             | (hero_card 계열)        | L18 cards            |

\* **전 15타입이 `content` dict를 받는다**(Phase B 파라미터화 완료). 텍스트만 교체되고 좌표·색·도형은 고정.

## B. 정량 물성 → chart L-ID (FT Visual Vocabulary 정합 — golden에 없는 절반)

| 물성 (FT 범주)            | 판정 질문(한 문장 테스트)               | 부품 L-ID                                    |
| ------------------------- | --------------------------------------- | -------------------------------------------- |
| 크기 비교 (magnitude)     | "A가 B보다 크다/많다"(동질 항목)        | L01 bar·L02 hbar·L31 bubble                  |
| 부분-전체 (part-to-whole) | "X가 전체의 N%"·"합이 이 조각들로"      | L04 pie·L05 stacked·L07 waterfall·L33 waffle |
| 시계열 변화 (over time)   | "시간에 따라 늘었다/줄었다"             | L03 line·L10 slope·L17 timeline              |
| 상관 (correlation)        | "X와 Y가 상관인데 저 점만 예외"         | L12 annotated_scatter·L31 bubble             |
| 분포 (distribution)       | "값들이 이렇게 퍼져 있다"               | L13 histogram                                |
| 흐름/전환 (flow)          | "단계마다 이만큼 빠져나간다"(전환·이탈) | L11 funnel·L20 process_band                  |
| 전후 비교 (paired)        | "두 시점 사이 항목별로 이만큼 이동"     | L09 dumbbell·L10 slope                       |
| 교차 밀도 (matrix)        | "행×열 어디가 뜨겁고 차가운가"          | L08 heatmap·L21 band_table                   |
| KPI 각인                  | "임팩트 숫자 자체가 메시지"(2~4개)      | L28 metrics·L36 stat_split                   |
| **편차 (deviation)**      | "기준선 대비 +/-"                       | ⬜ **미보유 → §검색**                        |
| **공간 (spatial)**        | "지리·위치가 핵심"                      | ⬜ **미보유(보류) → §검색**                  |

근거·안티패턴 상세: `pptx-visuals/references/visual-selection.md`(Zelazny 5비교·Abela 4목적).
L-ID 정형·세트제약: `pptx-visuals/references/archetype-catalog.md`. **셋은 같은 축의 다른 표현.**

## 축 크로스워크 (물성 ↔ 형상 ↔ 주장형태)

recommend_archetypes.py의 `형상`(시계열·구성비·전후·범주비교·교차다축·흐름전환·분포상관·정성·
시간여정·숫자)은 이 표 **B의 정량 물성과 1:1**이다(형상=물성의 정량 이름). visual-selection의
`주장형태`는 그 근거(왜 그 형식인가). **입력 축은 물성 하나** — 형상·주장형태는 별도 축이 아니라
같은 것의 코드·산문 표현이다.

> **content는 선택이 아니라 필수다.** `content`를 생략하면 골든 기본값(= 다른 프로젝트 글)이
> 조용히 나가므로, 공장 문은 content가 골든 DEFAULT 키를 **전부** 덮지 않으면 빌드를 중단한다
> (`goldenfab/content_contract.py`). 부분 주입도 FAIL — 누락 키가 골든 글로 메워지는 게
> "계약 0% 이행 덱이 전 게이트 통과" 사고의 본질이었다.

각 타입의 content 키는 [golden-content-contract.md](golden-content-contract.md)(코드 DEFAULT 덤프) 참조.
필수 키의 단일 출처는 `goldenfab/content_contract.required_keys()`(15/15 해석).
단일 소스는 `goldenfab/`이고, compare_golden.py는 **goldenfab 보호 전용** 회귀 게이트다 —
실전 덱의 adapted·novel 장은 스냅샷 대상이 아니다(audit_deck + 병렬 채점이 진다, design-rules Ⅲ부 P5).

## 판정 절차 (2026-07-16 개정 — 매칭 ≠ 텍스트 스왑 강제)

골든은 복제 템플릿이 아니라 **변형 가능한 출발점**이다(design-rules 골든셋 원칙).
매칭 결과는 셋 중 하나로 갈린다:

1. 재료(FINAL-REPORT)의 각 장 후보에 대해 **물성을 한 줄로 선언**한다
   ("이 내용 = 비교 / 흐름 / 분해 / 경계 / 실물 화면 / 시계열 / 분포 / …") — P1 물성 선언과 같은 형식.
2. 결정표(A→B 순)에서 그 물성의 행을 찾는다. 판정 질문에 "예"면 그 부품이 **출발점**이다.
   내용의 항목 수·구획 구성이 골든과 그대로 맞으면 `golden.<layout>` + content(텍스트 전량
   교체), 다르면 **`adapted.<layout>`** — 그 타입의 코드·렌더 PNG를 먼저 Read하고 goldenfab
   부품(kit·grid·피치 파생)을 재사용한 장 스크립트로 변형한다. 안 맞는 내용을 content에
   우겨넣는 것이 사용자가 기각한 "텍스트 스왑"이다.
3. **두 행에 걸치면** 주 물성(그 장이 답할 독자 질문에 직결된 쪽)을 따른다 —
   한 장에 두 물성을 우겨넣지 말고 장을 쪼갠다.
4. **어느 행에도 안 맞으면 끼워맞추지 않는다** — **§검색 플라이휠**(아래).

## 검색 플라이휠 (매칭 안 됨·⬜빈칸 — novel 부품 생성)

물성이 이 표에 있지만 부품이 ⬜(편차·공간 등)거나, 물성 자체가 표에 없는 경우:

1. 그 물성의 **검증된 시각 문법**을 확인 — 정량이면 FT Visual Vocabulary 해당 범주의 차트 leaf,
   구조면 다이어그램 관례(계층=tree, 네트워크=node-link VOWL/Graffoo 등). **즉흥 발명 금지**(P1).
   물성이 가장 가까운 골든 렌더 PNG 2~3장을 Read해 밀도·구성 앵커로 삼는다.
2. **`novel`**: P1 물성 선언부터 신규 설계(design-rules P5). 통과작 리치니스(밀도·마감)에 맞춘다.
3. `preflight_dense.py`(또는 audit_deck) green 확인 후에만 제시.
4. 골든 창고 편입은 덱 완성 후 **사용자가 고른 장만**: goldenfab 타입 승격 + design-rules 박제
   + compare_golden 케이스 추가 + **이 표에 행 추가**(⬜→✅). 다음 프로젝트가 재사용. (골든 1장
   = 유지보수 계약 1개, 05-plan 결정 #8)

## 강제 장치

- 물성 선언은 계약(contract) 1번 항목 — 완성 장이 선언과 다르면 FAIL (design-rules P1).
- 셀프 반려 체크 ②"메커니즘이 형태로 보이나"·⑦"주제와 무관한 채움인가"가 억지 매칭을
  제시 전에 차단한다 (design-rules P4).
- 세트 제약(≥5종·쿨다운·박스 30%)은 consistency-qa 게이트가 기계 강제(archetype-catalog 세트제약).
