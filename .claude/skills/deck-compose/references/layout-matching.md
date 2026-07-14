# 물성 → 골든 레이아웃 타입 매칭 결정표 (단일 출처)

deck-spec의 각 장에 `"type": "golden.<layout>"`을 배정할 때 이 표로 판정한다.
판정 기준은 **주제가 아니라 물성**(내용이 실제로 어떤 형태의 것인가) — design-rules Ⅲ부 P1.

## 결정표 — 15타입 전수

| 물성 (내용의 실제 형태)                         | 판정 질문                                 | 타입                    | 콘텐츠 |
| ----------------------------------------------- | ----------------------------------------- | ----------------------- | ------ |
| 덱의 첫 인상 — 제목·가치 제안                   | 이 장이 덱의 표지인가                     | `golden.cover`          | dict   |
| 부 목록과 이정표                                | 전체 구성을 안내하는 장인가               | `golden.toc`            | dict   |
| 부의 시작 선언                                  | 새 부(챕터)를 여는 간지인가               | `golden.part`           | dict   |
| 결격·문제의 다구획 해부(원인 4~5개 병렬)        | "왜 안 되는가"를 여러 구획으로 쪼개는가   | `golden.problem_grid`   | dict*  |
| 입력→출력의 분기 있는 실행 흐름                 | 무언가가 단계·분기를 거쳐 흐르는가        | `golden.exec_graph`     | dict*  |
| 주장 + 실물 데이터 증거(표·JSON 카드)           | 실제 데이터 엔트리를 펼쳐 보여야 하는가   | `golden.tech_evidence`  | dict*  |
| 위계 구조(트리·fan-out) + 구축 방법             | 계층 관계 자체가 내용인가                 | `golden.tech_tree`      | dict*  |
| 실물 화면(캡처)이 주인공                        | 스크린샷을 크게 보여주는 장인가           | `golden.screenshot`     | dict   |
| 메커니즘(걸림·주입 규칙) + 실물 적용 한 줄 서사 | "언제·어떻게 발동하나"를 설명하는가       | `golden.tech_mechanism` | dict*  |
| 실물 결과 화면 + 스키마·검증 보조               | 산출물 화면과 그 형식 계약을 함께 보이나  | `golden.tech_capture`   | dict*  |
| 구 방식 vs 신 방식 — 같은 입력의 동작 대비      | 전/후·경쟁 방식을 시뮬레이션으로 대비하나 | `golden.ab_simulation`  | dict*  |
| 시험 구성(재료 분해) + 최종 기록 + 실패 해부    | 검증·테스트 결과를 보고하는 장인가        | `golden.validation`     | dict*  |
| 축별 좌우 대결(자사 vs 경쟁)                    | 여러 축에서 둘을 나란히 비교하나          | `golden.mirror_matrix`  | dict*  |
| 경계 — 안(보장)과 밖(거절·유보)                 | 못 하는 것·범위 한계를 선언하는 장인가    | `golden.boundary`       | dict*  |
| 덱을 닫는 마침 문장(수미상관)                   | 마지막 장인가                             | `golden.closing`        | dict   |

\* **전 15타입이 `content` dict를 받는다**(Phase B 파라미터화 완료). `golden.problem_grid`~`boundary`
10종은 **텍스트 override** — `content` 없으면 골든 기본값(회귀 0), 주면 텍스트만 교체(좌표·색·도형 고정).
각 타입의 content 키는 [golden-content-contract.md](golden-content-contract.md)(코드 DEFAULT 덤프) 참조.
새 프로젝트는 물성 매칭 후 그 타입에 자기 텍스트를 override로 넣는다 — 골든 원본(golden/)은 동결,
공장(goldenfab/)만 확장. compare_golden.py가 회귀·override 유효성을 게이트한다.

## 판정 절차

1. 재료(FINAL-REPORT)의 각 장 후보에 대해 **물성을 한 줄로 선언**한다
   ("이 내용 = 비교 / 흐름 / 분해 / 경계 / 실물 화면 / …") — P1 물성 선언과 같은 형식.
2. 결정표에서 그 물성의 행을 찾는다. 판정 질문에 "예"면 그 타입.
3. **두 행에 걸치면** 주 물성(그 장이 답할 독자 질문에 직결된 쪽)을 따른다 —
   한 장에 두 물성을 우겨넣지 말고 장을 쪼갠다.
4. **어느 행에도 안 맞으면 끼워맞추지 않는다** — 골든 확장 절차(pull):
   물성 선언 → 새 아키타입 시안 → 사용자 승인 → design-rules 박제 + goldenfab 타입 추가
   + compare_golden 케이스 추가. (골든 1장 = 유지보수 계약 1개, 05-plan 결정 #8)

## 강제 장치

- 물성 선언은 계약(contract) 1번 항목 — 완성 장이 선언과 다르면 FAIL (design-rules P1).
- 셀프 반려 체크 ②"메커니즘이 형태로 보이나"·⑦"주제와 무관한 채움인가"가 억지 매칭을
  제시 전에 차단한다 (design-rules P4).
