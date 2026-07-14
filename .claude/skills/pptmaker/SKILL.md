---
name: pptmaker
description: FINAL-REPORT를 재료로 일관 서식의 B2B 셀링 PPT를 찍어내는 오케스트레이터. 목차 인터뷰→통합 구성→네이티브 pptx 빌드→일관성 검증의 파이프라인으로 조율한다. "프로젝트 PPT 만들어줘/소개 자료 생성/발표자료 만들어/pptx 생성/이 프로젝트로 덱 만들어" 및 후속 "PPT 다시 만들어/특정 슬라이드 수정/브랜드 바꿔서 재빌드/스토리 다시/이 프로젝트로 갱신" 요청 시 반드시 사용.
---

# pptmaker — PPT 서식 공장 오케스트레이터

콘텐츠는 사용자가 프로젝트별 `FINAL-REPORT/`(보고서형 md 묶음)로 미리 정리해둔다. 이 하네스는
그 FINAL-REPORT를 **목차 인터뷰로 셀링 골격에 앉히고 일관 서식(같은 골격·brand-kit)으로 빌드**하는
공장이다. 전체 설계는 `docs/PIPELINE.md` 참조. 중간 산출물은 `_workspace/`에 남긴다.

**실행 모드:** 서브 에이전트 파이프라인. 각 에이전트를 `Agent` 도구로 순차 호출하고
(`model: "opus"` 필수), 결과는 `_workspace/` 파일로 다음 단계에 넘긴다. QA 실패 시 지정된
앞 단계로 1회 되돌린다.

## Phase 0: 컨텍스트 확인 (초기 / 후속 / 부분 재실행 판별)

작업 시작 전 입력과 `_workspace/` 상태를 본다:

- **입력(필수):** 프로젝트별 `FINAL-REPORT/`(보고서형 md 묶음 — 사용자가 미리 정리해둔
  단일 재료 원천). 경로 예: `<프로젝트>/FINAL-REPORT/*.md`. FINAL-REPORT가 없으면 진행하지
  않고 "먼저 FINAL-REPORT로 정리해달라"고 사용자에 요청한다(빈 재료 진행 금지).
- `_workspace/` 없음 → **초기 실행** (Phase 1부터 전체)
- `_workspace/` 있음 + "특정 슬라이드/스토리만 수정" → **부분 재실행**
  (카피·구성이면 deck-composer→pptx-builder, 색·폰트면 brand-kit→pptx-builder,
  차트·다이어그램 수치면 deck-spec 수정→pptx-builder — pptx 직접 수정 금지)
- `_workspace/` 있음 + "다른 프로젝트/새로" → 기존 `_workspace/`를 `_workspace_prev/`로
  옮기고 **새 실행**
- brand-kit 값 변경 요청 → `pptx-build/assets/brand-kit.yaml`만 수정 후 재빌드(추출·통합 건너뜀)

## 파이프라인 (초기 실행)

```
입력: <프로젝트>/FINAL-REPORT/*.md  (사용자가 정리해둔 단일 재료 원천 — 추출 단계 없음)
① 목차·콘텐츠 인터뷰 게이트(deck-outline-grill) — 오케스트레이터가 메인 대화에서 직접 수행
                      (서브에이전트 금지 — 사용자와의 인터뷰 필요). FINAL-REPORT 기반 목차
                      추천 1안 → 부·장 단위 "가장 전하고 싶은 것" 채록 →
                      _workspace/01.5_outline.md 확정. **아웃라인 "확정" 전 ② 진행 금지.**
② deck-composer      → _workspace/03_exhibit-candidates.json (슬라이드별 시각 후보 2~3안)
                      입력: FINAL-REPORT + 01.5_outline.md(필수 계약) — 장 추가·삭제·메시지 변경 금지
②.5 익스히빗 승인 게이트(v4) — make_mockups.py로 _workspace/mockups/gallery.html 생성,
                      사용자에게 열어주고 "s07=B" 회신 대기. **승인 전 ③ 진행 금지.**
                      승인 반영 → _workspace/02_deck-spec.json 확정
③ pptx-builder       → _workspace/deck.pptx         (하이브리드: 네이티브 + mpl 익스히빗)
④ consistency-qa     → _workspace/03_qa-report.md   (일관성·품질·다양성 게이트 3종)
                      audit 기계검사 + 전장 PNG 렌더 눈검증(겹침·잘림·공백) 필수
                      FAIL → 지정 단계로 1회 되돌림 → 재빌드 → 재검증
```

②.5 게이트는 자동 진행 금지 항목이다(§9 Decision Gating) — 사용자가 명시적으로 "목업 생략"을
요청한 경우에만 deck-composer가 결정표 최적안으로 단독 확정한다.

① 게이트도 동일하게 자동 진행 금지 — "PPT 만들어줘" 요청 시 빌드 전 반드시 1회 발동한다.
역할 분담: ①은 "무엇을 말할 것인가"(구성·강조점), ②.5는 "어떻게 보일 것인가"(시각)를
각각 승인받는다. 상세 인터뷰 절차는 `deck-outline-grill` 스킬.

각 에이전트는 자기 스킬(deck-compose / pptx-build / consistency-qa)을 읽고 수행한다.
차트·다이어그램 레시피는 `pptx-visuals` 스킬이 단일 출처. ①(목차 인터뷰)만은 사용자 대화가
필요해 오케스트레이터가 메인에서 직접 수행한다. 오케스트레이터는 순서·되돌림·최종 종합을 담당한다.

## 데이터 흐름 (파일 기반 전달)

- 작업 디렉토리 하위 `_workspace/`에 중간 산출물 저장, 파일명은 `{순번}_{산출물}` 규칙.
- 계약: `02_deck-spec.json`은 `skills/pptx-build/references/deck-spec-schema.md` 스키마 준수.
- 최종 산출물 `deck.pptx`만 `results/`(기본 `pptmaker/results/<프로젝트명>-소개.pptx`)로 복사,
  `_workspace/`는 감사용 보존. 사용자가 다른 경로를 지정하면 그쪽 우선. 한글 파일명 복사는
  PowerShell 사용(bash는 인코딩 깨짐).
- 골격 검증: deck-spec이 cover→toc로 시작하고 cta로 끝나는지 오케스트레이터가 ③ 전에 확인.

## 에러 핸들링

| 상황                    | 대응                                                                                                                |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------- |
| FINAL-REPORT 없음       | 사용자에 FINAL-REPORT 정리 요청, 빈 재료로 진행 금지                                                                |
| FINAL-REPORT가 얇음     | GRILL 인터뷰에서 보강 재료 요청(추가 문서·스크린샷), 억지 창작 금지                                                 |
| 수치 근거가 없음        | 해당 차트/성과는 넣지 않거나 정성 강점으로 대체(없는 수치 창작 금지)                                                |
| deck-spec이 스키마 위반 | deck-composer에 스키마 명시하여 1회 재요청, 재실패 시 위반 항목 보고                                                |
| 빌드 예외(python-pptx)  | 예외 메시지와 문제 슬라이드를 pptx-builder에 전달해 1회 재시도                                                      |
| QA FAIL                 | 리포트의 [되돌릴 대상]으로 1회 되돌림. 재실패 시 리포트에 잔여 결함 명시하고 산출물과 함께 전달(조용히 넘기지 않음) |
| python-pptx 미설치      | `uv sync` 안내 후 사용자 승인받아 실행                                                                              |

원칙: 1회 재시도 후 재실패면 그 결함을 **명시**하고 진행한다. 상충/불확실 데이터는 삭제하지
않고 출처와 함께 남긴다.

## 테스트 시나리오

**정상 흐름:** <프로젝트>/FINAL-REPORT/*.md → ① 목차 인터뷰(추천 1안 → 장별 강조 채록 →
01.5_outline 확정) → ② deck-composer가 8슬라이드 deck-spec(cover→toc→문제→솔루션diagram→
기능→차별점table→성과chart→cta) → ②.5 익스히빗 승인 → ③ deck.pptx 빌드 → ④ QA: 표/차트
네이티브 확인·폰트 위반 0·골격 준수·FINAL-REPORT 근거 대응 = PASS → 사용자에게 경로 보고.

**에러 흐름(되돌림):** ③ 빌드 후 ④ QA에서 chart 슬라이드가 그림(PICTURE)으로 잡힘 → FAIL,
[되돌릴 대상: pptx-builder] → deck-spec의 chart를 네이티브로 재빌드 → 재검증 PASS → 진행.
1회 재시도로도 실패하면 03_qa-report.md에 "chart N개 네이티브화 실패" 명시하고 사용자에 보고.

## 후속 작업

description의 후속 키워드로 재진입한다. Phase 0에서 부분 재실행을 판별해 필요한 에이전트만
재호출한다 — 전체를 다시 돌리지 않는다. **모든 수정은 deck-spec/brand-kit 경유 재빌드** —
완성 pptx를 직접 뜯지 않는다.
