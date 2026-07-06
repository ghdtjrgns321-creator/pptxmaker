---
name: pptmaker
description: 내가 만든 프로젝트를 소개하는 B2B 셀링 PPT를 자동 생성하는 오케스트레이터. 프로젝트를 분석→셀링 관점 큐레이션→네이티브 pptx 빌드→일관성 검증의 파이프라인으로 조율한다. "프로젝트 PPT 만들어줘/소개 자료 생성/발표자료 만들어/pptx 생성" 및 후속 "PPT 다시 만들어/특정 슬라이드 수정/브랜드 바꿔서 재빌드/스토리 다시/이 프로젝트로 갱신" 요청 시 반드시 사용.
---

# pptmaker — PPT 자동생성 오케스트레이터

기존 프로젝트를 읽어 B2B 셀링 PPT를 찍어낸다. **공장 일관성**(같은 골격·brand-kit)과 **셀링
전달력**이 목표다. 서브 에이전트 파이프라인으로 조율하며, 중간 산출물은 `_workspace/`에 남긴다.

**실행 모드:** 서브 에이전트 파이프라인. 각 에이전트를 `Agent` 도구로 순차 호출하고
(`model: "opus"` 필수), 결과는 `_workspace/` 파일로 다음 단계에 넘긴다. QA 실패 시 지정된
앞 단계로 1회 되돌린다.

## Phase 0: 컨텍스트 확인 (초기 / 후속 / 부분 재실행 판별)

작업 시작 전 대상 프로젝트 경로를 확인하고 `_workspace/` 상태를 본다:

- 프로젝트 경로 미지정 → 사용자에게 **분석할 프로젝트 폴더 경로**를 묻는다(필수 입력).
- `_workspace/` 없음 → **초기 실행** (Phase 1부터 전체)
- `_workspace/` 있음 + "특정 슬라이드/스토리만 수정" → **부분 재실행** (해당 에이전트만: 카피 수정이면 selling-curator→pptx-builder, 색·폰트면 brand-kit→pptx-builder)
- `_workspace/` 있음 + "다른 프로젝트/새로" → 기존 `_workspace/`를 `_workspace_prev/`로 옮기고 **새 실행**
- brand-kit 값 변경 요청 → `assets/brand-kit.yaml`만 수정 후 pptx-builder 재빌드(분석·큐레이션 건너뜀)

## 파이프라인 (초기 실행)

```
① project-analyst   → _workspace/01_facts.md        (프로젝트 사실 추출)
② selling-curator   → _workspace/02_deck-spec.json  (셀링 스토리 큐레이션)
③ pptx-builder      → _workspace/deck.pptx          (네이티브 pptx 빌드)
④ consistency-qa    → _workspace/03_qa-report.md    (일관성·품질 검증)
                      FAIL → 지정 단계로 1회 되돌림 → 재빌드 → 재검증
```

각 에이전트는 자기 스킬(project-analysis / selling-curation / pptx-build / consistency-qa)을
읽고 수행한다. 오케스트레이터는 순서·되돌림·최종 종합만 담당한다.

## 데이터 흐름 (파일 기반 전달)

- 작업 디렉토리 하위 `_workspace/`에 중간 산출물 저장, 파일명은 `{순번}_{산출물}` 규칙.
- 계약: `02_deck-spec.json`은 `skills/pptx-build/references/deck-spec-schema.md` 스키마 준수.
- 최종 산출물 `deck.pptx`만 사용자 지정 경로(기본 프로젝트 루트)로 복사, `_workspace/`는 감사용 보존.
- 골격 검증: deck-spec이 cover→toc로 시작하고 cta로 끝나는지 오케스트레이터가 ③ 전에 확인.

## 에러 핸들링

| 상황                                  | 대응                                                                                                                |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------------------- |
| project-analyst가 사실을 거의 못 찾음 | 사용자에게 프로젝트 경로·범위 재확인, 빈 재료로 진행 금지                                                           |
| deck-spec이 스키마 위반               | selling-curator에 스키마 명시하여 1회 재요청, 재실패 시 위반 항목 보고                                              |
| 빌드 예외(python-pptx)                | 예외 메시지와 문제 슬라이드를 pptx-builder에 전달해 1회 재시도                                                      |
| QA FAIL                               | 리포트의 [되돌릴 대상]으로 1회 되돌림. 재실패 시 리포트에 잔여 결함 명시하고 산출물과 함께 전달(조용히 넘기지 않음) |
| python-pptx 미설치                    | `pip install python-pptx` 안내 후 사용자 승인받아 설치                                                              |

원칙: 1회 재시도 후 재실패면 그 결함을 **명시**하고 진행한다. 상충/불확실 데이터는 삭제하지
않고 출처와 함께 남긴다.

## 테스트 시나리오

**정상 흐름:** 사용자가 프로젝트 경로 제공 → ① 사실 추출(기능 4·스택·성과 숫자 2) → ②
8슬라이드 deck-spec(cover→toc→문제→솔루션→기능→스택→성과chart→cta) → ③ deck.pptx 빌드 →
④ QA: 표/차트 네이티브 확인·폰트 위반 0·골격 준수 = PASS → 사용자에게 경로 보고.

**에러 흐름(되돌림):** ③ 빌드 후 ④ QA에서 chart 슬라이드가 그림(PICTURE)으로 잡힘 → FAIL,
[되돌릴 대상: pptx-builder] → deck-spec의 chart를 네이티브로 재빌드 → 재검증 PASS → 진행.
1회 재시도로도 실패하면 03_qa-report.md에 "chart N개 네이티브화 실패" 명시하고 사용자에 보고.

## 후속 작업

description의 후속 키워드로 재진입한다. Phase 0에서 부분 재실행을 판별해 필요한 에이전트만
재호출한다 — 전체를 다시 돌리지 않는다.
