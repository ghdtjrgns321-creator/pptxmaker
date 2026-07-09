---
name: pptmaker
description: NotebookLM pptx 초안들을 통합해 일관 서식의 B2B 셀링 PPT를 찍어내는 오케스트레이터. 초안 추출→통합 큐레이션→네이티브 pptx 빌드→일관성 검증의 파이프라인으로 조율한다. "프로젝트 PPT 만들어줘/소개 자료 생성/발표자료 만들어/pptx 생성/초안 통합해줘" 및 후속 "PPT 다시 만들어/특정 슬라이드 수정/브랜드 바꿔서 재빌드/스토리 다시/이 프로젝트로 갱신" 요청 시 반드시 사용.
---

# pptmaker — PPT 서식 공장 오케스트레이터

콘텐츠 생산(리서치·초안)은 사용자가 NotebookLM에서 해온다. 이 하네스는 그 초안 3~5개를
**통합하고 일관 서식(같은 골격·brand-kit)으로 재빌드**하는 공장이다. 전체 설계는
`docs/PIPELINE.md` 참조. 중간 산출물은 `_workspace/`에 남긴다.

**실행 모드:** 서브 에이전트 파이프라인. 각 에이전트를 `Agent` 도구로 순차 호출하고
(`model: "opus"` 필수), 결과는 `_workspace/` 파일로 다음 단계에 넘긴다. QA 실패 시 지정된
앞 단계로 1회 되돌린다.

## Phase 0: 컨텍스트 확인 (초기 / 후속 / 부분 재실행 판별)

작업 시작 전 입력과 `_workspace/` 상태를 본다:

- **입력 우선순위(2026-07-08 A/B 확정):** ① 로컬 프로젝트 소스 경로(→ content-extract
  모드 A 직접 추출, 기본) ② `input/<프로젝트명>/*.md` NotebookLM 보고서(보조·관점 변주)
  ③ pptx(최후 수단·이미지 판독). 소스 경로도 input도 없으면 사용자에게 프로젝트 소스
  경로를 묻는다(필수 입력).
- `_workspace/` 없음 → **초기 실행** (Phase 1부터 전체)
- `_workspace/` 있음 + "특정 슬라이드/스토리만 수정" → **부분 재실행**
  (카피·구성이면 deck-composer→pptx-builder, 색·폰트면 brand-kit→pptx-builder,
  차트·다이어그램 수치면 deck-spec 수정→pptx-builder — pptx 직접 수정 금지)
- `_workspace/` 있음 + "다른 프로젝트/새로" → 기존 `_workspace/`를 `_workspace_prev/`로
  옮기고 **새 실행**
- brand-kit 값 변경 요청 → `pptx-build/assets/brand-kit.yaml`만 수정 후 재빌드(추출·통합 건너뜀)

## 파이프라인 (초기 실행)

```
입력: 로컬 소스(기본) 또는 input/<프로젝트명>/  (NotebookLM 보조)
① content-extractor  → _workspace/01_extracted.md   (+extract/charts.json, images/)
② deck-composer      → _workspace/03_exhibit-candidates.json (슬라이드별 시각 후보 2~3안)
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

각 에이전트는 자기 스킬(content-extract / deck-compose / pptx-build / consistency-qa)을
읽고 수행한다. 차트·다이어그램 레시피는 `pptx-visuals` 스킬이 단일 출처. 오케스트레이터는
순서·되돌림·최종 종합만 담당한다.

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
| input/에 pptx 없음      | NotebookLM 초안 생성 안내(docs/notebooklm-prompt.md), 빈 재료로 진행 금지                                           |
| 추출 결과가 거의 빈약   | 사용자에게 초안 품질·개수 재확인 요청(프롬프트 변주 추가 생성 제안)                                                 |
| 차트 수치 추출 실패     | 텍스트 복원 시도 → 실패 시 이미지 폴백을 리포트에 명시(조용히 버리지 않음)                                          |
| deck-spec이 스키마 위반 | deck-composer에 스키마 명시하여 1회 재요청, 재실패 시 위반 항목 보고                                                |
| 빌드 예외(python-pptx)  | 예외 메시지와 문제 슬라이드를 pptx-builder에 전달해 1회 재시도                                                      |
| QA FAIL                 | 리포트의 [되돌릴 대상]으로 1회 되돌림. 재실패 시 리포트에 잔여 결함 명시하고 산출물과 함께 전달(조용히 넘기지 않음) |
| python-pptx 미설치      | `uv sync` 안내 후 사용자 승인받아 실행                                                                              |

원칙: 1회 재시도 후 재실패면 그 결함을 **명시**하고 진행한다. 상충/불확실 데이터는 삭제하지
않고 출처와 함께 남긴다.

## 테스트 시나리오

**정상 흐름:** input/에 초안 4개 → ① 추출(텍스트·차트 수치 6건·이미지 폴백 1건) → ②
8슬라이드 deck-spec(cover→toc→문제→솔루션diagram→기능→차별점table→성과chart→cta) → ③
deck.pptx 빌드 → ④ QA: 표/차트 네이티브 확인·폰트 위반 0·골격 준수 = PASS → 사용자에게
경로 보고.

**에러 흐름(되돌림):** ③ 빌드 후 ④ QA에서 chart 슬라이드가 그림(PICTURE)으로 잡힘 → FAIL,
[되돌릴 대상: pptx-builder] → deck-spec의 chart를 네이티브로 재빌드 → 재검증 PASS → 진행.
1회 재시도로도 실패하면 03_qa-report.md에 "chart N개 네이티브화 실패" 명시하고 사용자에 보고.

## 후속 작업

description의 후속 키워드로 재진입한다. Phase 0에서 부분 재실행을 판별해 필요한 에이전트만
재호출한다 — 전체를 다시 돌리지 않는다. **모든 수정은 deck-spec/brand-kit 경유 재빌드** —
완성 pptx를 직접 뜯지 않는다.
