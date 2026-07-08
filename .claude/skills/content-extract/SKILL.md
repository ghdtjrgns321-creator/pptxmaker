---
name: content-extract
description: NotebookLM이 생성한 pptx 초안 3~5개에서 텍스트·표·차트데이터·이미지를 추출해 통합 재료(01_extracted.md)를 만든다. 새 정보를 창작하지 않고 파일에 있는 것만 뽑는다. content-extractor 에이전트가 파이프라인 1단계를 수행할 때, 또는 "pptx에서 내용 추출/재추출/초안 다시 파싱" 요청 시 반드시 사용.
---

# content-extract — 외부 pptx 재료 추출

NotebookLM pptx 초안들을 파싱해 다음 단계(deck-compose)가 쓸 **재료**를 만든다.
창작 금지 — 파일에 존재하는 텍스트·수치만, 출처(파일명·슬라이드 번호)와 함께 뽑는다.

## 실행

```
uv run python .claude/skills/content-extract/scripts/extract_pptx.py \
    <작업디렉토리>/_workspace/extract  input/<프로젝트명>/*.pptx
```

산출물: `extract/extracted.md`(슬라이드별 텍스트·표·차트·이미지 목록),
`extract/charts.json`(네이티브 차트 수치), `extract/images/`(그림 폴백).

## 차트 3단계 폴백 (A안 박제)

1. **네이티브 차트** → 스크립트가 categories/series 수치를 charts.json으로 직접 추출
2. **도형 뭉치 차트** → 수치가 텍스트로 잡힌다. extracted.md의 해당 슬라이드 텍스트에서
   수치를 복원해 01_extracted.md에 "복원(출처: 파일·슬라이드)"로 표기
3. **이미지 차트** → images/에 저장됨. 수치 복원이 불가하면 01_extracted.md에
   **"재생성 불가, 이미지 폴백"**을 명시(조용히 버리지 않는다)

## 출력: `_workspace/01_extracted.md`

스크립트 산출물(extracted.md)을 그대로 넘기지 말고, 에이전트가 **정리**해서 쓴다:

- 초안별(3~5개) 핵심 메시지·수치·근거를 카테고리(문제/솔루션/기능/차별점/성과)로 분류
- 각 항목에 출처 표기: `(초안2 s5)` — deck-compose가 상충 시 원본을 추적할 수 있게
- 차트 재료는 charts.json 인덱스 또는 복원 수치를 명시, 이미지 폴백은 경로 명시
- 초안 간 **상충 수치는 삭제하지 않고 병기** — 선택은 deck-compose 몫

## 에러 핸들링

- pptx가 0개면 진행하지 않고 "input/ 비어 있음"을 오케스트레이터에 보고
- 파싱 예외 파일은 건너뛰고 목록에 명시(전체 중단 금지)
- NotebookLM 샘플 미확보 상태로 설계됨 — 실제 파일에서 새 형태(SmartArt 등)를 만나면
  extracted.md에 "미지원 형태" 표기하고 스킬/스크립트 보강을 제안
