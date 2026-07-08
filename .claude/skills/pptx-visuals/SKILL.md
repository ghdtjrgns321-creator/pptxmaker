---
name: pptx-visuals
description: PPTX 안의 네이티브 차트(bar/line/pie)와 다이어그램(flow/layers 도형 DSL)을 생성·수정하는 레시피 단일 출처. pptx-builder가 차트·다이어그램 슬라이드를 렌더할 때, 또는 "차트 수치 바꿔/다이어그램 수정/그래프 추가/시각자료 손봐" 요청 시 반드시 사용.
---

# pptx-visuals — 차트·다이어그램 단일 출처

PPT 시각자료의 생성·수정 레시피를 박제한 스킬. 구현은 `scripts/visuals.py` 하나이며
`build_pptx.py`가 import한다. **시각자료 코드는 이 모듈에만 존재한다** — 빌더나 다른 곳에
차트 코드를 중복 작성하지 않는다.

## 차트 (네이티브 — 이미지 금지)

- `add_chart(slide, spec, brand, x, y, w, h)` — deck-spec의 `chart` 슬라이드를 렌더.
- 타입: `bar`(비교/성장) · `line`(추세) · `pie`(비중). 값은 spec의 숫자 그대로 —
  PPT에서 데이터만 고치면 자동 반영되는 것이 핵심 가치.
- 시리즈 색은 brand 팔레트(accent→primary→muted 순환)를 자동 적용. Office 기본 테마색 금지.

## 다이어그램 (도형 DSL — python-pptx에 다이어그램 객체가 없어 도형+화살표로 렌더)

- `add_diagram(slide, spec, brand, x, y, w, h)` — deck-spec의 `diagram` 슬라이드를 렌더.
- 레이아웃 2종(YAGNI — 필요 시 추가):
  - `flow`: 노드 좌→우 + accent 화살표. 프로세스·파이프라인 흐름용. 노드 3~5개 권장.
  - `layers`: 노드 상→하 적층. 아키텍처 레이어용. 노드 2~5개 권장.
- 노드: `{"label": "굵은 제목", "sub": "설명 한 줄(선택)"}`. 스키마 상세는
  `pptx-build/references/deck-spec-schema.md`.

## 수정 원칙 (박제)

**수정은 항상 deck-spec을 고치고 재빌드한다.** 완성된 .pptx를 직접 뜯어고치지 않는다 —
deck-spec이 단일 출처여야 "브랜드 바꿔 재빌드" 같은 공장 재현성이 유지된다.

- "3번 슬라이드 차트 수치 바꿔" → `02_deck-spec.json`의 `series` 값 수정 → 재빌드
- "다이어그램에 단계 추가" → `nodes`에 항목 추가 → 재빌드
- 색·폰트 불만 → `brand-kit.yaml` 수정 → 재빌드 (spec은 안 건드림)

## 확장 시 규칙

- 새 차트 타입·레이아웃은 `visuals.py`에 추가하고 스키마 문서에 필드를 등록한 뒤,
  `deck-compose` 스킬의 타입 선택 가이드에 사용 기준을 한 줄 추가한다(3곳 동기화).
- 렌더 좌표·크기 상수는 호출자(build_pptx.py)가 주므로 이 모듈에 캔버스 상수를 두지 않는다.
