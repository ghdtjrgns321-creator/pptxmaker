---
name: pptx-visuals
description: PPTX 안의 네이티브 차트(bar/line/pie)와 다이어그램(timeline/matrix_2x2/venn 등 도형 DSL 9종)을 생성·수정하는 레시피 단일 출처. pptx-builder가 차트·다이어그램 슬라이드를 렌더할 때, 또는 "차트 수치 바꿔/다이어그램 수정/그래프 추가/시각자료 손봐" 요청 시 반드시 사용.
---

# pptx-visuals — 차트·다이어그램 단일 출처

PPT 시각자료의 생성·수정 레시피를 박제한 스킬. 구현은 `scripts/visuals.py`(네이티브)·
`scripts/mpl_exhibits.py`(mpl 이미지)·`scripts/make_mockups.py`(승인 게이트 갤러리) 3개이며
`build_pptx.py`가 import한다. **시각자료 코드는 이 모듈에만 존재한다** — 빌더나 다른 곳에
차트 코드를 중복 작성하지 않는다.

## v4 하이브리드 익스히빗 (mpl_exhibits.py — 도구 천장 돌파)

- 네이티브가 못 만드는 7종: `waterfall`·`heatmap`·`dumbbell`·`slope`·`funnel`·
  `annotated_scatter`·`histogram` — `render(spec, brand, out_png, w_in, h_in)`이 220dpi PNG
  생성, 빌더가 그림으로 삽입(**수치편집불가** — 수정은 deck-spec 재빌드).
- **BCG 스타일 단일 경로**: 모든 렌더러는 `_fig()`→`_save()` 경유 + `_palette()` 색만 사용 —
  전부 회색(C7CBD1) + 강조만 accent 1색 + 콜아웃(`annotations`) + 각주(`note`).
- 네이티브 차트에도 같은 규칙: `emphasis` 필드(add_chart), 콜아웃은 `add_callouts()`가
  pptx 도형(편집가능)으로 얹는다.
- 목업 갤러리: 기본은 실물 렌더 `render_real_mockups.py`(빌더+PowerPoint COM, 인터랙티브
  세트 심사), 폴백은 `make_mockups.py`(mpl 스케치). 사용 절차는 deck-compose SKILL.
- 후보 검증기: `check_candidates.py` — 3안·shape·why_not·조합장치·게이트 시뮬레이션.

## v4.2 아키타입 카탈로그 (`references/archetype-catalog.md` 단일 출처)

- L01~L30 전 아키타입 + 형상→L-ID 매핑 + 세트 제약(쿨다운·≤2회·≥5종·박스 30%).
- **카탈로그 행 = 렌더러 실존 보증.** 새 아키타입은 렌더러 구현·스모크 후에만 등재.
- v4.2 신규 렌더러 4종(visuals.py): `matrix_2x2`(2×2 포지셔닝)·`spectrum`(성숙도 마커)·
  `harvey_table`(하비볼 0~4)·`check_matrix`(✓/✗/–) — 정성 서술의 시각 구조화 전담.

## 차트 (네이티브 — 이미지 금지)

- `add_chart(slide, spec, brand, x, y, w, h)` — deck-spec의 `chart` 슬라이드를 렌더.
- 값은 spec의 숫자 그대로 — PPT에서 데이터만 고치면 자동 반영되는 것이 핵심 가치.
- 시리즈 색은 brand 팔레트(accent→primary→muted 순환)를 자동 적용. Office 기본 테마색 금지.

### 타입 선택 — `references/visual-selection.md`가 단일 출처

**시각은 주장의 증거 형태다(설명+분석+설득 목적, 장식 아님).** 주장 한 문장 → 결정표 →
형식. 안티패턴(이종 재고를 bar로, 병렬 목록을 layers로, 합≠100%를 pie로)은 즉시 교체.
아래 표는 정량 차트 요약이며 전체 결정표·안티패턴은 references 문서 참조.

### 정량 차트 요약 (데이터 형태를 먼저 판별 — 기본값 bar 금지)

| 데이터 형태                         | 타입          |
| ----------------------------------- | ------------- |
| 항목 크기 비교, 카테고리명 짧음     | `bar`         |
| 순위·비중 비교, 카테고리명 김       | `hbar`        |
| 시계열 추세                         | `line`        |
| 구성비(합≈100%, 항목 ≤5)            | `pie`         |
| 구성비 + 도넛 선호                  | `doughnut`    |
| 구성이 여러 그룹에 걸쳐 변하는 비교 | `stacked_bar` |

## 다이어그램 (도형 DSL — python-pptx에 다이어그램 객체가 없어 도형+화살표로 렌더)

- `add_diagram(slide, spec, brand, x, y, w, h)` — deck-spec의 `diagram` 슬라이드를 렌더.
- 대표 레이아웃(전종은 `references/archetype-catalog.md`가 단일 출처):
  - (셰브런 단계 밴드·적층·분기·카드·From→To·그룹 밴드 표는 2026-07-25 폐기 — 사용자 전면 반려.
    단계·분기·병렬 물성의 정본은 (A) 골든 도해다: `deck-compose/references/layout-matching.md` 표 A)
    단계 3~6개. 라벨만 필요하면 `details`를 생략한다 — **별도 어휘를 쓰지 않는다**
    (옛 `flow`는 이것과 렌더 동일이라 2026-07-25 폐기).
  - `layers`: 노드 상→하 적층. 아키텍처 레이어용. 노드 2~5개 권장.
- 노드: `{"label": "굵은 제목", "sub": "설명 한 줄(선택)"}`. 스키마 상세는
  `pptx-build/references/deck-spec-schema.md`.

## 아이콘 (Lucide MIT — 벡터·플랫 픽토그램)

- `add_icon(slide, name, x, y, size, color)` — assets/icons/의 브랜드색 PNG 삽입.
  color: `primary`/`accent`/`white`. bullets·metrics의 `icon` 필드가 이 함수로 렌더된다.
- 아이콘 추가·브랜드 색 변경 시: `scripts/make_icons.py`의 ICONS 목록 수정 후
  `uv run --with pymupdf python make_icons.py` 재실행(브랜드 색은 brand-kit에서 자동 반영).
- AI 이미지 생성 아이콘 금지 — 스타일 비일관·재현 불가.

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
