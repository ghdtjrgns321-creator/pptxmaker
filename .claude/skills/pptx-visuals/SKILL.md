---
name: pptx-visuals
description: PPTX 안의 네이티브 차트(bar/line/pie 6종)·mpl 익스히빗(waterfall/slope/dumbbell 등 9종)·도형 다이어그램(timeline/matrix_2x2/venn 등 9종)을 생성하는 레시피 단일 출처. 슬라이드에 차트나 도해를 넣을 때, 또는 "차트 수치 바꿔/다이어그램 수정/그래프 추가/시각자료 손봐" 요청 시 반드시 사용.
---

# pptx-visuals — 차트·다이어그램 단일 출처

시각자료 코드는 **이 모듈에만 존재한다** — 다른 곳에 차트 코드를 중복 작성하지 않는다.

구현 2개: `scripts/visuals.py`(네이티브 차트 + 도형 DSL + 아이콘) · `scripts/mpl_exhibits.py`
(네이티브가 못 만드는 9종을 220dpi PNG로). 데이터 필드는 **`references/spec-fields.md`가 단일
출처**다.

## 호출

```python
visuals.add_chart(slide, spec, brand, x, y, w, h)      # 네이티브 — 수치편집가능
visuals.add_diagram(slide, spec, brand, x, y, w, h)    # 도형 DSL
visuals.add_callouts(slide, annotations, brand, x, y, w, h)
visuals.add_icon(slide, name, x, y, size, color="primary"|"accent"|"white")
mpl_exhibits.render(spec, brand, out_png, w_in, h_in)  # PNG 경로 반환 — 수치편집불가
```

**자리는 호출자가 준다**(인치). 이 모듈에 캔버스 상수를 두지 않는다. `brand`는
`goldenfab.kit.load_kit()` 반환값. `spec`은 평범한 dict — 빌더도 spec 파일도 필요 없다.

## 형식 선택

**시각은 주장의 증거 형태다**(설명+분석+설득 목적, 장식 아님). 주장 한 문장 → 결정표 → 형식.
결정표와 안티패턴은 `references/visual-selection.md`가 단일 출처.

정량 차트는 **데이터 형태를 먼저 판별**한다 — 기본값 bar 금지:

| 데이터 형태                         | 타입               |
| ----------------------------------- | ------------------ |
| 항목 크기 비교, 카테고리명 짧음     | `bar`              |
| 순위·비중 비교, 카테고리명 김       | `hbar`             |
| 시계열 추세                         | `line`             |
| 구성비(합≈100%, 항목 ≤5)            | `pie` / `doughnut` |
| 구성이 여러 그룹에 걸쳐 변하는 비교 | `stacked_bar`      |

네이티브가 못 하는 물성은 mpl 9종으로: 증분 분해 `waterfall` · 밀도 매트릭스 `heatmap` ·
전후 비교 `dumbbell` · 순위 역전 `slope` · 단계 이탈 `funnel` · 상관+예외 `annotated_scatter` ·
분포 `histogram` · 면적 비교 `bubble` · 비율 도트 `waffle`.

아키타입 L-ID·형상 매핑·세트 제약은 `references/archetype-catalog.md`.

## BCG 스타일 단일 경로

모든 mpl 렌더러는 `_fig()`→`_save()` 경유 + `_palette()` 색만 쓴다 — 전부 회색(C7CBD1) +
**강조만 accent 1색** + 콜아웃(`annotations`) + 각주(`note`). 네이티브 차트도 같은 규칙을
`emphasis` 필드로 따른다. 콜아웃은 `add_callouts()`가 pptx 도형(편집가능)으로 얹는다.

## 도형 DSL — layout은 필수

`layout` 누락·오타는 `ValueError`로 **빌드 시점에 시끄럽게** 죽는다. 조용한 폴백을 두지 않는
이유는 그게 어휘 수렴의 기계적 원인이었기 때문이다("어휘를 안 고르면 박스+화살표").

살아있는 9종: `timeline` · `matrix_2x2` · `spectrum` · `harvey_table` · `check_matrix` · `venn` ·
`icon_rows` · `stat_split` · `split_detail`.

**폐기 9종**(쓰면 ValueError): `flow` · `process_band` · `layers` · `branch` · `cards` ·
`from_to` · `band_table` · `pro_con` · `contrast_split`. 박스+화살표 계열 전수, 2026-07-25 사용자
확정 폐기("그냥 도형 나열 / 고등학생 ppt"). 실전 사용 0건이었다.

## 아이콘 (Lucide MIT — 벡터·플랫 픽토그램)

`add_icon`이 `pptx-build/assets/icons/`의 브랜드색 PNG를 삽입한다. 아이콘 추가·브랜드 색 변경 시
`scripts/make_icons.py`의 ICONS 목록을 고치고 `uv run --with pymupdf python make_icons.py` 재실행
(브랜드 색은 brand-kit에서 자동 반영). **AI 이미지 생성 아이콘 금지** — 스타일 비일관·재현 불가.

## 확장 시

새 타입은 `visuals.py`(또는 `mpl_exhibits.py`)에 구현하고 `references/spec-fields.md`에 필드를
등록한 뒤 `visual-selection.md`에 사용 기준을 한 줄 추가한다(3곳 동기화). 카탈로그 행은
**렌더러 실존 보증**이다 — 구현·스모크 후에만 등재한다.

## 확인된 결함

- `dumbbell` 범례가 마지막 행 데이터점에 겹친다(2026-07-29 스모크, 행 3개).
- `waffle`·`venn`·`matrix_2x2`·`icon_rows`·`split_detail`은 필드를 정확히 맞춰야 돈다 —
  `spec-fields.md` 표를 보고 호출할 것(임의 추측 시 KeyError/TypeError).
