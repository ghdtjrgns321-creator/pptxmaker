# 6. 시각 어휘 24종 — `pptx-visuals` (⑤)

```
 ① 재료 ─ ② 질문 후보 ─ ③ 갤러리 ─ ④ 선택 ─ ▶ ⑤ 조판 ◀ ─ ⑥ 눈검증 ─ ⑦ 채점
```

## 내부 흐름

```
 "이 주장을 무엇으로 증명하나"
        │
        ├─ 정량 · 데이터 형태로 판별 ──▶ 네이티브 차트 6종      편집 가능
        │   (기본값 bar 금지)              visuals.add_chart()
        │
        ├─ 네이티브가 못 만드는 물성 ──▶ mpl 익스히빗 9종       PNG(편집 불가)
        │   (증분·순위역전·전후·분포)      mpl_exhibits.render()
        │
        └─ 정성 · 구조 ───────────────▶ 도형 DSL 9종            편집 가능
            (시간축·좌표계·충족도·겹침)   visuals.add_diagram()
```

**시각은 장식이 아니라 주장의 증거 형태다.** 슬라이드마다 "이 장의 주장은 무엇인가"를 한
문장으로 쓰고, 그 주장을 무엇으로 증명하는지에 따라 형식을 고른다.

## 호출 규약

```python
visuals.add_chart(slide, spec, brand, x, y, w, h)      # 네이티브 — 수치 편집 가능
visuals.add_diagram(slide, spec, brand, x, y, w, h)    # 도형 DSL
visuals.add_callouts(slide, annotations, brand, x, y, w, h)
visuals.add_icon(slide, name, x, y, size, color="primary"|"accent"|"white")
mpl_exhibits.render(spec, brand, out_png, w_in, h_in)  # PNG 경로 반환
```

**자리는 호출자가 준다**(인치). 이 모듈에 캔버스 상수를 두지 않는다. `spec`은 평범한 dict —
빌더도 spec 파일도 필요 없다. 필드 정의는 `references/spec-fields.md`가 단일 출처.

## 어휘 목록

### 네이티브 차트 6종 (`CHART_TYPES`)

| 데이터 형태                         | 타입               |
| ----------------------------------- | ------------------ |
| 항목 크기 비교, 카테고리명 짧음     | `bar`              |
| 순위·비중 비교, 카테고리명 김       | `hbar`             |
| 시계열 추세                         | `line`             |
| 구성비(합≈100%, 항목 ≤5)            | `pie` · `doughnut` |
| 구성이 여러 그룹에 걸쳐 변하는 비교 | `stacked_bar`      |

**기본값 bar 금지** — 데이터 형태를 먼저 판별한다.

### mpl 익스히빗 9종 (`MPL_TYPES`)

| 타입                | 데이터 필드                              | 쓸 자리             |
| ------------------- | ---------------------------------------- | ------------------- |
| `waterfall`         | `categories` + `values` (+`total_label`) | 증분 누적·구성 분해 |
| `heatmap`           | `rows` + `cols` + `values`(2D)           | 행×열 밀도          |
| `dumbbell`          | `categories` + `series`(2개)             | 전후 비교           |
| `slope`             | `columns`(2) + `series{이름:[v1,v2]}`    | 두 시점 순위 역전   |
| `funnel`            | `stages` + `values`                      | 단계 전환·이탈      |
| `annotated_scatter` | `points[{x,y,label,emphasis?}]`          | 상관 + 예외         |
| `histogram`         | `values` (+`bins`)                       | 분포                |
| `bubble`            | `items[{label,value}]` 4~9개             | 원 면적 비교        |
| `waffle`            | `items[{label,percent}]` 1~3개           | 10×10 도트 비율     |

산출물은 220dpi PNG라 PPT에서 **수치 편집이 불가**하다. 고치려면 데이터를 바꿔 다시 부르고
PNG를 갈아끼운다.

### 도형 DSL 9종 (`add_diagram`)

`timeline` · `matrix_2x2` · `spectrum` · `harvey_table` · `check_matrix` · `venn` ·
`icon_rows` · `stat_split` · `split_detail`

`layout`은 **필수**다. 누락·오타는 `ValueError`로 빌드 시점에 죽는다.

## 조용한 폴백을 두지 않는 이유

기본값을 두면 **어휘가 수렴한다.** "어휘를 안 고르면 박스+화살표"가 실제로 일어났고, 그 결과
박스+화살표 계열 9종이 전수 폐기됐다.

**폐기 9종**(쓰면 `ValueError`): `flow` · `process_band` · `layers` · `branch` · `cards` ·
`from_to` · `band_table` · `pro_con` · `contrast_split`

사유는 하나다 — 팔레트 전수 렌더에서 **전면 반려**("그냥 도형 나열 / 고등학생 ppt"). 노드
문법을 개선한 판까지 반려됐고 실전 사용은 0건이었다(2026-07-25 사용자 확정).

## BCG 스타일 단일 경로

모든 mpl 렌더러는 `_fig()`로 시작해 `_save()`로 끝나고, 색은 `_palette(brand)`가 주는 것만
쓴다.

```
 전부 회색(C7CBD1)  +  강조만 accent 1색  +  콜아웃(annotations)  +  각주(note)
```

네이티브 차트도 같은 규칙을 `emphasis` 필드로 따른다. 콜아웃은 `add_callouts()`가 pptx
도형(편집 가능)으로 얹는다.

## 실증 예시 — 렌더러 24종 단독 호출 스모크

2026-07-29, 이 모듈이 아카이브된 빌더 없이 단독으로 도는지 확인했다.

```
단독 호출 성공 19 / 24
  mpl 8종      waterfall · heatmap · dumbbell · slope · funnel · scatter · histogram · bubble
  네이티브 6종  bar · hbar · line · pie · doughnut · stacked_bar
  도형 5종      timeline · spectrum · harvey_table · check_matrix · stat_split

실패 5건      waffle(KeyError 'percent') · venn/matrix_2x2(TypeError) ·
              icon_rows(KeyError 'lead') · split_detail(KeyError 'visual')
```

**실패 5건은 전부 필드명 오류이지 결합 문제가 아니었다.** 스모크를 짜면서 필드를 추측했고,
실제 스키마(`x_axis`는 리스트, waffle은 `percent`)와 달랐다. 문서가 맞았고 호출이 틀렸다.

이 결과가 판단을 바꿨다. 처음에는 이 모듈이 `deck-spec` 파이프라인에 묶여 있다고 보고 아카이브
대상으로 분류했는데, 인자가 `(slide, spec, brand, x, y, w, h)`라 **자리를 인자로 받고 데이터
dict를 받는** 구조였다. 아카이브한 `figures/`처럼 골든 좌표를 기본값으로 물고 있지도 않았다.
2차 실험에서 안 쓴 건 코드 결함이 아니라 그냥 안 썼기 때문이었다.

실물 품질도 확인했다 — `slope`는 두 시점의 **순위 역전을 선 교차로 인코딩**하고, 회색 베이스에
강조 계열만 accent로 칠한다. 숫자를 나열하는 대신 관계를 그리는 어휘다.

## 배선 상태 (정직하게)

| 모듈 | 하네스·조판 스크립트에서 import한 곳 | 상태 |
| --- | --- | --- |
| `visuals.py` | **0곳** | 구현·단독 호출은 확인됨. **실전 덱에서 쓴 적 없음** |
| `mpl_exhibits.py` | **0곳** | 같음 |
| `make_icons.py` | 0곳 | 독립 CLI(아이콘 재생성용) — 정상 |

2차 실험판 17장은 이 모듈들을 **한 번도 부르지 않았다.** 표·대시보드·도해를 전부 `kit`·`grid`·
`dense`로 직접 그렸기 때문이다. 즉 이 장이 서술하는 24종 어휘는 **쓸 수 있는 상태**이지
**쓰이고 있는 상태**가 아니다.

이 구분이 중요한 이유는 [11_JOURNEY](11_JOURNEY.md)에 있다 — 옛 하네스에서 13,710줄 중 700줄만
값을 하고 나머지가 0회 호출이었던 것이 폐기 사유였다. 같은 기준을 이 모듈에도 적용하면
"아직 증명되지 않은 자산"이다. 아카이브하지 않은 근거는 단독 호출 스모크 19/24와,
숫자 나열 대신 관계를 인코딩하는 어휘라는 판단이다.

## 확인된 결함

- `dumbbell` 범례가 마지막 행 데이터점에 겹친다(행 3개일 때 실측).
- 위 실패 5종은 필드를 정확히 맞춰야 돈다 — `spec-fields.md` 표를 보고 호출한다.

## 아이콘

`add_icon`이 `pptx-build/assets/icons/`의 브랜드색 PNG를 삽입한다. Lucide MIT 벡터를 브랜드
3색으로 구운 73장. 추가·색 변경 시 `make_icons.py`의 ICONS 목록을 고치고 재실행한다
(`uv run --with pymupdf python make_icons.py`).

**AI 이미지 생성 아이콘 금지** — 스타일이 일관되지 않고 재현이 불가능하다.
