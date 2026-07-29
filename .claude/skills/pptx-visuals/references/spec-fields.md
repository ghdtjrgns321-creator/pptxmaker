# 렌더러 데이터 필드 — visuals.py · mpl_exhibits.py 사용법 단일 출처

`deck-spec-schema.md`(deck-spec 파이프라인과 함께 2026-07-29 아카이브)에서 차트·다이어그램
필드만 발췌해 옮긴 것. **여기 적힌 것이 이 스킬 렌더러의 전부다** — 빌더·spec 파일 없이
파이썬에서 직접 호출한다.

## 호출 규약

```python
visuals.add_chart(slide, spec, brand, x, y, w, h)      # 네이티브(수치편집가능)
visuals.add_diagram(slide, spec, brand, x, y, w, h)    # 도형 DSL
visuals.add_callouts(slide, annotations, brand, x, y, w, h)
visuals.add_icon(slide, name, x, y, size, color="primary"|"accent"|"white")
mpl_exhibits.render(spec, brand, out_png, w_in, h_in)  # PNG 경로 반환(수치편집불가)
```

- 자리(`x,y,w,h`)는 인치, **호출자가 준다** — 렌더러는 캔버스 상수를 갖지 않는다.
- `brand`는 `goldenfab.kit.load_kit()` 반환값(`brand["colors"]["accent"]` 등).
- `spec`은 평범한 dict다. 아래 필드 외에 무엇이 더 들어있어도 무시된다.

## 네이티브 차트 — `chart_type` 6종

`bar` · `hbar` · `line` · `pie` · `doughnut` · `stacked_bar`

```python
{"chart_type": "hbar",
 "categories": ["상한 밀림", "검색 실패", "생성 누락"],
 "series": {"누락 원인 비중 (%)": [45, 33, 18]},   # dict — 이름: 값 배열
 "emphasis": ["상한 밀림"]}                        # 이것만 accent, 나머지 회색
```

타입 선택은 `visual-selection.md` 결정표를 따른다. **기본값 bar 금지** — 데이터 형태를 먼저 판별.

## mpl 익스히빗 — `chart_type` 9종 (PNG로 굽는다)

| chart_type          | 데이터 필드                                            | 쓸 자리             |
| ------------------- | ------------------------------------------------------ | ------------------- |
| `waterfall`         | `categories` + `values` (+`total_label`)               | 증분 누적·구성 분해 |
| `heatmap`           | `rows` + `cols` + `values`(2D)                         | 행×열 밀도          |
| `dumbbell`          | `categories` + `series`(2개)                           | 전후 비교           |
| `slope`             | `columns`(2) + `series{이름:[v1,v2]}`                  | 두 시점 순위 역전   |
| `funnel`            | `stages` + `values`                                    | 단계 전환·이탈      |
| `annotated_scatter` | `points[{x,y,label,emphasis?}]` (+`x_label`,`y_label`) | 상관 + 예외         |
| `histogram`         | `values` (+`bins`,`x_label`)                           | 분포                |
| `bubble`            | `items[{label,value,emphasis?}]` 4~9개                 | 원 면적 비교        |
| `waffle`            | `items[{label,percent,emphasis?}]` 1~3개               | 10×10 도트 비율     |

## 도형 DSL — `layout` 9종

`layout`은 **필수**다. 누락·오타는 `ValueError`로 빌드 시점에 죽는다 — 기본값을 두지 않는 이유는
디폴트 어휘 수렴("어휘를 안 고르면 박스+화살표")이 실제로 일어났기 때문이다.

| layout         | 필드                                                                                             | 쓸 자리               |
| -------------- | ------------------------------------------------------------------------------------------------ | --------------------- |
| `timeline`     | `nodes[{label, sub?}]` 3~5개                                                                     | 좌→우 마일스톤        |
| `matrix_2x2`   | `x_axis[좌,우]` · `y_axis[하,상]` · `quadrants[4라벨]?` · `items[{label,x:0~1,y:0~1,emphasis?}]` | 2×2 포지셔닝          |
| `spectrum`     | `stages[{label, sub?}]` 4~6개 · `marker`(0-base 현위치)                                          | 성숙도 위 현위치      |
| `harvey_table` | `cols[]` · `rows[{label, scores:[0~4]}]` · `emphasis_col?`                                       | 대안별 충족도         |
| `check_matrix` | 위와 동일, `scores`에 `"y"`/`"n"`/`"-"`                                                          | 요건 충족 비교        |
| `venn`         | `sets[{label, sub?}]` 2~3개 · `overlap`(문자열 라벨)                                             | 개념 겹침             |
| `icon_rows`    | `rows[{icon?, badge?, lead, desc, tag?}]` 3~6개 · `tag_head?`                                    | 정성 나열의 기본 그릇 |
| `stat_split`   | `stat{value, label, desc?}` · `rows[icon_rows 행]`                                               | 빅넘버 + 근거 행      |
| `split_detail` | `visual`(중첩 diagram spec) · `rows[icon_rows 행]` · `tag_head?`                                 | 좌 도해 + 우 설명     |

`icon_rows`의 `badge`는 배지 자리에 **수량**을 넣는 옵션 — 합이 분모와 맞는 전수 귀속("떨어진
것이 각각 왜인가")을 표현한다. 없으면 순번(i+1)이 들어가는데, 순번은 귀속 수량을 말할 수 없다.

### 폐기 어휘 — 쓰면 ValueError

`flow` · `process_band` · `layers` · `branch` · `cards` · `from_to` · `band_table` · `pro_con` ·
`contrast_split`. 박스+화살표 계열 전수, 2026-07-25 사용자 확정 폐기("그냥 도형 나열 / 고등학생
ppt"). 실전 사용 0건이었다. 이 물성들이 필요하면 도형을 직접 그린다.

## 공통 필드

- `emphasis: ["항목명"]` — **전부 회색 + 강조만 accent 1색**. 네이티브·mpl 공통. 없으면 팔레트 순환.
- `annotations: [{"text": ...}]` — 콜아웃.
  - 네이티브: `{"at": [fx, fy], "point_to": [fx, fy]?}` — 분수 좌표, pptx 도형(편집가능)으로 얹음
  - mpl: `{"xy": [데이터x, 데이터y]?, "at": [fx, fy]?}` — 이미지 안에 굽힘
- `note` — mpl 경로 각주(이미지 내부).

## 확인된 결함

- `dumbbell` 범례가 마지막 행 데이터점에 겹친다(2026-07-29 스모크). 행 3개 이하일 때 확인 필요.
