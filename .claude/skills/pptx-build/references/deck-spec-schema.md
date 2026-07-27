# deck-spec.json 스키마 (에이전트 간 데이터 계약)

`deck-composer`가 산출하고 `pptx-build`가 소비하는 단일 계약. 이 스키마를 벗어난 키는
빌더가 무시하거나 실패한다. 슬라이드 **골격**(앞뒤 고정, 본문 유연)은 여기서 강제된다.

## 최상위 구조

```json
{
  "meta": { "project": "프로젝트명", "tagline": "한 줄 소개", "audience": "대상(선택)" },
  "slides": [ { "type": "...", ... }, ... ]
}
```

## 골격 규칙 (오케스트레이터가 검증)

- **맨 앞 고정:** `cover` → `toc` 로 시작
- **맨 뒤 고정:** `cta` 로 종료
- **본문(유연):** 그 사이는 프로젝트 성격에 맞게 `bullets/two_column/table/chart/diagram/metrics/section`을 자유 배치
- `toc`는 items를 비워두면 본문 섹션 제목으로 자동 채워진다 → 목차/본문 순서가 항상 일치

## 슬라이드 타입별 필드

| type         | 필수 필드                                                                                                                                                                         | 용도                                                         |
| ------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------ |
| `cover`      | `title`, `subtitle`                                                                                                                                                               | 표지                                                         |
| `toc`        | (없음) `items` 선택                                                                                                                                                               | 목차. 비우면 본문 제목 자동                                  |
| `section`    | `title`                                                                                                                                                                           | 섹션 구분(전환)                                              |
| `bullets`    | `title`, `bullets[]`                                                                                                                                                              | 불릿 3~5개. Pain/특징 등                                     |
| `two_column` | `title`, `left{heading,items[]}`, `right{heading,items[]}`                                                                                                                        | 문제/솔루션·Before/After 대비                                |
| `metrics`    | `title`, `items[{value,label}]`                                                                                                                                                   | 큰 숫자 KPI 최대 4개                                         |
| `table`      | `title`, `headers[]`, `rows[[]]`                                                                                                                                                  | 네이티브 표(기술스택·비교표)                                 |
| `chart`      | `title`, `chart_type`(bar/hbar/line/pie/doughnut/stacked_bar + 확장 waterfall/heatmap/dumbbell/slope/funnel/annotated_scatter/histogram), `categories[]`, `series{name:[values]}` | 차트. 기본 6종은 네이티브, 확장 7종은 mpl 이미지(하이브리드) |
| `diagram`    | `title`, `layout`(timeline/matrix_2x2/venn/icon_rows … **필수**, 9종), `nodes[{label,sub?}]`                                                                                      | 도형 다이어그램(흐름·아키텍처)                               |
| `part`       | `title`, `subtitle`(선택)                                                                                                                                                         | 간지(PART divider)                                           |
| `cta`        | `title`, `contact{email,site,...}`                                                                                                                                                | 마무리·연락처                                                |

## 골든 계열 타입 (2026-07-16 — 골든은 변형 가능한 출발점)

| type               | 필수 필드                        | 용도                                                                                                                      |
| ------------------ | -------------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `golden.<layout>`  | `content{...}` (DEFAULT 키 전량) | 내용 구조가 골든과 그대로 맞을 때만 — 텍스트 전량 교체, 좌표·색·도형 고정. 키 누락 시 빌드 중단(content_contract)         |
| `adapted.<layout>` | `script`(spec 기준 상대경로)     | 물성은 매칭되나 항목 수·구획이 다를 때 — 그 골든을 출발점으로 goldenfab 부품을 재사용한 장 스크립트 `build(prs, content)` |
| `novel`            | `script`                         | 골든에 없는 물성 — 가장 가까운 골든 렌더를 밀도 앵커로 신규 설계(design-rules Ⅲ부 P5)                                     |

- 골든 계열 장은 헤더·결론바·출처를 자체 포함 — 레거시 푸터 스탬프·배너가 붙지 않는다.
- `adapted.*`·`novel`의 선택 필드 `"audit": {"dup_allow": N, "known": {"규칙명": 기준선}}` —
  audit_deck.py의 장별 예외 선언(재등장이 메시지인 노드의 실측 수 등).
- 판정 기준은 deck-compose `references/layout-matching.md`, 게이트는 audit_deck + 병렬 채점.

## PART 내비게이션 체계 (우석진 템플릿 이식 — _workspace/06_reference-notes.md)

`part` 슬라이드를 본문 그룹 앞에 끼우면 빌더가 자동으로 처리한다:

- **간지**: primary 전면 배경 + "PART 0N" 칩 + "Ⅰ. 제목" + 해당 PART 소주제 프리뷰(넘버 서클)
- **상단 섹션 탭**: 모든 본문 슬라이드에 Ⅰ~Ⅳ 탭 스탬프, 현재 PART만 하이라이트(bg_alt+accent 밑줄)
- **계층 번호**: 본문 제목에 "N-M." 자동 접두 + 제목 위 소분류 kicker("N. PART명")
- **그룹 목차**: toc가 PART별 2단 목차로 자동 전환
- PART는 3~4개 권장. `part` 없이 쓰면 구형(평면) 레이아웃으로 렌더된다.

### `table` 매트릭스 스타일

`"style": "matrix"` — 비교 매트릭스: 마지막 열(제안 솔루션)을 accent 헤더 + 굵은 글씨로 강조,
데이터 열 중앙 정렬. 비교 우위 열은 반드시 마지막에 배치한다.

### `diagram` 상세 (pptx-visuals 스킬이 렌더)

- `layout`은 **필수**다(기본값 없음 — 누락 시 빌드가 ValueError로 죽는다). 기본값 자체가
  디폴트 어휘 수렴의 원인이었다("어휘를 안 고르면 박스+화살표").
- **폐기 9종(2026-07-25 사용자 확정)**: `flow`·`process_band`·`layers`·`branch`·`cards`·
  `from_to`·`band_table`·`pro_con`·`contrast_split`. 박스+화살표 계열 전수 — 팔레트 전수 렌더
  전면 반려, 실전 사용 0건. **이 값을 쓰면 빌드가 ValueError로 죽는다.** 단계·분기·병렬·전환·
  그룹 나열·2열 대비의 정본은 (A) 골든 도해다(`deck-compose/references/layout-matching.md` 표 A:
  G01·G04·G09·G10·G11·G14·G08 · 병렬 카드는 `dense.hero_card`).
- `layout: "timeline"` — 좌→우 마일스톤(넘버 원 + 상단 라벨/하단 설명). 로드맵·연혁. 3~5개.
- `chart`의 `panels[{title, chart_type, categories, series}]` — 미니 차트 2~3개 나란히(멀티패널).
  지표 여러 개의 전후 개선을 한 장에 담을 때. panels 사용 시 commentary 대신 intro 권장.
- `banner`: 본문 슬라이드 하단 전폭 결론 스트립(한 문장, `**강조**` 지원).
  **source·footnotes와 병용 금지**(같은 영역 점유 — 빌더가 배너 우선).
- 형식 선택은 `pptx-visuals/references/visual-selection.md` 결정표를 따른다(주장 한 문장 → 형식).
- `layout: "matrix_2x2"` (v4.2) — 2×2 사분면 포지셔닝. `x_axis:["좌","우"]`, `y_axis:["하","상"]`,
  `quadrants:[4라벨(선택)]`, `items:[{label, x:0~1, y:0~1, emphasis?}]`. 정성 경쟁 비교용.
- `layout: "spectrum"` (v4.2) — 성숙도 스펙트럼(점진 음영 밴드 + 현위치 accent 마커).
  `stages:[{label, sub?}]` 4~6개, `marker: <0-base 현위치>`. 수준·단계 위의 위치 주장용.
- `layout: "harvey_table"` (v4.2) — 하비볼 비교표. `cols[]`, `rows:[{label, scores:[0~4]}]`,
  `emphasis_col: <0-base>`. 대안별 충족도(다축 우월) 정성 비교용.
- `layout: "check_matrix"` (v4.2) — 체크매트릭스(✓/✗/–). harvey_table과 동일 형식,
  scores에 "y"/"n"/"-". 요건 충족 비교용.
- v4.4 채굴 승격 4종:
  - `chart_type:"bubble"` (mpl 이미지) — `items:[{label, value, emphasis?}]` 4~9개. 원 면적 비교.
  - `chart_type:"waffle"` (mpl 이미지) — `items:[{label, percent, emphasis?}]` 1~3개. 10×10 도트 %.
  - `layout:"venn"` (도형) — `sets:[{label, sub?}]` 2~3개 + `overlap:"교집합 라벨"`. 개념 겹침.
- ~~`meta.frame_style: "v3"|"v44"`~~ 폐기 2026-07-25 — legacy `cover`·`part` 렌더러 전용
  스위치였고 그 두 프레임이 골든 중복으로 폐기됐다. 표지·목차·간지는 `golden.cover`/
  `golden.toc`/`golden.part`가 정본이며 스타일 스위치가 없다(스냅샷으로 박제).
- 아키타입 L-ID·형상 매핑·세트 제약은 `pptx-visuals/references/archetype-catalog.md`가 단일 출처.
- 노드: `{"label": "굵은 제목", "sub": "설명 한 줄(선택)"}`. `sub`는 생략 가능.

```json
{"type":"diagram","title":"도입 이정표","layout":"timeline",
 "nodes":[{"label":"1분기","sub":"파일럿"},{"label":"2분기","sub":"전사 확대"},
          {"label":"3분기","sub":"운영 이관"}]}
```

## v4 하이브리드 익스히빗 필드 (chart 타입 확장)

- **렌더 경로 자동 판별**: `chart_type`이 확장 7종(waterfall·heatmap·dumbbell·slope·funnel·
  annotated_scatter·histogram)이면 mpl 220dpi PNG로 삽입된다(**수치편집불가** — 수정은
  deck-spec 재빌드). 기본 6종은 네이티브 유지(수치편집가능). `"render": "image"`로 강제 가능.
- `emphasis: ["항목명"]` — **전부 회색 + 강조만 accent 1색** BCG 규칙. 네이티브(카테고리/
  시리즈명 매칭)·mpl 공통. 없으면 기존 팔레트 순환.
- `annotations: [{"text": "...", ...}]` — 콜아웃 주석.
  - 네이티브 경로: `{"at": [fx, fy], "point_to": [fx, fy]?}` — 개체 영역 분수 좌표에 pptx
    도형(편집가능)으로 얹는다.
  - mpl 경로: `{"xy": [데이터x, 데이터y]?, "at": [fx, fy]?}` — 이미지 안에 굽는다.
- `ref: "ref/dallas_p18.png"` — 레퍼런스 앵커(ref/catalog.md). 목업 갤러리에 표시되고
  레이아웃 기준이 된다. 빌더는 무시(스타일은 코드가 보장).
- `sub_table: {"headers":[..], "rows":[[..]]}` (v4.1, BCG p18) — chart 슬라이드 차트 아래
  부속 데이터 행(네이티브 미니 표, 1~3행). "한 장 안의 조합"(차트+부속표)로 밀도를 만든다.
- 확장 유형별 데이터 필드: waterfall `categories+values(+total_label)` · heatmap
  `rows+cols+values[2D]` · dumbbell `categories+series(2개)` · slope `columns(2)+series{name:[v1,v2]}`
  · funnel `stages+values` · annotated_scatter `points[{x,y,label,emphasis?}](+x_label,y_label)`
  · histogram `values(+bins,x_label)`. mpl 경로 출처는 `note` 필드(이미지 내 각주).

## 공장 표준 골격 (박제 — 문제→솔루션→증거)

확정된 본문 배치. 앞뒤 고정(cover→toc … cta), 본문 5단은 이 순서가 기본:

문제(bullets) → 솔루션(two_column) → 핵심기능(bullets/table) → 차별점(table/bullets)
→ 성과(metrics/chart). 순서 자체가 설득 장치다.

성과 숫자가 없으면 성과 슬라이드는 정성 강점으로 대체하거나 뺀다(문제→솔루션 순서는 불변).

## 밀도·구성 (NYCHA BCG 실측 반영 — references/reference-metrics.md)

기관 문서 밀도를 목표로 한다(슬라이드당 120~200 단어). 지원 필드:

- `subtitle`: 제목 아래 설명 한 줄(15pt). 대부분의 본문 슬라이드 권장.
- `commentary`: **table/chart/diagram/metrics 필수 권장** — 개체 옆(표·차트·layers는 좌측
  칼럼)이나 아래(process_band·metrics는 전폭 블록)에 붙는 분석 불릿(중첩 지원, bullets와 동일 형식).
  개체만 덜렁 있는 슬라이드는 밀도 미달(QA density 검사 FAIL 60단어 하한)이 된다.
  "차트가 말하는 결론 + 근거 + 시사점"을 3~5개 불릿으로 쓴다.
- **중첩 불릿**: `bullets`/`two_column` items를 문자열 대신 객체로 — `{"text":"리드 문장","sub":["근거1","근거2"]}`.
  최상위 `•` + 하위 `–`(들여쓰기). 그룹 블록으로 밀도를 만든다.
- `footnotes`: `["단서·참조", ...]` → 하단에 7~8pt 각주.
- `icon`: `bullets` 항목·`metrics` 항목에 `"icon": "shield-check"` — 넘버 서클 대신 브랜드색
  픽토그램(Lucide) 렌더. 사용 가능 이름은 `pptx-visuals/scripts/make_icons.py`의 ICONS 목록
  (assets/icons/에 사전 변환된 24종). 없는 이름은 빌드 에러.
- `intro`: bullets/table/chart/diagram 슬라이드의 **전폭 서술 문단**(2~4문장, `**강조**` 지원).
  거버닝 메시지 아래·개체 위에 배치 — 컨설팅 보고서식 리드. 재료(보고서)의 배경 서술을 쓴다.
- 해설 칼럼 좌/우는 빌더가 **자동 교차**한다(연속 슬라이드 동일 배치 방지) — spec에서 지정 불가.
- `source`: 정량 슬라이드 좌하단 출처.

```json
{"type":"bullets","title":"문제","subtitle":"수기 프로세스가 만드는 비용",
 "bullets":[
   {"text":"주문 접수가 **분산**돼 있다","sub":["이메일·전화 **40%**","폼·팩스 잔존"]},
   "오입력 반품 **월 12건**"],
 "footnotes":["2026-Q2 내부 집계 기준"]}
```

## 인라인 강조 (핵심 데이터 — BCG식)

`bullets`와 `two_column` items 안에서 `**...**`로 감싼 부분은 **볼드 + accent색**으로 렌더된다.
핵심 수치·차별 키워드만 감싼다(문장 전체 금지 — 강조가 강조가 아니게 됨). 실제 기관 덱은
숫자를 본문 안에 인라인으로 강조하지, 큰 색 숫자 카드를 쓰지 않는다.

```json
{"type":"bullets","title":"문제","bullets":["전체 주문의 **40%**가 수기 입력","반품 **월 12건**"]}
```

## 마스터 틀 (자동 스탬프 — 컨설팅 표준)

모든 본문 슬라이드에 코드가 동일 위치로 자동 삽입한다(spec에 안 적어도 됨):
- **헤더:** 제목 아래 얇은 구분선
- **푸터:** 좌 `회사명 · 프로젝트명`, 우하단 페이지번호 `NN / MM`, 위에 헤어라인
- **표지(cover)·마무리(cta)는 페이지번호·푸터 생략** (브랜드 마크만)

선택 필드 `"source"`: 표/차트/metrics 등 정량 슬라이드에 넣으면 좌하단에 `출처: …`로 렌더된다.
정량 주장에는 출처를 붙이는 것이 셀링 신뢰의 기본이다.

```json
{"type":"chart","title":"성과","source":"내부 로그 2026-Q2", "...":"..."}
```

## 주의

- 표·차트는 **네이티브 객체**로 생성된다(이미지 아님). 값은 spec의 숫자를 그대로 넣어라 —
  이미지로 굽지 말 것. 데이터만 고치면 PPT에서 자동 반영되는 게 핵심 가치다.
- `bullets`는 슬라이드당 5개 이하. 넘치면 슬라이드를 나눠라(정보 밀도 과다는 셀링 역효과).
