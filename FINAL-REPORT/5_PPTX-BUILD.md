# 5. PPTX-BUILD — ③ 네이티브 빌더

> deck-spec.json(구성 계약)과 brand-kit.yaml(브랜드 SSOT)을 받아 **네이티브 .pptx**를 찍어내는
> 공장의 본체. 표·차트·다이어그램을 이미지가 아니라 편집 가능한 PowerPoint 객체로 만들고,
> 같은 spec이면 같은 결과가 나오도록 일관성을 코드에 박제한다.
> 본체 코드는 `.claude/skills/pptx-build/scripts/build_pptx.py` 하나이며, 골든 계열 렌더는
> `goldenfab/` 패키지(레이아웃 15종 registry)로 위임한다.

## 5.1 파이프라인 내 위치

```
①GRILL ──▶ ②COMPOSE ──▶ ②.5 승인 ──▶ ┌────────────────────┐ ──▶ ④QA
목차 인터뷰    deck-spec      목업 갤러리   │  ③ PPTX-BUILD  ★  │     계약 대조·오딧
(3장)         (4장)          (4장)        │  spec → .pptx      │     (8장)
                                          └────────────────────┘
```

②가 확정한 deck-spec.json이 유일한 입력 계약이고, 산출된 .pptx는 ④ consistency-qa가
계약(01.5_outline.md)과 3자 대조한다. 수정은 항상 deck-spec/brand-kit 경유 재빌드 —
완성 pptx를 직접 뜯지 않는다(프로젝트 원칙, `CLAUDE.md`).

## 5.2 내부 흐름 — deck-spec → 네이티브 .pptx

```
deck-spec.json ──┐
brand-kit.yaml ──┴─▶ build()  build_pptx.py:1402~1461 (슬라이드 루프 :1433~1460)
                        │
                        ├── PART pre-pass (:1414~1426)
                        │   part 타입 장에서 파트 번호·본문 계층번호(N-M)·소주제 수집
                        │   (파트가 있으면 본문 상단을 NAV_BODY_TOP 2.56으로 내림)
                        │
                        ▼  슬라이드 루프 — type 문자열로 3계열 라우팅
     ┌──────────────────┼──────────────────────────────┐
     │                  │                              │
 레거시 11종        golden.<layout>              adapted.<layout> / novel
 RENDERERS dict    _render_golden (:1336)       _render_scripted (:1367~1400)
 (:1322~1334)           │                              │
     │             registry.LAYOUTS 15종 조회      sl["script"] 상대경로 필수
     │             content_contract               importlib로 장 스크립트 로드
     │             .assert_content (:1344)        build(prs, content) 호출
     │             누락·빈 값 → ValueError        _flatten 자체 적용
     │             빌드 중단(시끄럽게)                 │
     │             fn(prs, content)                    │
     ▼                  │                              │
 마스터 스탬프           └──────────────┬───────────────┘
 _flatten + _frame                     ▼
 (레거시 타입만)              자체 완결 — 푸터 스탬프·배너 미적용
     └────────────┬────────────────────┘
                  ▼
            prs.save(out.pptx)          진입점: python build_pptx.py <spec> <out> [--brand]
```

핵심은 **type 문자열 하나로 3계열이 갈라진다**는 것: 접두어 없는 레거시 타입은 RENDERERS
dict, `golden.` 접두는 goldenfab registry(공장 문), `adapted.`/`novel`은 장 스크립트다.
차트·다이어그램의 실제 그리기는 `visuals.py`·`mpl_exhibits.py`에 위임한다
(build_pptx.py:26~28이 sys.path 주입 후 import — 상세: 6_VISUALS.md).

## 5.3 구조 표

| 구성요소                  | 개수                                    | 출처 파일                                                                                                                                                                                                             |
| ------------------------- | --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 레거시 렌더러             | 11종                                    | `.claude/skills/pptx-build/scripts/build_pptx.py:1322~1334` (RENDERERS)                                                                                                                                               |
| 골든 계열 디스패치        | 3계열(golden.\*/adapted.\*/novel)       | `build_pptx.py:1336`(_render_golden)·`:1367`(_render_scripted)·`:1439`                                                                                                                                                |
| goldenfab 레이아웃        | 15종                                    | `.claude/skills/pptx-build/scripts/goldenfab/registry.py:24~40`                                                                                                                                                       |
| content 계약(명시 키)     | 3종(cover 5키·toc 4키·part 4키)         | `goldenfab/content_contract.py:22~30` (EXPLICIT_KEYS)                                                                                                                                                                 |
| content 계약표 커버리지   | 10타입 150키(15타입 중 — 5종은 표 없음) | `.claude/skills/deck-compose/references/golden-content-contract.md`                                                                                                                                                   |
| 브랜드 색                 | 6종                                     | `.claude/skills/pptx-build/assets/brand-kit.yaml:9-15`                                                                                                                                                                |
| 브랜드 폰트               | 3종(head/body/mono)                     | `brand-kit.yaml:17-20`                                                                                                                                                                                                |
| 글자 크기 단계            | 9단(v4.5 8단 + 미커밋 실험 1)           | `brand-kit.yaml:22-31`                                                                                                                                                                                                |
| brand-kit 코드 로더       | 8종(직접 6 + load_kit 경유 2)           | 직접 파싱: `build_pptx.py:54`·`goldenfab/kit.py:21`·`audit_pptx.py:177`·`make_icons.py:56`·`make_mockups.py:233`·`render_real_mockups.py:39` / `goldenfab.kit.load_kit` 경유: `audit_deck.py:57`·`audit_golden.py:28` |
| brand dict 인자 수신자    | 2종(로드하지 않음)                      | `visuals.py`·`mpl_exhibits.py` — 호출자가 넘긴 brand dict를 받아 쓴다(`visuals.py:5` 도크스트링)                                                                                                                      |
| GRID 좌표 상수(단일 출처) | 전수 표 §5.7                            | `goldenfab/grid.py:6~29`                                                                                                                                                                                              |
| design-rules 규칙         | 126항목(불릿 110 + 번호 16)             | `.claude/skills/pptx-build/references/design-rules.md` 2026-07-19 실측(`grep -c '^\s*- '` 110 · `grep -cE '^\s*[0-9]+\. '` 16). 문서 표기 "96항목"(`docs/user/00-INDEX.md:36`)은 낡음                                 |

## 5.4 슬라이드 타입 라우팅

### 5.4.1 레거시 11종 — RENDERERS dict

`build_pptx.py:1322~1334`의 RENDERERS가 정본이다:

| 타입       | 렌더 특기사항                                                                                             |
| ---------- | --------------------------------------------------------------------------------------------------------- |
| cover      | meta.frame_style=="v44"면 `_cover_v44`로 전환(기본 "v3")                                                  |
| toc        | items 비우면 본문 제목 자동 수집                                                                          |
| section    | 간지 — 워터마크 280pt, hex "2A2E35" 리터럴(:637~640)                                                      |
| bullets    | 최후 수단 타입(덱 전체 2장 이하 — QA 게이트)                                                              |
| two_column | 좌우 2단                                                                                                  |
| metrics    | items **[:4] 절단**(조용한 상한)                                                                          |
| table      | 네이티브 표, 행 높이 0.42~0.68(:1062)                                                                     |
| chart      | 4분기: panels(멀티패널 2~3, [:3] 절단)/sub_table(미니 표)/mpl 이미지 경로/네이티브(visuals.add_chart)     |
| diagram    | layout별 3분기 — layers·branch·cards는 해설 칼럼 좌우 교차, icon_rows 등 4종은 전폭, 그 외 상단 도해+해설 |
| part       | PART pre-pass의 단일 출처, frame_style "v44"면 `_part_v44`                                                |
| cta        | 마무리 — 레거시 골격의 종결 타입                                                                          |

chart의 mpl 판별은 `_is_image_exhibit`: chart_type∈MPL_TYPES 또는 `render:"image"`면 PNG를
종횡비 보존 삽입, 아니면 네이티브 차트+add_callouts 경로다(:1208·1256).

> 낡은 문서 주의: `.claude/agents/pptx-builder.md:17`에는 "스키마의 9종 슬라이드 타입"으로
> 남아 있으나 이는 v3 잔재다. 실측 정본은 스키마(`deck-spec-schema.md:24-36·40-44`)와
> 코드(RENDERERS 11종 + 골든 계열 3종)의 **레거시 11종 + 골든 계열 3종**이다.

### 5.4.2 golden.\<layout> — 공장 문과 content 계약

`_render_golden`(build_pptx.py:1336~1365)은 `golden.` 접두를 떼어낸 레이아웃 키로
`goldenfab/registry.py`의 LAYOUTS 15종을 조회한다:

cover·toc·part(layouts.py) / problem_grid(_variant_k) / exec_graph(s06 variant_c) /
tech_evidence(s08 variant_c) / tech_tree(s09 variant_a) / screenshot(s10 variant_a) /
tech_mechanism(s11 variant_d) / tech_capture(s12 variant_b) / ab_simulation(s14 variant_c) /
validation(s15 variant_c) / mirror_matrix(s17 variant_c — 골든 덱에서는 제외, 실전 창고용) /
boundary(s18 variant_b) / closing(s21 variant_a)

렌더 전에 `content_contract.assert_content`가 강제된다(:1344). 필수 키는
`required_keys()`가 해석한다 — cover·toc·part는 EXPLICIT_KEYS(5·4·4키), 나머지는 레이아웃
모듈의 DEFAULT dict 키 전량. **키 존재는 주입이 아니다** — `None`·빈 문자열·빈 컬렉션은
미주입으로 판정하고, 누락 목록·빈 값 목록·계약에 없는 키(오타 의심)를 담아 ValueError로
빌드를 중단한다(`goldenfab/content_contract.py`). 골든 기본값은 회귀 게이트 전용 기준점이라
"골든의 글과 숫자는 공장 문 밖으로 한 글자도 못 나간다"(gate-repair 스펙)를 이 계약이 지킨다.

part의 `total` 키는 코드상 `c.get("total", 6)`의 조용한 기본값(layouts.py:194)이 있지만,
계약상 필수로 승격돼 있다 — 안 주면 5부 덱에 진행 도트 6개가 찍히는 조용한 오류.

`_render_golden` 안에는 "레이아웃 함수 파라미터가 2개 미만이면 content가 있을 때 에러"
분기(:1356~1360)가 있으나, **registry 15종 전부 `(prs, c)` 서명이라 진입 조건이 성립하지
않는 도달 불가 분기(죽은 분기)다**. 동작하는 코드로 서술하지 않는다.

### 5.4.3 adapted.\<layout> / novel — 장 스크립트 경로

`_render_scripted`(build_pptx.py:1367~1400)가 담당한다. spec의 `sl["script"]`(스펙 기준
상대경로)가 필수이고, importlib로 로드한 모듈의 `build(prs, content)` 함수를 호출한다 —
"장 스크립트가 레이아웃 본체다"(:1381). 반환 slide가 없으면 마지막 슬라이드를 사용하고
`_flatten`(그림자 제거)을 적용한다. adapted는 골든 장 스크립트를 변형한 출발점, novel은
골든 렌더 밀도를 앵커로 한 신규 설계다 — 3갈래 판정 기준은
`.claude/skills/deck-compose/references/layout-matching.md`(상세: 4_DECK-COMPOSE.md),
빌드 후 게이트는 audit_deck+병렬 채점(상세: 8_QA-GATES.md).

## 5.5 마스터 틀 스탬프

`build()`(:1402~1461)의 슬라이드 루프 후반(:1449~1460)이 렌더 후 **레거시 타입에만** 마스터 틀을 얹는다:

```
┌──────────────────────────────────────────────┐
│ [PART 탭 0.2] [킥커 0.62] [타이틀 0.82]       │ ← 파트 내비 헤더
│ ───────────── NAV_RULE_Y 2.42 ─────────────  │   (구형 헤더는 HEADER_HAIR_Y 1.42)
│                                              │
│              (본문 렌더 영역)                 │   BODY_BOTTOM 6.55
│                                              │
│ 출처선 SRC_Y 6.65 · 각주 FOOTNOTE_Y 6.88     │
│ ────────────── HAIR_Y 7.0 ──────────────     │ ← 푸터 헤어라인          세로 저작권
│ 좌 문서명   중앙 워드마크   우 페이지번호 7.08 │ ← _frame                (7pt·270°)
└──────────────────────────────────────────────┘
```

- 순서: `_flatten`(전 도형 그림자 제거) → cover/part/cta 제외 전 장에 `_frame`(푸터
  헤어라인+좌 문서명+중앙 워드마크+우 페이지번호+우측 세로 저작권 7pt rotation 270).
- banner는 cover/toc/part/cta 제외 장에서 source·footnotes와 **배타** 적용(빌더가 배너 우선).
- 헤더는 상단 룰(구형 HEADER_HAIR_Y 1.42 / 파트 내비 NAV_RULE_Y 2.42)과 푸터 헤어라인
  HAIR_Y 7.0의 이중 수평선 구조 — v1의 BCG 실측 개편("세로 rail→헤더 이중룰",
  `CLAUDE.md` 변경 이력)이 상수로 남은 형태다. 캔버스는 13.333×7.5in(:31).
- **golden.\*·adapted.\*·novel 장은 스탬프·배너를 얹지 않는다** — 골든 장은 킥커·헤드라인·
  헤어라인·결론 바·출처선을 자체 포함(자체 완결)하기 때문이다. 좌표도 다르다: 레거시
  출처선은 SRC_Y 6.65, 골든 계열은 GRID의 SOURCE_Y 7.15.

## 5.6 brand-kit.yaml — 브랜드 SSOT

`.claude/skills/pptx-build/assets/brand-kit.yaml` 하나만 고치면 전 덱의 색·폰트·크기가
일괄 반영된다(로더 8종 + brand dict 인자 수신자 2종 실측 — §5.3 표). `check_contract.py`는
brand-kit을 로드하지 않는다 — 금지어 안내 문자열(`check_contract.py:32`)에 파일명만 등장한다.
컨셉은 "흑백 + 포인트 하나(Muted Ember)".

| 축        | 값(실측 정본)                                                                                              |
| --------- | ---------------------------------------------------------------------------------------------------------- |
| 색 6종    | primary #15171B · accent #D66E3A(Muted Ember) · text #1F2329 · muted #8A9099 · bg #FFFFFF · bg_alt #F2F3F5 |
| 폰트 3종  | head/body Pretendard · mono D2Coding — 뷰어 환경 의존(미설치 시 대체 폰트 렌더)                            |
| sizes 9단 | display 40 · title 32 · section 28 · **compact 19** · head 15 · sub 14 · body 12 · caption 9 · foot 8      |
| layout    | margin 0.6in · logo ""(빈 값이면 텍스트 로고)                                                              |

- **compact 19는 미커밋 실험값**이다(git status M — 2026-07-18 s06_proto_f 고밀도 골격
  탐색에서 파생). 백업 `golden/backup/v5-monochrome-2026-07-18/brand-kit.yaml`은 compact
  없는 8단이며, diff 실측 유일 차이가 이 한 줄이다.
- 타이포 위계는 v4.5 4단(McKinsey 실측 대비 ≈3:1): section=위계1(완결 주장문) ·
  sub=위계2 · body=위계3(볼드 리드-인) · caption=위계4.
- **낡음 명시**: `references/reference-metrics.md:17`에는 "brand-kit 적용값
  section=24·sub=15·head=16·body=13"으로 남아 있으나 이는 v3 시절 값이며, 실측 정본은
  현행 yaml(section 28·sub 14·head 15·body 12)이다. `deck-spec-schema.md`의 "subtitle
  15pt"도 현행 sub 14pt와 1pt 어긋난 낡은 수치다. reference-metrics의 BCG NYCHA
  111슬라이드 실측 데이터부(단어/슬라이드 median 197, 텍스트 면적 median 35%)는 밀도
  목표 120~200단어/장의 근거로 유효하다.

로더는 `goldenfab/kit.py`(goldenfab에서 배선 밀도가 가장 높은 모듈): load_kit이 colors를
RGBColor로 변환한 "rgb" 키를 추가하고, 파생색은 `mix`(두 브랜드 색 선형 보간 — 신규 hex
금지 하의 유일 통로), 한글 어절 줄바꿈은 `add_text`가 eaLnBrk="0"+MSO_LANGUAGE_ID.KOREAN
을 코드로 강제(런 언어 태그 없으면 eaLnBrk 무효), 이미지는 `fit_picture`가 PIL로 실비율을
읽어 박스 안에 축소한다(한 축만 주면 비율 가정이 레이아웃을 1.74" 침범한 s10 버그의 근본
수정). 모듈 자체에 hex·pt 리터럴을 두지 않는 것이 규율이다.

## 5.7 GRID 상수 — 골든 좌표 단일 출처

`goldenfab/grid.py:6~29`. "스크립트는 이 값만 파생 사용, 임의 좌표 금지."

| 상수           | 값(inch)     | 의미                               |
| -------------- | ------------ | ---------------------------------- |
| 슬라이드       | 13.333 × 7.5 | 16:9 캔버스(kit.py:17)             |
| MARGIN_L       | 0.6          | 좌 여백                            |
| RIGHT_EDGE     | 12.733       | 우변(13.333−0.6, "12.x 착각 금지") |
| COL_L (X/W)    | 0.6 / 3.1    | 좌 칼럼                            |
| COL_R (X/W)    | 4.2 / 8.533  | 우 칼럼                            |
| V_RULE_X       | 3.95         | 칼럼 사이 세로 룰(파생식)          |
| RULE_Y         | 1.45         | 헤더 룰                            |
| CONTENT_TOP    | 1.75         | 본문 상단                          |
| CONTENT_BOTTOM | 6.35         | 본문 하한(BOX_Y+BOX_H 파생식)      |
| BAR_Y / BAR_H  | 6.60 / 0.45  | 결론 바                            |
| SOURCE_Y       | 7.15         | 출처선                             |

비고: `goldenfab/audit.py`는 grid를 import하지 않고 CONTENT_BOTTOM 6.35·RIGHT_EDGE
12.733을 자체 리터럴로 중복 보유한다 — 값은 일치하나 grid 변경 시 수동 동기화가 필요한
이중 출처(알려진 빚).

## 5.8 design-rules — 규칙 체계 편제

`.claude/skills/pptx-build/references/design-rules.md`는 사용자 승인으로 확정된 디자인
판단의 단일 출처로, **126항목(불릿 110 + 번호 16 — 2026-07-19 실측)**이 3부로 편제돼 있다.
문서 표기 "96항목"(`docs/user/00-INDEX.md:36`)은 갱신되지 않은 낡은 계수다:

| 부  | 범위           | 절                                                                                                 |
| --- | -------------- | -------------------------------------------------------------------------------------------------- |
| Ⅰ부 | 전역 규칙      | §1 GRID · §2 색 · §3 타이포 · §4 리듬 · §5 도형 어휘 · §6 구성 · §7 콘텐츠                         |
| Ⅱ부 | 장 유형별 문법 | A 표지 ~ J 클로징 — A~J 10계열이되 실제 `###` 절은 12개(D-2 문제 정의 장·E-2 관계 도해 포함)       |
| Ⅲ부 | 제작 프로세스  | P1 물성 선언 · P1.5 채점 하강 · P2 오딧 · P3 렌더 루프 · P4 제시 전 셀프 반려 체크 · P5 실전 3갈래 |

**전 장 병렬 채점(장당 1개 에이전트 동시 파견, FAIL 장만 재채점)은 design-rules의 절이 아니라
`pptx-build/SKILL.md:76`의 운용 지시다** — design-rules P4는 "제시 전 셀프 반려 체크"이고,
SKILL이 그 P4를 근거로 병렬 채점 운용을 지시하는 구조다.

이 규칙은 pptx-build SKILL 경유로만 로드된다 — "거치지 않으면 규칙이 적용되지 않는다"
(SKILL description). 산문 규칙 중 기계화 가능한 것은 오딧 러너로 하강시키는 것이 원칙이다
("규칙이 있어도 러너에 안 물리면 없는 규칙", `audit_golden.py:184~185` — 오딧 세부는
8_QA-GATES.md, 골든 장별 문법 세부는 7_GOLDEN-DECK.md에 위임).

낡음 명시: design-rules 절 제목의 S번호("S17 vC"·"S18 vB"·"S21 vC")와 Ⅱ부 C절의
"S3·5·7·13·16·19"는 6부 19장 시절 번호로 남아 있으나, 실측 골든 덱은 5부 17장이다
(규칙 내용은 유효, 번호 좌표만 낡음). 같은 맥락으로 `pptx-build/SKILL.md:78-79`의
"레퍼런스 덱 19장·490도형"도 낡은 서술이며 실측 스냅샷은 17장·519도형이다.

## 5.9 실증 예시(walked example) — K-IFRS 파일럿 slides[14]가 빌더를 통과하는 과정

실물 spec `_workspace_kifrs/02_deck-spec.json`(16장, meta.frame_style="v3")의
`slides[14]`를 추적한다. 실값:

```json
{ "type": "golden.ab_simulation", "content": { "kicker": "5. 핵심 의사결정" } }
```

**① 라우팅.** `build()`의 슬라이드 루프가 type 문자열을 본다. `golden.` 접두 →
`_render_golden`(build_pptx.py:1336). 접두를 뗀 키 `ab_simulation`으로
`registry.LAYOUTS`를 조회하면 `s14_variants.variant_c`(registry.py:24~40)가 나온다.
서명은 `(prs, c)` 파라미터 2개 — "파라미터 <2" 분기(:1356~1360)는 지나가지 않는다
(15종 전부 2개 서명이라 애초에 도달 불가).

**② 계약 검사.** 렌더 전에 `content_contract.assert_content`(:1344)가 실행된다.
`required_keys("ab_simulation")`은 EXPLICIT_KEYS에 없으므로 s14 모듈의 DEFAULT dict
키 전량 — 계약표 기준 **15키**다. 이 slide의 주입은 kicker 1개뿐이므로 **현행 코드에서는
나머지 14키 누락 목록을 담은 ValueError로 빌드가 즉시 중단된다.** 이것이 정상 동작이다 —
"빈 장이 아니라 그럴듯하게 채워진 장이 나오므로 아무도 못 본다 … 시끄러운 게 옳다"
(`content_contract.py:6~10`).

**③ 반면교사(과거형).** 2026-07-15 계약 신설 전에는 `{**DEFAULT, **(c or {})}` 침묵
폴백이 누락 키를 골든 기본값(다른 프로젝트의 글)으로 조용히 메웠고, 이 spec 그대로
빌드가 통과했다. 그 결과가 파일럿 QA FAIL 17건 중 골든오염 9장이다(현재 동결 상태 —
상세: 9_KIFRS-PILOT.md). kicker override "5. 핵심 의사결정"은 골든 기본
"4. 문제와 해결"을 덮은 부분 주입의 실물 흔적으로, 지금은 "부분 주입도 FAIL"의 표본이다.

**④ 15키 전량 주입 시의 렌더·스탬프.** 계약을 통과하면 `fn(prs, content)`로
variant_c가 호출되고, 장이 자체 완결로 그려진다: GRID 상수 파생 좌표 위에 킥커+헤드라인+
헤어라인(CONTENT_TOP 1.75 위쪽), 본문 도해, 결론 바(BAR_Y 6.60·높이 0.45 — 구 S17
7축 비교의 주장을 흡수한 그 바), 출처선(SOURCE_Y 7.15). 마지막 스탬프 단계에서
`build()`는 이 장이 golden.\* 이므로 **푸터 스탬프·배너를 얹지 않는다**. 동일 레이아웃이
골든 기본값으로 렌더될 때의 도형 규모는 스냅샷 S14(트러블슈팅) **37도형**이 기준점이다
(`assets/golden-snapshot.json` — 17장 총 519도형의 한 장).

**⑤ 같은 spec의 다른 장.** 간지 5장은 전부 `total=5`를 명시했다 — layouts.py:194의
조용한 기본값 6(5부 덱에 도트 6개)을 회피하는, 계약이 total을 필수로 승격시킨 이유의
실물이다. 반면 slides[3]·[5]·[7]~[10]·closing은 content 키 자체가 없어, 현행 코드로
이 spec을 다시 빌드하면 첫 골든 본문 장(S4)에서 즉시 중단된다 — 계약 이행 0/8의 물증이
그대로 동결돼 있다.

## 5.10 알려진 빚·엣지 케이스 (build_pptx.py 실측)

| 항목                                 | 상태                                                                        |
| ------------------------------------ | --------------------------------------------------------------------------- |
| `_render_golden` "파라미터 <2" 분기  | 도달 불가(죽은 분기) — registry 15종 전부 (prs, c) 서명                     |
| 간지 워터마크 hex "2A2E35"(:637~640) | goldenfab "리터럴 금지" 규율 밖의 레거시 하드코딩                           |
| metrics items[:4]·chart panels[:3]   | 조용한 절단 상한(에러 없이 잘림)                                            |
| `main()` 반환 없음                   | `sys.exit(main())`이 항상 exit 0 — 실패는 예외로만 전파                     |
| `_intro` 반환 오프셋 1.0 고정        | 텍스트 길이 무관 고정값                                                     |
| 폰트 뷰어 의존                       | Pretendard·D2Coding 미설치 환경에서 대체 폰트 렌더                          |
| 실전 완주 산출물 부재                | 이 빌더로 QA까지 통과한 .pptx 완성본은 아직 없다(파일럿 FAIL 후 동결 — 9장) |

## 교차 링크

- deck-spec 구성·레이아웃 매칭(golden/adapted/novel 판정): 4_DECK-COMPOSE.md
- 차트 6종·mpl 9종·다이어그램 18종·아이콘: 6_VISUALS.md
- 골든 덱 17장·goldenfab 레이아웃 세부·스냅샷: 7_GOLDEN-DECK.md
- compare_golden·audit_deck·audit_golden·check_contract 게이트 전 층: 8_QA-GATES.md
- K-IFRS 파일럿 FAIL 17건·동결 상태: 9_KIFRS-PILOT.md
