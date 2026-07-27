# 5. PPTX-BUILD — ③ 네이티브 빌더

> deck-spec.json(구성 계약)과 brand-kit.yaml(브랜드 SSOT)을 받아 **네이티브 .pptx**를 찍어내는
> 공장의 본체. 표·차트·다이어그램을 이미지가 아니라 편집 가능한 PowerPoint 객체로 만들고,
> 같은 spec이면 같은 결과가 나오도록 일관성을 코드에 박제한다.
> 본체는 `.claude/skills/pptx-build/scripts/build_pptx.py`(1,300줄) 하나이고, 골든 계열 렌더와
> 부품 조립은 `goldenfab/` 패키지로 위임한다.

## 5.1 파이프라인 내 위치

```
①GRILL ──▶ ②COMPOSE ──▶ ②.5 승인 ──▶ ┌────────────────────┐ ──▶ ④QA
목차 인터뷰    deck-spec      목업 갤러리   │  ③ PPTX-BUILD  ★  │     계약 대조·오딧
(3장)         (4장)          (4장)        │  spec → .pptx      │     (8장)
                                          └────────────────────┘
```

②가 확정한 deck-spec.json이 유일한 입력 계약이고, 산출된 .pptx는 ④ consistency-qa가 계약(`01.5_outline.md`)과 3자 대조한다. 수정은 항상 deck-spec/brand-kit 경유 재빌드 — 완성 pptx를 직접 뜯지 않는다(`CLAUDE.md` 원칙).

## 5.2 내부 흐름 — deck-spec → 네이티브 .pptx

```
deck-spec.json ──┐
brand-kit.yaml ──┴─▶ build()
                        │
                        ▼  슬라이드 루프 — type 문자열로 4계열 라우팅
   ┌──────────────┬──────────────┬───────────────────┬───────────────────┐
   │              │              │                   │                   │
golden.<layout>  composed    adapted.<layout>      레거시 8종
_render_golden  _render_       / novel             RENDERERS dict
   │            composed      _render_scripted        │
   │              │              │                   │
registry.LAYOUTS  frames 6종 +   sl["script"] 필수    section·bullets·
15종 조회         figures 부품   importlib 로드       two_column·metrics·
   │              │              build(prs, content)  table·chart·
content_contract  measure로      │                    diagram·cta
.assert_content   자리 조이기    assert_scripted_      │
   │              │              content              ▼
누락·빈 값 →      채움률 <96%    │                  마스터 스탬프
ValueError 중단   → ValueError   │                  _flatten + _frame
   │              중단           │                  (+ banner/source/footnotes)
   ▼              ▼              ▼                    ▼
   └──────────────┴──────────────┴────────────────────┘
                        ▼
              _gate_density — 빌드가 스스로 채점(generic_checks)
                        ▼
                  prs.save(out.pptx)
```

핵심은 **type 문자열 하나로 4계열이 갈라진다**는 것이다. 앞의 셋은 자체 완결이라 푸터 스탬프·배너를 얹지 않고, 레거시만 마스터 틀을 받는다. 차트·다이어그램의 실제 그리기는 `visuals.py`·`mpl_exhibits.py`에 위임한다(상세: [6_VISUALS.md](6_VISUALS.md)).

## 5.3 구조 표

| 구성요소              | 개수·값                                      | 출처                                          |
| --------------------- | -------------------------------------------- | --------------------------------------------- |
| 장 타입 계열          | **4계열**                                    | `build_pptx.py` 슬라이드 루프                 |
| 레거시 렌더러         | **8종**                                      | `build_pptx.py` RENDERERS                     |
| goldenfab 레이아웃    | **15종**                                     | `goldenfab/registry.py` LAYOUTS               |
| 배치 틀               | **6종**                                      | `goldenfab/frames.py` FRAMES                  |
| 부품(도해)            | **10종**                                     | `goldenfab/figures/` · `select.PARTS`         |
| content 계약(명시 키) | 3종(cover 5키·toc 4키·part 4키)              | `goldenfab/content_contract.py` EXPLICIT_KEYS |
| 브랜드 색·폰트·크기   | 색 6 · 폰트 3 · 크기 **8단**                 | `assets/brand-kit.yaml`                       |
| brand-kit 소비 코드   | **32파일**(로더·경유 포함)                   | `grep -rln "load_kit(\|brand-kit"` 실측       |
| GRID 좌표 상수        | 전수 §5.7                                    | `goldenfab/grid.py`(84줄)                     |
| 공유 컴포넌트         | dense 9함수(`hero_card`·`compact_header` 등) | `goldenfab/dense.py`(196줄)                   |
| design-rules 규칙     | **148항목**(불릿 124 + 번호 24)              | `references/design-rules.md`(630줄) 실측      |

## 5.4 타입 라우팅 — 4계열

### 5.4.1 `composed` — 배치 틀 + 부품 조립 (2026-07-27 신설)

이 경로가 실전 덱의 기본이다. spec은 이렇게 생겼다.

```json
{"type": "composed", "frame": "twin_top_cards",
 "kicker": "...", "headline": "...", "source": "...",
 "figures": [{"part": "bipartite_map", "data": {...}},
             {"part": "relation_catalog", "data": {...}}],
 "cards": {"part": "card_row", "data": {"cards": [...]}}}
```

**왜 생겼나.** 골든 장 통째 배정(`golden.<layout>`)이 실전 덱을 골든 복제로 만드는 기계적 원인이었다 — 실전 덱의 타입 분포가 골든 `SLIDE_ORDER`와 그대로 겹쳤다. 이 경로는 자리는 틀이, 그림은 부품이 맡아 **장 전체를 가져오지 않는다.**

**렌더 순서**: 틀 이름 검증 → 부품 모듈 로드 → 각 부품의 `measure()`로 필요 크기 계산 → 틀이 자리(`slots`)를 배분 → 배경·헤더 → 소제목 밴드 → 부품 `draw()` → 카드 띠 → 출처선 → `_flatten`.

**자리를 조인다.** 틀이 주는 밴드는 **상한**이고 실제 높이는 부품이 안다. 안 조이면 부품이 안 쓴 자리가 그대로 빈칸으로 남는다 — 첫 조립 시험에서 도해가 2.20" 자리에 1.00"만 써서 가운데가 통째로 비었다. 부품은 자리에 비례해 늘어나지 않으므로(계약 ①) 자리를 줄이는 쪽이 맞고, 남는 만큼 카드가 위로 올라온다. **카드를 두껍게 하지는 않는다** — 2차 시험에서 두껍게 했더니 불릿이 벌어지고 아이콘 둘레가 비어 카드 안이 헐거워졌다. 카드 높이는 글자·겹 수에 묶인 치수다.

**채움률 게이트.** 부품과 틀이 각자 정상이어도 **조합이 얕으면** 장이 비어 보인다. 그 판정을 눈에 맡기지 않는다 — 조립 결과가 본문 구간의 96%(`FILL_MIN`, 골든 실측 최소) 미만이면 ValueError로 중단하고, 무엇을 하라는지까지 메시지에 적는다("밴드가 얕은 틀로 바꾸거나, 부품에 정보 층을 얹거나, 장을 합친다").

**모르는 이름은 조용히 넘기지 않는다.** 틀·부품 이름의 단일 출처는 `frames.FRAMES`·`select.PARTS`이고, 없는 이름이면 등록 목록을 붙여 죽는다. 조용한 폴백이 어휘를 3종으로 수렴시킨 원인이었다.

### 5.4.2 `golden.<layout>` — 공장 문과 content 계약

`_render_golden`이 `golden.` 접두를 뗀 키로 `registry.LAYOUTS` 15종을 조회한다. 렌더 전에 `content_contract.assert_content`가 강제된다. 필수 키는 `required_keys()`가 해석한다 — cover·toc·part는 `EXPLICIT_KEYS`, 나머지 12종은 레이아웃 모듈의 DEFAULT dict 키 전량.

**키 존재는 주입이 아니다** — `None`·빈 문자열·빈 컬렉션은 미주입으로 판정하고, 누락 목록·빈 값 목록·계약에 없는 키(오타 의심)를 담아 ValueError로 빌드를 중단한다. 골든 기본값은 회귀 게이트 전용 기준점이라 **골든의 글과 숫자는 공장 문 밖으로 한 글자도 못 나간다.**

### 5.4.3 `adapted.<layout>` / `novel` — 장 스크립트

`_render_scripted`가 담당한다. spec의 `sl["script"]`(spec 기준 상대경로)가 필수이고, importlib로 로드한 모듈의 `build(prs, content)`를 호출한다 — "장 스크립트가 레이아웃 본체다".

**이 문에는 2026-07-25까지 계약 검사가 없었다.** 골든 출발 스크립트를 그대로 쓰면 K-IFRS 글이 남아도 전 게이트가 green이었다. 지금은 `assert_scripted_content`가 "그 스크립트가 실제로 읽는 골든 기본값 키"를 요구한다. DEFAULT 전량이 아닌 이유는 실측이다 — 정본 dense 장이 주입해도 안 읽는 키가 22건이라 전량을 요구하면 즉시 상시 적색이 된다. 그 구멍 수는 `test_wiring` 통합F가 감소 전용 래칫으로 고정한다.

### 5.4.4 레거시 8종 — RENDERERS dict

`section` · `bullets` · `two_column` · `metrics` · `table` · `chart` · `diagram` · `cta`.

| 타입       | 렌더 특기사항                                                                  |
| ---------- | ------------------------------------------------------------------------------ |
| section    | 간지 — 워터마크 280pt, hex "2A2E35" 리터럴(goldenfab 리터럴 금지 규율 밖의 빚) |
| bullets    | 최후 수단 타입(덱 전체 2장 이하 — QA 게이트)                                   |
| two_column | 좌우 2단                                                                       |
| metrics    | items **[:4] 절단**(조용한 상한)                                               |
| table      | 네이티브 표, 행 높이 0.42~0.68                                                 |
| chart      | 4분기: 멀티패널(2~3, [:3] 절단) / 미니 표 / mpl 이미지 경로 / 네이티브 차트    |
| diagram    | layout별 분기 — `visuals.add_diagram` 9종에 위임                               |
| cta        | 마무리 — 레거시 골격의 종결 타입                                               |

**`cover`·`toc`·`part`는 2026-07-25에 폐기됐다.** 골든 프레임(`golden.cover`/`golden.toc`/`golden.part`)이 정본인데 이쪽은 v3 시절 사본이라 "어느 쪽이 표준인가"가 갈려 있었고, 구세대 견본 spec이 그 타입을 쓰고 있었다. 지금 이 타입을 쓰면 unknown slide type ValueError로 죽는다.

**부수 효과 — 도달 불가가 된 내비 기계.** `build()`의 PART pre-pass(간지 수집·본문 계층번호 N-M 부여·`NAV_BODY_TOP`)는 legacy `part` 렌더러를 전제로 하는데 그 타입이 폐기돼 **현재 발화하지 않는다.** 골든 덱은 `golden.part`가 자체 간지·탭을 그린다. 지울지 살릴지는 "legacy 경로를 어디까지 유지하나"라는 별개 결정이라 코드에 주석으로 명시만 해 두었다 — 안 적으면 다음 사람이 살아 있는 기계로 착각하고 손댄다.

## 5.5 마스터 틀 스탬프 — 레거시 전용

```
┌──────────────────────────────────────────────┐
│              (본문 렌더 영역)                 │   BODY_BOTTOM 6.55
│                                              │
│ 출처선 SRC_Y 6.65 · 각주 FOOTNOTE_Y 6.88     │
│ ────────────── HAIR_Y 7.0 ──────────────     │ ← 푸터 헤어라인          세로 저작권
│ 좌 문서명   중앙 워드마크   우 페이지번호 7.08 │ ← _frame                (7pt·270°)
└──────────────────────────────────────────────┘
```

- 순서: `_flatten`(전 도형 그림자 제거) → cta 제외 전 장에 `_frame`.
- `banner`는 `source`·`footnotes`와 **배타** 적용(빌더가 배너 우선) — 하단 스트립을 같이 쓰면 겹친다.
- **`composed`·`golden.*`·`adapted.*`·`novel` 장은 스탬프·배너를 얹지 않는다** — 킥커·헤드라인·헤어라인·결론 바·출처선을 자체 포함하기 때문이다. 좌표도 다르다: 레거시 출처선은 SRC_Y 6.65, 골든 계열은 GRID의 SOURCE_Y 7.15.
- 캔버스는 13.333×7.5in(16:9).

## 5.6 brand-kit.yaml — 브랜드 SSOT

이 파일 하나만 고치면 전 덱의 색·폰트·크기가 일괄 반영된다. 컨셉은 "흑백 + 포인트 하나(Muted Ember)".

| 축        | 값(실측 정본)                                                                                              |
| --------- | ---------------------------------------------------------------------------------------------------------- |
| 색 6종    | primary #15171B · accent #D66E3A(Muted Ember) · text #1F2329 · muted #8A9099 · bg #FFFFFF · bg_alt #F2F3F5 |
| 폰트 3종  | head/body Pretendard · mono D2Coding — 뷰어 환경 의존(미설치 시 대체 폰트 렌더)                            |
| sizes 8단 | display 40 · title 32 · section 28 · head 15 · sub 14 · body 12 · caption 9 · foot 8                       |
| layout    | margin 0.6in · logo ""(빈 값이면 텍스트 로고)                                                              |

타이포 위계는 4단(McKinsey 실측 대비 ≈3:1): section=위계1(완결 주장문) · sub=위계2 · body=위계3(볼드 리드-인) · caption=위계4.

로더 정본은 `goldenfab/kit.py`(144줄)다. `load_kit`이 colors를 RGBColor로 변환한 `rgb` 키를 추가하고, 파생색은 `mix`(두 브랜드 색 선형 보간 — 신규 hex 금지 하의 유일 통로), 한글 어절 줄바꿈은 `add_text`가 `eaLnBrk="0"` + `MSO_LANGUAGE_ID.KOREAN`을 코드로 강제하며(런 언어 태그 없으면 eaLnBrk 무효), 이미지는 `fit_picture`가 실비율을 읽어 박스 안에 축소한다(한 축만 주면 비율 가정이 레이아웃을 1.74" 침범한 s10 버그의 근본 수정). 모듈 자체에 hex·pt 리터럴을 두지 않는 것이 규율이고, 그 규율은 `test_wiring` 통합D가 **하드코딩 래칫**으로 지킨다 — 현재 visuals 0 · build_pptx 0 · mpl 2(승인 예외).

## 5.7 GRID 상수 — 골든 좌표 단일 출처

`goldenfab/grid.py`. "스크립트는 이 값만 파생 사용, 임의 좌표 금지."

| 상수           | 값(inch)     | 의미                               |
| -------------- | ------------ | ---------------------------------- |
| 슬라이드       | 13.333 × 7.5 | 16:9 캔버스                        |
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
| CAP_PITCH      | 0.42         | 세로 피치 상한 = 골든 리듬         |

`CAP_PITCH`가 이 표에서 성격이 다르다 — **자리가 남아도 이보다 벌어지지 않는다.** 빈 자리를 균등 분배하면 골든의 리듬이 깨지고 장이 헐거워지기 때문이다.

배치 틀(`frames.py`)의 좌표는 GRID 파생이 아니라 **골든 실측**이다: `CARD_Y 3.62 · CARD_H 3.32`(S8·S9·S11 공통), `BODY_TOP 1.05 ~ BODY_BOTTOM 6.60`, 틀마다 다른 `foot`(골든 장별 실제 본문 하단), `FILL_MIN 0.96`.

## 5.8 design-rules — 규칙 체계 편제

`references/design-rules.md`는 사용자 승인으로 확정된 디자인 판단의 단일 출처로, **148항목(불릿 124 + 번호 24)**이 P0 + 3부로 편제돼 있다.

| 부  | 범위           | 절                                                                                                                                  |
| --- | -------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| P0  | 북극성         | 이 장이 왜 존재하나 — 설계 **전** 필수                                                                                              |
| Ⅰ부 | 전역 규칙      | §1 GRID · §2 색 · §3 타이포 · §4 리듬 · §5 도형 어휘 · §6 구성 · §7 콘텐츠 · **§8 고밀도 규율**                                     |
| Ⅱ부 | 장 유형별 문법 | A 표지 · B 목차 · C 간지 · D 본문공통 · D-2 문제정의 · E 기술설명 · E-2 관계 · F 스크린샷 · G 전환비교 · H 미러 · I 경계 · J 클로징 |
| Ⅲ부 | 제작 프로세스  | P1 물성 선언 · P1.5 채점 하강 · P2 오딧 · P3 렌더 루프 · P4 셀프 반려 체크 · P5 실전 3갈래                                          |

§8(고밀도 개편 규율)이 이 문서에서 가장 최근에 자란 절이고, 공유 헬퍼 `goldenfab/dense.py`가 그 코드 대응물이다.

이 규칙은 pptx-build SKILL 경유로만 로드된다 — "거치지 않으면 규칙이 적용되지 않는다". 산문 규칙 중 기계화 가능한 것은 오딧 러너로 하강시키는 것이 원칙이고("규칙이 있어도 러너에 안 물리면 없는 규칙"), 그 하강 여부를 `test_wiring`이 검사한다.

**낡음 명시**: Ⅱ부 절 제목의 S번호("S17 vC"·"S18 vB"·"S21 vC", C절의 "S3·5·7·13·16·19")는 6부 19장 시절 좌표로 남아 있다. 규칙 내용은 유효하고 번호만 낡았다.

## 5.9 빌드 시 자체 채점 — `_gate_density`

빌드가 **스스로** `goldenfab.audit.generic_checks`로 채점한다. 오케스트레이터가 ④ consistency-qa를 안 돌려도 성김·겹침·넘침이 잡힌다.

| 장 계열                   | 프로파일 | 밀도 밴드        |
| ------------------------- | -------- | ---------------- |
| `golden.*`(정형 4종 제외) | sparse   | 골든 스냅샷 파생 |
| `adapted.*` · `novel`     | dense    | 골든 스냅샷 파생 |
| 레거시 8종                | legacy   | **미적용**       |
| `composed`                | —        | **대상 아님**    |

레거시는 2026-07-25까지 이 게이트의 대상이 아니었고 — 남은 레거시 어휘(차트 6·mpl 9·표·문서형·다이어그램 9종)가 무검증으로 나갔다 — 그때 배선됐다. 밀도 밴드는 골든 파생이라 레거시에 걸지 않는다(93/111 상시 적색이 된다); 레거시 밀도의 정본은 `audit_pptx`의 본문 60단어 하한이다.

**`composed`는 이 게이트에 안 걸린다** — 프로파일 배정 분기 어디에도 안 맞아 건너뛴다. 대신 자체 채움률 게이트(96%)가 있다. 두 검사가 보는 것이 다르므로(채움률은 세로 점유, generic_checks는 겹침·성김·accent 상한) 이건 커버리지 갭이다(`상세: 8_QA-GATES.md`).

## 5.10 알려진 빚·엣지 케이스

| 항목                               | 상태                                                                            |
| ---------------------------------- | ------------------------------------------------------------------------------- |
| PART pre-pass·내비 기계            | 도달 불가 — legacy `part` 폐기로 발화 안 함(코드 주석에 명시)                   |
| `composed`가 `_gate_density` 밖    | 채움률만 보고 generic_checks는 안 돈다                                          |
| `composed`를 ④ QA·계약이 모른다    | `audit_pptx`·`check_contract`·`content_contract` 참조 0 — 완주 시 여기서 막힌다 |
| 간지 워터마크 hex "2A2E35"         | goldenfab "리터럴 금지" 규율 밖의 레거시 하드코딩                               |
| metrics items[:4]·chart panels[:3] | 조용한 절단 상한(에러 없이 잘림)                                                |
| 폰트 뷰어 의존                     | Pretendard·D2Coding 미설치 환경에서 대체 폰트 렌더                              |
| 실전 완주 산출물 부재              | 이 빌더로 ④까지 통과한 .pptx 완성본은 아직 없다                                 |

## 교차 링크

- deck-spec 구성·부품/틀 선택 규칙: [4_DECK-COMPOSE.md](4_DECK-COMPOSE.md)
- 부품 10종·틀 6종의 실체, 차트·다이어그램 어휘: [6_VISUALS.md](6_VISUALS.md)
- 골든 덱 17장·dense 승격·스냅샷: [7_GOLDEN-DECK.md](7_GOLDEN-DECK.md)
- compare_golden·audit_deck·audit_golden·check_contract·test_wiring: [8_QA-GATES.md](8_QA-GATES.md)
