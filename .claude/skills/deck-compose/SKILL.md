---
name: deck-compose
description: FINAL-REPORT 재료를 아웃라인 계약에 따라 B2B 셀링 골격에 배치해 deck-spec.json을 만든다. 창작이 아니라 선별·배치가 역할. deck-composer 에이전트가 슬라이드를 구성할 때, 또는 "슬라이드 구성 다시/스토리 수정/특정 슬라이드만 재작성/카피 손봐" 요청 시 반드시 사용.
---

# deck-compose — 초안 통합·골격 배치

> **골든 레이아웃 배정(Phase 2~):** 각 장의 `type`은 물성 기준으로 `golden.<layout>`을 우선
> 배정한다 — 판정은 [references/layout-matching.md](references/layout-matching.md)의 15타입
> 결정표. 어느 타입에도 안 맞으면 끼워맞추지 말고 골든 확장 절차(pull)로 올린다.

FINAL-REPORT 재료를 `01.5_outline.md`(아웃라인 계약)의 목차·강조점에 따라 셀링 덱으로
**배치**한다. 콘텐츠는 사용자가 FINAL-REPORT로 이미 확정했다 — 여기서는 재료를
**선별·배열**만 한다. 산출물은 `_workspace/02_deck-spec.json`.

## 구성 원칙

1. **아웃라인 계약을 지킨다.** `01.5_outline.md`의 장 구성·채록 메시지가 계약이다 —
   장을 임의로 추가·삭제하거나 채록된 강조점을 바꾸지 않는다.
2. **창작 금지.** FINAL-REPORT에 없는 성과·기능을 지어내지 않는다. 셀링은 신뢰가 자산이다.
3. **숫자로 말한다.** FINAL-REPORT의 수치는 네이티브 `chart`로, 임팩트 숫자 2~4개는 `metrics`로.
   본문 핵심 수치는 `**...**` 인라인 강조(문장 전체 금지).
4. **한 메시지, 빽빽한 근거.** 슬라이드당 결론 하나 + **120~200 단어**(BCG 실측).
   부제 + 중첩 불릿(`{"text":리드,"sub":[근거…]}`) + 각주로 밀도를 만든다.
5. **개체만 덜렁 금지.** table/chart/diagram/metrics 슬라이드에는 `commentary`(분석 불릿
   3~5개)를 반드시 붙인다 — "이 차트/표가 말하는 결론·근거·시사점". QA가 본문 슬라이드
   60단어 미만을 FAIL로 잡는다.
6. **서술 문단으로 전문성.** FINAL-REPORT의 배경 서술은 `intro`(2~4문장 전폭 문단)로 살린다 —
   불릿만 있는 페이지보다 보고서 문체가 컨설팅 톤을 만든다. 배경 서술이 얇으면 억지로
   창작하지 말고 오케스트레이터에 "재료 부족(GRILL에서 보강 요청)"을 보고.

## 공장 표준 골격 — 문제→솔루션→증거 (박제)

**앞뒤는 고정, 본문 5단은 이 순서 기본, 내용만 프로젝트별 유연.**

```
[고정]  1 cover        표지: 프로젝트명 + 한 줄 가치제안
[고정]  2 toc          목차 (본문 자동)
[본문]  3 문제(Pain)    · bullets        — 누가 무엇 때문에 곤란한가 3~5
[본문]  4 솔루션        · two_column/diagram — 문제→해결 (As-Is/To-Be, 처리 흐름)
[본문]  5 핵심 기능     · bullets/table  — 실제로 얻는 것 3~4
[본문]  6 차별점        · table/bullets  — 왜 하필 이것(경쟁 비교)
[본문]  7 성과·증거     · metrics/chart  — 근거 있는 숫자 (있을 때만)
[고정]  8 cta          마무리 + 연락처
```

**고정 규칙:** `cover`→`toc` 시작, `cta` 종료. 성과는 CTA 직전. 문제→솔루션 순서 불변.
**유연 규칙:** 성과 숫자가 없으면 7번을 정성 강점으로 대체하거나 뺀다. 전체 6~10장 권장.
표준 예시: `assets/standard-deck-spec.json`.

## PART 골격 (필수 — 내비게이션 체계)

본문을 **PART 3~4개**로 묶고 각 그룹 앞에 `part` 간지를 끼운다(예: Ⅰ 문제 정의 → Ⅱ 솔루션
→ Ⅲ 차별점·증거 → Ⅳ 로드맵). 빌더가 상단 탭·계층번호(N-M)·그룹 목차를 자동 스탬프한다.
심사자가 "지금 어디를 보고 있는지" 알게 하는 장치다.

## 타입 선택 가이드 — 시각은 장식이 아니라 주장의 증거 형태

**슬라이드마다 주장 한 문장을 먼저 쓰고, `pptx-visuals/references/visual-selection.md`의
결정표(Zelazny 5비교·Abela 4목적 기반)로 형식을 고른다.** 안티패턴 표(이종 재고를 bar로,
병렬 목록을 layers로 등)에 걸리면 즉시 교체. 같은 형식 3장 연속이면 재점검.

- 프로세스·단계 흐름 → `diagram` flow / 방어·아키텍처 적층 → `diagram` layers
- 1→N 분기(라우팅·선택) → `diagram` branch / 로드맵·연혁 → `diagram` timeline
- 경쟁·전후 비교 → `table`(비교 우위가 핵심이면 `"style":"matrix"`)
- 수치 → `chart`: **데이터 형태를 먼저 판별**하고 타입을 고른다(기본값 bar 금지) —
  크기 비교=bar, 카테고리명 긴 순위·비중=hbar, 시계열=line, 구성비(합≈100%)=pie/doughnut,
  그룹별 구성 변화=stacked_bar. 상세 결정표는 pptx-visuals SKILL
- **확장 정량 유형(v4·mpl 이미지)**: 기여도 분해=waterfall, 전후 비교=dumbbell/slope(표 금지),
  교차 밀도=heatmap, 전환·이탈=funnel, 상관+예외=annotated_scatter, 분포=histogram.
  네이티브 차트엔 `emphasis`(회색+강조 1색)와 `annotations`(콜아웃)를 적극 사용
- 임팩트 숫자 2~4개 → `metrics` / 대비 서술 → `two_column`
- `bullets`는 **위 어느 도형에도 안 걸릴 때만** — 덱 전체에서 2장 이하로 제한
- bullets 리드·metrics 항목에는 의미가 맞는 `icon`(Lucide 24종, make_icons.py ICONS 목록)을
  붙여 픽토그램 모티프를 만든다 — 의미가 안 맞으면 억지로 붙이지 말고 넘버 서클 유지

## 익스히빗 사양서 패스 (v4.2 — 아키타입 배정 + 사용자 승인 게이트)

deck-spec 확정 전에 시각 슬라이드마다 후보를 만들어 사용자 승인을 받는다. 다양성 판단을
모델 디폴트에 맡기지 않는 장치다. **아키타입 단일 출처는
`pptx-visuals/references/archetype-catalog.md`** (L01~L30 + 형상 매핑 + 세트 제약).

1. **형상 분류 먼저**: 슬라이드마다 `shape`(시계열/범주비교/구성비/흐름전환/전후/교차다축/
   분포상관/정성/시간여정/숫자)를 분류하고, 카탈로그의 형상→L-ID 매핑에서 후보군을 고른다.
   다양성은 장식이 아니라 데이터 형상에서 나온다 — 정성 서술은 L22~L25(2×2·스펙트럼·
   하비볼·체크매트릭스)로 시각 구조화한다.
2. 산출: `_workspace/03_exhibit-candidates.json` —
   `{"slides":[{"no","title","message","shape","current","visual_candidates":[...]}]}`.
   **슬라이드당 3안**(1·2안 상이 유형, 3안째는 같은 유형의 조합 변주 허용). 모든 후보에
   조합 장치(emphasis/annotations/sub_table/banner/ref) ≥1. **맨 현행 재출품 금지.**
   디폴트 어휘(bar·flow·cards) 후보에는 `why_not`(검토한 대안 L-ID와 기각 사유) 의무.
3. **자기 심사(보고 전 필수)**: `uv run python
   .claude/skills/pptx-visuals/scripts/check_candidates.py <후보.json> <deck-spec.json>`
   — RESULT: PASS가 나올 때까지 고친다. 히스토그램에서 단일 유형이 슬라이드의 30%를
   넘으면 해당 슬라이드를 다른 아키타입으로 재설계.
4. 렌더(실물): `uv run python .claude/skills/pptx-visuals/scripts/render_real_mockups.py
   _workspace/03_exhibit-candidates.json _workspace/mockups` → 인터랙티브 gallery.html
   (라디오 선택 + 실시간 히스토그램·게이트 판정 + 선택 복사). "sNN=X" 회신 대기.
   **회신 전 deck-spec 확정 금지.** (COM 불가 환경 폴백: make_mockups.py 스케치)

## 다양성 강제 조항 (consistency-qa 게이트 4종이 기계 검증 — 위반이면 되돌아온다)

- 동일 유형 **간격 ≥3장(쿨다운)** · 동일 유형 **덱 전체 ≤2회** · 덱 전체 **최소 5종** ·
  **박스 다이어그램(flow/layers/cards/branch/from_to) 본문의 30% 이하**.
  후보 설계 단계에서 게이트 통과 가능한 조합이 존재하게 구성한다(check_candidates가 시뮬레이션).

## 출력: `_workspace/02_deck-spec.json`

`pptx-build`의 `references/deck-spec-schema.md` 계약을 정확히 따른다. 표·차트 값은 숫자로.
스키마에 없는 필드·타입을 만들지 않는다. 이미지 폴백 재료를 쓸 땐 그 사실을 오케스트레이터에
보고한다(일관성 예외이므로).

## 자체 점검

① 표지만 보고 "무엇을 파는지" 이해되나? ② 문제→해결 흐름이 끊기지 않나? ③ 근거 없는
숫자가 섞이지 않았나(모든 수치가 FINAL-REPORT에 있나)? ④ 한 슬라이드에 메시지가 하나인가?
⑤ 아웃라인 계약의 장 구성·채록 메시지를 그대로 지켰나? 하나라도 아니오면 고친다.
