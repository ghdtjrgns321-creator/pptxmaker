# 8. QA 게이트 — 오딧·검증층 전체

> 이 장은 pptmaker의 검증층 전체를 다룬다: 어느 산출물에 어느 게이트가 물리는지(게이트 지도), 각 게이트의 검사 항목·판정 기준, 커버리지가 닿지 않는 곳(정직 기재), 그리고 "규칙이 있어도 러너에 안 물리면 없는 규칙"을 막는 현재 장치.
> 전제: 이 프로젝트는 미완성이며 실전 최종 산출물(.pptx 완성본)은 없다. 검증층의 유일한 실전 판정 기록은 K-IFRS 파일럿의 **QA FAIL(계약 위반 17건, 2026-07-15)**이고, 그 후 재빌드 기록 없이 동결 상태다(상세: 9_KIFRS-PILOT.md).

## 8.1 위치

```
① GRILL ─▶ ② compose ─▶ ②.5 승인 ─▶ ③ build ─▶ ④ QA
                                      │            │
                                ┌─────┴─────┐ ┏━━━━┷━━━━━━━━━━━━━┓
                                │ 빌드 중단  │ ┃ 이 장: 오딧·게이트 ┃
                                │ 게이트     │ ┃ 전 층             ┃
                                └───────────┘ ┗━━━━━━━━━━━━━━━━━━┛
```

검증층은 ④ 한 단계에 몰려 있지 않다. ③ 빌드 시점에 박힌 중단 게이트(content_contract), goldenfab 소스를 보호하는 회귀·오딧 게이트(compare_golden·audit_golden), 실전 덱 러너(audit_deck), 레거시 QA(audit_pptx), 계약 3자 대조(check_contract), 그리고 이들을 순서대로 구동하는 consistency-qa 에이전트가 층을 이룬다.

## 8.2 게이트 지도 — 어느 산출물에 어느 게이트가 물리는가

```
┌─ 보호 대상 ──────────────────┐      ┌─ 게이트 ──────────────────────────────────┐
│                              │      │                                            │
│ goldenfab 레이아웃 15종      │──┬──▶│ audit_golden (CLI) — 레이아웃별 SPECS 7종  │
│ (registry.LAYOUTS)           │  │   │  + 검사 규칙 11종 + --selftest 11케이스    │
│                              │  └──▶│ compare_golden (CLI) — 스냅샷 회귀          │
│                              │      │  17장·519도형 전수 + param 유효성 12종     │
│                              │      │                                            │
│ content_contract 계약 자체   │─────▶│ verify_content_contract (CLI)              │
│                              │      │  — 15종 × 5케이스 민감도 검증              │
│                              │      │                                            │
│ deck-spec → .pptx (③ 빌드)   │─────▶│ content_contract.assert_content            │
│  golden.* 장                 │      │  — 누락·빈 값이면 ValueError 빌드 중단     │
│                              │      │                                            │
│ 실전 덱 golden.*(비정형)     │─────▶│ audit_deck (CLI) — 전역 오딧 8항목         │
│  ·adapted.*·novel 장         │      │  + 골든 파생 밀도 밴드                     │
│                              │      │                                            │
│ 레거시 타입 spec 덱          │─────▶│ audit_pptx (CLI) — 검사 13항               │
│                              │      │  (다양성 게이트 4종 포함)                  │
│                              │      │                                            │
│ 계약·spec·pptx 3자           │─────▶│ check_contract (CLI) — 검사 6종            │
│ (01.5_outline.md 기준)       │      │  "정답은 계약"                             │
│                              │      │                                            │
│ ④ 전체 판정                  │─────▶│ consistency-qa 에이전트                    │
│                              │      │  — check_contract → audit_pptx 순서 강제   │
└──────────────────────────────┘      └────────────────────────────────────────────┘
```

역할 분담의 핵심: **compare_golden(스냅샷 회귀)은 goldenfab 보호 전용**이고, 실전 장(adapted/novel 포함)은 스냅샷 대상이 아니다. 실전 장은 3겹으로 검증한다 — ① audit_deck(전역 오딧+밀도 밴드), ② P4 전 장 병렬 채점(FAIL 장만 재채점, 골든 앵커 PNG 대비), ③ 창고 편입은 사용자 선택(출처: `docs/superpowers/specs/2026-07-16-golden-adaptive-wiring-design.md`, `design-rules.md` P5). 병렬 채점의 비용 근거: 순차 2시간+(10장 2시간 30분) → 병렬 최장 1장 약 15분(2026-07-16 스펙 L23). 문서에는 `docs/user/08-system-overview.md` 제목이 "게이트 4겹"으로 남아 있으나 같은 문서의 표 행은 6개로 계수가 불일치한다 — 이 장은 위 지도를 정본으로 쓴다.

## 8.3 게이트별 구조 표

| 게이트                        | 대상                                                | 검사 규모                                                                                              | 판정 기준                                                          | 출처(파일:라인)                                                                                             |
| ----------------------------- | --------------------------------------------------- | ------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------- |
| audit_golden                  | goldenfab 레이아웃(골든 기본값 렌더)                | SPECS 7종 등록(registry 15종 중), 검사 규칙 11종, selftest 11케이스                                    | FAIL이어도 known 기준선 이하면 DEBT/GOOD(래칫), 전체 실행 약 2초   | `.claude/skills/pptx-build/scripts/audit_golden.py:122~200`(SPECS)·`:203~221`(audit)                        |
| compare_golden                | 골든 레퍼런스 덱 17장·519도형                       | 비교 축 8종(도형 수·type·text·fill·run·run_colors·cells·좌표 ±0.005") + names 순서 + param 유효성 12종 | 불일치 1건이라도 FAIL(assert)                                      | `.claude/skills/pptx-build/scripts/compare_golden.py:255`(리포트 경로)·스냅샷 `assets/golden-snapshot.json` |
| audit_deck                    | 실전 장: golden.\*(FIXED 4종 제외)·adapted.\*·novel | generic_checks 8항목 + 밀도 밴드(리터럴 0, 스냅샷 파생)                                                | exit 1=FAIL, spec `"audit": {dup_allow, known}`으로 장별 예외 선언 | `.claude/skills/pptx-build/scripts/audit_deck.py:34`(FIXED)·`:57~58`(accent·밴드)                           |
| audit_pptx                    | 레거시 타입 spec 덱                                 | 검사 13항(다양성 게이트 4종 포함), MIN_BODY_WORDS=60                                                   | exit 0=PASS/1=FAIL                                                 | `.claude/skills/consistency-qa/scripts/audit_pptx.py:44~65`(다양성)·`:145`(60단어)                          |
| check_contract                | 계약(01.5_outline.md)·deck-spec·pptx 3자            | 검사 6종(⓪장 수 ①제목 ②골든 오염 ③계약 이미지 md5 ④금지어 5종 ⑤앵커 수치)                              | FAIL 존재 시 exit 1, 리포트 04_contract-check.md                   | `.claude/skills/consistency-qa/scripts/check_contract.py:31~48`(FORBIDDEN)·`:98~135`(오염)                  |
| verify_content_contract       | content_contract 계약 자체                          | 분모 N=15(레지스트리 전수) × 5케이스                                                                   | exit 0=PASS(케이스 ③④는 ValueError 검출이 성공 조건)               | `.claude/skills/pptx-build/scripts/verify_content_contract.py`                                              |
| content_contract(빌드 게이트) | golden.\* 장의 content 주입                         | 레이아웃별 필수 키 전량(누락·빈 값·오타 키 검출)                                                       | ValueError로 빌드 중단("키 존재는 주입이 아니다")                  | `goldenfab/content_contract.py:94~95`·`build_pptx.py:1344`(import)·`:1353`(강제 호출)                       |
| consistency-qa 에이전트       | ④ 전체 판정                                         | 실행 순서 강제 + 정성 기준표 5항목                                                                     | 계약 대조 FAIL이면 나머지 전부 PASS여도 FAIL, 되돌림 1회           | `.claude/agents/consistency-qa.md:12-13`(순서)·`:33`(되돌림)                                                |

세 오딧의 공용 엔진은 `goldenfab/audit.py`다. 공용 임계: AIR_MIN 0.12 · ACCENT_MAX 4 · CONTRAST_MIN 4.5(WCAG AA) · 채움률 min_ratio 0.30(폭 ≥1.5") · 한글 1자 ≈ 0.8×pt/72 (`goldenfab/audit.py:31~35`·`:151~169`).

## 8.4 세부 — 게이트가 실제로 보는 것

### 8.4.1 audit_golden: 레이아웃별 기대값 + 규칙 11종

골든 기본값으로 각 레이아웃을 렌더한 뒤 SPECS의 기대값을 건다. 검사 규칙 11종(`audit_golden.py:203~221`): §5 진행형 도형 / §2 accent 상한 / P4⑤ 판정 대비 / P4④ 채움률 / P4③ 노드 재탕 / P4⑩ 노드 클래스 / P4⑩ 검정 충돌 / P4⑧ 노드 높이 / §F 그림 침범 / P2③④ 경계 / §6 공기.

SPECS 등록 7종의 기대값 실물(`audit_golden.py:122~200`):

| 레이아웃      | 기대값                                          |
| ------------- | ----------------------------------------------- |
| problem_grid  | progress 0 · ink_allow 1 · node_top_max 4.6     |
| screenshot    | known {"§2 accent 상한": 5}                     |
| tech_capture  | known {"§2 accent 상한": 6, "P4⑩ 검정 충돌": 2} |
| tech_evidence | ink_allow 2                                     |
| tech_tree     | ink_allow 2 · dup_allow 8                       |
| mirror_matrix | ink_allow 8                                     |
| boundary      | ink_allow 3                                     |

**known 래칫**: FAIL이어도 위반 수 n이 등록된 기준선 이하면 DEBT/GOOD으로 통과, 기준선 초과 시 FAIL(`goldenfab/audit.py:421~426`). 알려진 빚(예: screenshot accent 5는 상한 4 초과)을 기준선으로 박아 두고, 더 나빠지는 것만 막는 설계다. `--selftest` 11케이스는 규칙 민감도를 뮤테이션으로 자체 검증한다 — 셰브런 검출·흐린 판정 검출·산문 오탐 0·그림 침범(구 s10 방식 재현)·클래스 충돌·의도된 대조 오탐 0(S4 두 ◇)·§6-D 희소 검출(도형 4개)·골든 오탐 0(SPECS 7종 전수)·대비 산수(muted<4.5<primary) 등.

### 8.4.2 compare_golden: 스냅샷 회귀

goldenfab으로 지금 빌드한 골든 레퍼런스 덱(17장)을 `assets/golden-snapshot.json`(519도형: TEXT_BOX 226·AUTO_SHAPE 216·LINE 75·PICTURE 2)과 전수 대조한다. 비교 축: 도형 수 / type / text / fill(솔리드 fore_color, enum 직접 대조 — str 비교 금지) / run(첫 런 pt·bold) / run_colors(전 런) / cells(표 셀) / l·t·w·h ±0.005", 그리고 names 순서(SLIDE_ORDER 변경 감지). 게이트 민감도 실측: MARGIN_L 0.6→0.61 변조 시 197건 검출, brand-kit accent 변조(D66E3A→00FF00) 시 55건 검출(gate-repair 스펙 L189·L211-213). 추가로 param_efficacy 12종 — closing·screenshot 커스텀 content 반영 + 10종 headline override 반영을 확인한다. 의도적 변경은 `--update-snapshot` 후 git diff 리뷰가 절차이고, 스냅샷 손편집은 금지다(스냅샷 `_comment`).

이 게이트가 못 보는 것도 명시돼 있다: **폰트 이름은 비교하지 않는다**(크기·볼드만 — `docs/user/07-factory-port.md` L4-5, gate-repair 스펙의 알려진 공백). "형태가 내용에 맞나"도 못 본다 — 그건 P1(물성 선언)·P4(제3자 채점) 몫이다.

### 8.4.3 audit_deck: 실전 장 러너 + 밀도 밴드

spec에서 golden.\*(정형 FIXED={cover,toc,part,closing} 제외)·adapted.\*·novel 장을 골라 generic_checks 8항목을 적용한다(`audit_deck.py:34`·`:57~58`): §2 accent 상한 / P4⑤ 판정 대비 / P4④ 채움률 / P4③ 노드 재탕(dup_allow) / P4⑩ 노드 클래스 / §F 그림 침범 / P2③④ 경계(skip_tops 6.60·7.15 — 결론 바·출처선 위치는 예외) / §6-D 밀도 밴드.

**밀도 밴드는 리터럴 0**이다 — golden-snapshot.json에서 fixed 제외 본문 장의 도형 수·비어있지 않은 텍스트 프레임 수의 **최솟값을 파생**한다(`goldenfab/audit.py` density_band, L365 부근; 2026-07-16 스펙 L48). 골든이 고밀도로 진화하면 실전 장의 하한도 자동으로 따라 올라간다. 장별 예외는 코드 수정 없이 spec의 `"audit": {"dup_allow": N, "known": {규칙명: 기준선}}` 선언으로 처리한다. 단, audit_deck은 CLI다 — 빌드에 자동 후크로 물려 있지는 않다.

### 8.4.4 audit_pptx: 레거시 QA + 다양성 게이트 4종

레거시 타입 spec 전용 기계 QA. 검사 13항: slide_count · native_tables · native_charts(멀티패널은 패널 수만큼 계수, mpl 이미지 경로는 제외) · image_exhibits · skeleton(cover→toc…cta) · density(본문 장 <60단어 FAIL, MIN_BODY_WORDS=60) · 다양성 4종 · bullets≤2장 · part_navigation(slides≥12 시 part≥3) · font_consistency(--brand 지정 시).

다양성 게이트 4종(`audit_pptx.py:44~65`): 쿨다운(동일 유형 간격 ≥3) · 최소 5종(본문 ≥5장일 때) · 박스 비율 ≤30%(BOX={flow,layers,cards,branch,from_to}) · 동일 유형 ≤2회. 문서 곳곳(pptmaker SKILL ④, visual-selection.md C-2, audit_pptx.py 자체 main 주석 L163)에는 "3종"으로 남아 있으나 구현은 4종이 정본이다(diversity_checks 도크스트링 L40).

### 8.4.5 check_contract: "정답은 계약이다"

아웃라인 계약(01.5_outline.md 표) vs deck-spec vs 빌드 pptx 3자 대조. 검사 6종: ⓪ 장 수 ① 제목 대조 ② 골든 오염 — spec type의 레이아웃 키로 **살아있는** `content_contract.golden_defaults()`에서 len≥12 문자열을 수집해 슬라이드 텍스트 포함 시 FAIL(adapted.\*도 동일 잣대, `check_contract.py:129~135`; golden_defaults.json 스냅샷은 스테일로 선언하고 안 쓴다 — `check_contract.py:100`) ③ 계약 이미지(md5 대조 — "아무 그림"이 아니라 그 그림) ④ 금지어 5종(My Company/회사 로고/Lorem ipsum/TODO/샘플 텍스트) ⑤ 앵커 수치(계약 message의 `\d+\.\d+%|\d+/\d+|\d{3,}` 패턴이 슬라이드 텍스트에 실존). 원칙 문장이 코드에 있다: "정답은 계약이다. 골든 기본값·팩트시트 대응은 통과 근거가 아니다"(`check_contract.py:168`).

### 8.4.6 verify_content_contract: 게이트의 게이트

content_contract 자체의 민감도를 검증하는 CLI(코드 import 0곳 — 수동 실행 전용). 분모 N=15(레지스트리 전수) × 5케이스: ① required_keys 해석 가능 ② content 완비 → 통과(오탐 0) ③ 1키 누락 → ValueError 검출 ④ 전 키 None → ValueError 검출("키 존재≠주입" 회귀 방지) ⑤ 레이아웃 소스의 폴백 패턴(`c["x"] or [...]`) 잔존 스캔. exit 0=PASS.

### 8.4.7 consistency-qa 에이전트: 층의 지휘자

④ 단계의 실행 주체(에이전트 정의 실존, model opus). 규칙: 실행 순서 강제(check_contract 먼저 → audit_pptx 다음, `.claude/agents/consistency-qa.md:12-13`) / 계약 대조 FAIL이면 나머지 전부 PASS여도 판정 FAIL / 게이트 수치는 이번 실행 결과만 인용(기억·이전 리포트 복사 금지) / 집합 주장에 분모 명시 / 되돌림 1회, 재실패면 잔여 결함 명시 후 전달(무한 루프·조용한 통과 금지). 기계로 셀 수 있는 것은 스크립트, 판단이 필요한 것은 스킬 기준표(정성 5항목)로 나눈다.

## 8.5 "규칙이 있어도 러너에 안 물리면 없는 규칙" — 현재 방지 장치

이 원칙은 코드에 문장으로 박혀 있다: `audit_golden.py:184~185`. 배경만 요약하면 — 채움률 규칙(P4④)이 존재했는데도 러너에 등록되지 않아 미러 매트릭스의 채움률 14~29% FAIL 9건이 방치된 사례가 있었다(CLAUDE.md 2026-07-16 이력; 사고 경위는 13_TROUBLESHOOTING.md 소관). 현재 장치는 다음과 같다.

| 장치                         | 내용                                                                                                                   | 근거                                                 |
| ---------------------------- | ---------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------- |
| SPECS 등록으로 러너에 물리기 | mirror_matrix(ink_allow 8)·boundary(ink_allow 3)를 audit_golden SPECS에 등록 — 2026-07-16에 2종 추가 등록으로 누계 7종 | `audit_golden.py:122~200`; CLAUDE.md 2026-07-16 이력 |
| 규칙의 실행 코드 하강        | 채점자 FAIL 9건 중 7건이 순수 산수 → 산수는 audit.py 규칙으로 이식(0.5초), "채점자는 검사기가 아니라 광산"             | `goldenfab/audit.py:9~17`; design-rules.md P1.5·P2   |
| --selftest로 죽은 규칙 검출  | 규칙이 결함을 못 잡으면 selftest FAIL — "안 잡으면 죽은 코드"(11케이스 뮤테이션)                                       | `audit_golden.py` selftest; design-rules.md P2       |
| 전역 규칙의 실전 러너 배선   | generic_checks 8항목이 audit_deck 경유로 실전 장에도 돈다 — 골든 전용 규칙으로 고립되지 않게                           | `audit_deck.py:34`·`:57~58`                          |
| 위반 수의 직접 반환          | report가 위반 수를 규칙 함수에서 직접 받는다(문자열 파싱 금지) — 계수 누락으로 규칙이 무력화되는 경로 차단             | `goldenfab/audit.py:421~426`                         |

같은 원칙의 **미봉합 지점도 있다**: `golden/_pilot_s06/ref_anatomy.md`의 고밀도 규칙 R1~R12는 임계 표까지 있으나 design-rules·오딧 러너에 미등록 상태다(파일럿 s06 통과 대기) — 문서 스스로 "러너에 안 물리면 없는 규칙" 상태라고 적어 두었다.

## 8.6 커버리지 갭 (정직 기재)

| #   | 갭                                 | 내용                                                                                                                                                                                                                         | 출처                                                                     |
| --- | ---------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------ |
| 1   | audit_golden SPECS 부분 커버       | registry 15종 중 **7종만** 등록. exec_graph·tech_mechanism·ab_simulation·validation·closing은 레이아웃별 오딧 미등록(cover·toc·part는 정형이라 의도적 제외로 읽힘) — "전 타입 오딧"이 아니다                                 | `audit_golden.py:122~200`                                                |
| 2   | verify_content_contract 케이스⑤ 갭 | 폴백 스캔 정규식이 `c["x"] or [`/`(`/`"` 형태만 탐지 — `c.get("x") or`·dict 리터럴(`or {`) 폴백은 미커버                                                                                                                     | `verify_content_contract.py:91`                                          |
| 3   | s11·s12 자체 오딧의 accent 리터럴  | 골든 채택 변형의 빌드 시 자체 오딧이 accent를 "D66E3A" 리터럴로 검사 — kit의 C["accent"]와 이중화. 현재 값은 일치하나 brand-kit 변경 시 오딧이 침묵 불일치할 소지                                                            | `goldenfab/s11_variants.py:997`; s12 동일 패턴 `s12_variants.py:444~460` |
| 4   | audit.py 이중 출처 상수            | CONTENT_BOTTOM 6.35·RIGHT_EDGE 12.733을 grid.py에서 import하지 않고 자체 리터럴로 중복 보유 — 값은 일치하나 grid 변경 시 수동 동기화 필요                                                                                    | `goldenfab/audit.py:33~34` vs `goldenfab/grid.py`                        |
| 5   | 폰트 이름 미검사                   | compare_golden이 pt·bold만 보고 폰트 이름은 비교 안 함 — 알려진 공백, 미해결                                                                                                                                                 | gate-repair 스펙; `docs/user/07-factory-port.md` L4-5                    |
| 6   | audit_pptx는 골든 경로에 못 씀     | skeleton 검사가 cta 종결을 요구 — closing으로 끝나는 골든 경로 덱에는 그대로 적용 불가(레거시 spec 전용)                                                                                                                     | `audit_pptx.py` skeleton 검사                                            |
| 7   | audit_deck의 규칙 부분집합         | 레이아웃 특정 규칙(진행형 도형·ink_allow·공기 쌍·노드 높이)은 audit_deck에서 안 돈다 — audit_golden과의 의도된 커버리지 차이지만, 실전 장은 그만큼 얇게 검사된다                                                             | `audit_deck.py` generic_checks 8항목 vs `audit_golden.py` 11종           |
| 8   | golden-content-contract 부분 덤프  | 계약 참조 문서는 15타입 중 10타입만 덤프(cover/toc/part/screenshot/closing 표 없음) — 필수 키의 단일 출처는 `content_contract.required_keys()`(15/15 해석)라 강제에는 구멍 없음, 사람용 문서만 부분 커버                     | `.claude/skills/deck-compose/references/golden-content-contract.md`      |
| 9   | 다양성 게이트 정의 3곳 분산        | audit_pptx.diversity_checks / render_real_mockups의 JS 채점 / 그리디 페널티가 별도 구현 — 드리프트 위험                                                                                                                      | `render_real_mockups.py` 비고(notes_F)                                   |
| 10  | 자동 후크 부재                     | audit_deck·audit_golden·compare_golden·verify_content_contract는 전부 CLI — 빌드가 자동으로 부르지 않으며, 절차(SKILL·design-rules)가 실행을 지시한다. build_pptx의 main()은 반환값이 없어 항상 exit 0(실패는 예외로만 전파) | notes_D 배선 실측; build_pptx.py 비고                                    |
| 11  | shape_kind 죽은 모듈               | 물성 선언 assert 모듈은 존재하나 호출 0곳(미배선) — sequence kind는 검사 로직 자체가 없음. 같은 취지는 audit_golden SPECS의 progress 기대값이 대체 수행                                                                      | `goldenfab/shape_kind.py`(notes_D)                                       |

## 8.7 실증 예시 (walked example) — FAIL 17건은 어느 검사에 걸렸나

2026-07-15, K-IFRS 16장 덱의 check_contract 실행 1건을 실값으로 추적한다(출처: `_workspace_kifrs/04_contract-check.md`, `_workspace_kifrs/03_qa-report.md`).

입력 3개: 계약 `01.5_outline.md`(16장·5부, 사용자 확정 2026-07-14) + `02_deck-spec.json`(16장, 본문 8타입 중 6장 content 키 부재) + 빌드된 deck.pptx.

- **검사 ⓪ 장 수**: 계약 16행 = pptx 16장 → **PASS(16/16)**.
- **검사 ① 제목 대조**: 역할 라벨 행(표지 S1·목차 S2·마무리 S16) 스킵 후 13장 대조 → **PASS(13/13)**. 여기까지만 보면 멀쩡한 덱이다. (원천 리포트의 "14/14"는 역할 라벨을 2행으로 오산한 값 — 분모 재계수 근거는 [9_KIFRS-PILOT.md](9_KIFRS-PILOT.md) §9.4.)
- **검사 ② 골든 오염**: spec type(예: slides[7]의 `golden.tech_evidence`)에서 접두를 뗀 레이아웃 키로 살아있는 `golden_defaults()`의 len≥12 문자열을 수집해 슬라이드 텍스트와 대조 → **9개 장에서 검출**: S4 10건 · S6 19건 · S8 24건(최다) · S9 17건 · S10 17건 · S11 20건 · S13 17건 · S15 17건 · S16 2건. content 미주입 장들이 골든 덱(다른 프로젝트)의 문장을 그대로 싣고 있었다 → 장 수 기준 **FAIL 9건**.
- **검사 ③ 계약 이미지**: 계약이 확정한 3종(screenshot_answer.png→S6, knowledge_graph_3d.png+screenshot_graph_node.png→S9)의 md5가 pptx 안 그림과 불일치(누락) → **FAIL 2건**(S6 1·S9 2, 파일 3종).
- **검사 ④ 금지어**: "My Company"(S1·S16)·"회사 로고"(S2·S16) → **FAIL 4건**. 이 중 "회사 로고"는 content·brand-kit으로 못 고치는 `layouts.py`(toc)·`s21_closing.py` 하드코딩 — FORBIDDEN에 등록된 알려진 빚이다(`check_contract.py:33~37`).
- **검사 ⑤ 앵커 수치**: 계약 message의 수치 패턴 중 S9의 '2694'(간선 총계)·S15의 '1115'(기준서 번호)가 슬라이드 텍스트에 부재 → **FAIL 2건**.

합계 **FAIL 17건(=9+2+4+2), exit 1** — 리포트 `04_contract-check.md` 생성, consistency-qa 규칙("계약 대조 FAIL이면 전체 FAIL")에 따라 덱 판정 FAIL. 이 판정 이후 재빌드 기록이 없어 파일럿은 이 상태로 동결이다. 반면교사도 같은 spec에 있다: slides[3]·[5]·[7]~[10]·closing처럼 content 키가 아예 없는 장은, 2026-07-15 이후로는 check_contract까지 갈 것도 없이 ③ 빌드 시점에 content_contract.assert_content가 S4에서 ValueError로 중단시킨다(`build_pptx.py:1353`; gate-repair 스펙 "현 spec으로 빌드하면 S4에서 즉시 실패").

대조 사례로, goldenfab 보호 게이트의 최근 판정 1건: compare_golden 재실행 결과 **17/17장 · 519/519도형 · 불일치 0 — PASS**, param 유효성 12종 PASS 동봉(`golden/variants/compare_full.md`, 5부 개편 후 재실행분).

---

관련 장: 골든 덱·스냅샷의 생성 측은 7_GOLDEN-DECK.md, 파일럿 동결 상태는 9_KIFRS-PILOT.md, 게이트가 이렇게 설계된 사고 경위는 13_TROUBLESHOOTING.md.
