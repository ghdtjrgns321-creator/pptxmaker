# 2. PIPELINE — 입력(FINAL-REPORT)에서 종착(.pptx)까지

> 이 장은 전체 흐름 1장짜리 지도다. 각 단계의 세부 규칙은 상세 장([3_OUTLINE-GRILL.md](3_OUTLINE-GRILL.md)~[8_QA-GATES.md](8_QA-GATES.md))에 위임하고, 여기서는 단계당 1문단까지만 쓴다. 마지막 절은 실제 입력 1건(K-IFRS 1115)이 각 단계에서 무엇으로 변환됐는지의 실값 추적이다.
>
> **정직 고지**: 아래 배선은 스킬·에이전트·스크립트 실존이 실측으로 확인된 현행 구조다. 다만 이 흐름을 끝(`results/`)까지 PASS로 완주한 실전 사례는 아직 0건이다 — `results/`에 남은 빌드 실물 4건은 전부 골든 체계 이전이거나 QA FAIL 판정분이다([9_KIFRS-PILOT.md](9_KIFRS-PILOT.md) 서두 표).

## 1. 전체 흐름

```
 <프로젝트>/FINAL-REPORT/  (보고서형 md 묶음 — 사용자가 사전 정리. 없으면 진행 금지)
        │
        ▼
┌─ ① deck-outline-grill (목차·콘텐츠 인터뷰 — 메인 대화 직접 수행, 자동 진행 금지) ─┐
│   목차 1안 선제시 → 부·장 단위 육성 채록 → 사용자 "확정" 전 ② 진행 금지          │
└────────────────┬──────────────────────────────────────────────────────────────────┘
                 ▼
          01.5_outline.md  ◀─ 계약: 장 구성·채록 원문·확정 이미지·앵커 수치
                 │            (이후 모든 단계가 이 계약을 소비 — QA의 정답지)
                 ▼
┌─ ② deck-composer (에이전트) ──────────────────────────────────────────────────────┐
│   FINAL-REPORT에서 선별·배치(창작 0) + layout-matching 결정표(15타입)로 3갈래 판정 │
│     구조 일치      ──▶ golden.<layout> + content 전량 주입                        │
│     구획·항목 다름 ──▶ adapted.<layout> + 장 스크립트                             │
│     매칭 없음      ──▶ novel + 장 스크립트(골든 렌더 밀도 앵커)                   │
└────────────────┬──────────────────────────────────────────────────────────────────┘
                 ▼
          03_exhibit-candidates.json (본문 장당 3안 + check_candidates.py 자기 심사)
                 │
                 ▼
┌─ ②.5 익스히빗 승인 게이트 (사용자) ───────────────────────────────────────────────┐
│   render_real_mockups.py(실물 COM 렌더 갤러리; COM 불가 시 폴백 make_mockups.py)  │
│   → gallery.html → 사용자 회신("sNN=X") 전 ③ 진행 금지                            │
└────────────────┬──────────────────────────────────────────────────────────────────┘
                 ▼
          02_deck-spec.json 확정  (구성 SSOT)
                 │
                 ▼
┌─ ③ pptx-builder (에이전트) — build_pptx.py + brand-kit.yaml ──────────────────────┐
│   타입 디스패치 3계열:                                                            │
│     레거시 11종(cover/toc/…/cta)  ──▶ 마스터 틀 자동 스탬프                       │
│     golden.<layout>               ──▶ goldenfab registry 15종                     │
│         └ content_contract: content 전량 주입 강제 — 누락·빈 값 ValueError 중단   │
│     adapted.<layout> / novel      ──▶ 장 스크립트 build(prs, content) 로드        │
│   실전 장 오딧: audit_deck.py (generic 8항목 + 골든 스냅샷 파생 밀도 밴드)        │
└────────────────┬──────────────────────────────────────────────────────────────────┘
                 ▼
             deck.pptx
                 │
                 ▼
┌─ ④ consistency-qa (에이전트) — 실행 순서 강제 ────────────────────────────────────┐
│   1) check_contract.py: 계약·spec·pptx 3자 대조(검사 6종) — FAIL이면 전체 FAIL    │
│   2) audit_pptx.py: 네이티브 표·차트/밀도/다양성 게이트 4종 등 13항               │
│   3) 전 장 COM 렌더 PNG 눈검증 + 정성 기준표 5항목                                │
└──────┬──────────────────────────────┬─────────────────────────────────────────────┘
      PASS                          FAIL
       │                              │
       ▼                              ▼
 results/<프로젝트명>-소개.pptx   원인 단계 되돌림(1회) — 재실패 시 잔여 결함 명시
 (완주 PASS 0건 — 잔존 4건은     (K-IFRS 파일럿: 계약 위반 17건 FAIL 후 동결)
  게이트 미통과분)
```

파이프라인 밖에 회귀 게이트가 하나 더 있다: `compare_golden.py`는 goldenfab(골든 렌더 코드)을 수정했을 때 골든 스냅샷(17장 519도형)과 전수 대조하는 **goldenfab 보호 전용** 게이트로, 실전 덱 빌드 흐름에는 끼지 않는다(최근 판정 17/17장·519/519도형·불일치 0 PASS). 상세: [8_QA-GATES.md](8_QA-GATES.md).

## 2. 단계 표

| 단계     | 실행 주체                     | 입력                       | 처리                                                            | 출력                                                | 상세 장                                  |
| -------- | ----------------------------- | -------------------------- | --------------------------------------------------------------- | --------------------------------------------------- | ---------------------------------------- |
| ① 인터뷰 | deck-outline-grill(메인 대화) | FINAL-REPORT               | 목차 1안 제시 → 부·장 육성 채록 → 사용자 "확정"                 | `01.5_outline.md`(계약)                             | [3_OUTLINE-GRILL.md](3_OUTLINE-GRILL.md) |
| ② 구성   | deck-composer(에이전트)       | 계약 + FINAL-REPORT        | 선별·배치(창작 0), 매칭 3갈래 판정, 본문 장당 후보 3안          | `03_exhibit-candidates.json`                        | [4_DECK-COMPOSE.md](4_DECK-COMPOSE.md)   |
| ②.5 승인 | 사용자                        | 후보 사양서                | 실물 렌더 목업 갤러리 확인 → "sNN=X" 회신                       | `02_deck-spec.json` 확정                            | [4_DECK-COMPOSE.md](4_DECK-COMPOSE.md)   |
| ③ 빌드   | pptx-builder(에이전트)        | deck-spec + brand-kit.yaml | build_pptx 3계열 디스패치 + content_contract + audit_deck       | `deck.pptx`                                         | [5_PPTX-BUILD.md](5_PPTX-BUILD.md)       |
| ④ QA     | consistency-qa(에이전트)      | 계약 + spec + deck.pptx    | check_contract(6종) → audit_pptx(13항) → 렌더 눈검증·정성 5항목 | `04_contract-check.md`·`03_qa-report.md`, PASS/FAIL | [8_QA-GATES.md](8_QA-GATES.md)           |

**①** 은 승인 게이트가 아니라 콘텐츠 인터뷰다 — 장별 "가장 전하고 싶은 것"을 육성으로 채록해 원문 그대로 계약에 싣고, 전 장이 "확정"되기 전에는 ②로 못 넘어간다. 계약에는 장 구성뿐 아니라 확정 이미지(md5 대조 대상)와 앵커 수치까지 박힌다.

**②** 는 창작이 아니라 선별·배치다 — FINAL-REPORT에 없는 수치·성과는 만들지 않고, 각 장의 물성(비교·흐름·분해·경계·실물 화면)을 layout-matching 결정표에 대조해 golden/adapted/novel 3갈래 중 출발점을 정한다. 안 맞는 내용을 골든 틀에 우겨넣는 "텍스트 스왑"은 기각된 방식이다.

**②.5** 는 사용자 승인 게이트다 — 본문 장당 서로 다른 유형 3안을 실물 렌더 갤러리로 제시하고, 회신이 있어야 deck-spec이 확정된다(구성 반려를 빌드 전 텍스트 단계로 앞당기는 장치).

**③** 은 결정론 빌드다 — 같은 spec과 brand-kit이면 같은 .pptx가 나온다. golden.\* 경로는 content 전량 주입을 content_contract가 강제해, 누락 키가 골든 기본값(다른 프로젝트의 글)으로 메워지는 사고를 ValueError로 차단한다. adapted/novel은 장 스크립트를 importlib로 로드해 실행하고, 실전 장 검증은 audit_deck(전역 오딧 + 골든 스냅샷 파생 밀도 밴드)이 맡는다.

**④** 는 "존재 확인이 아니라 교차 검증"이다 — 정답은 골든 기본값도 팩트시트도 아닌 계약(01.5_outline.md)이며, 계약 대조가 FAIL이면 나머지가 전부 PASS여도 판정은 FAIL이다. 되돌림은 1회로 제한되고, 재실패면 잔여 결함을 명시한 채 보고한다.

## 3. 실증 예시 (walked example) — K-IFRS 1115: 입력에서 동결까지

유일한 실전 완주 시도의 실값 추적이다. 산출물 실물은 `_workspace_kifrs/`에 동결돼 있다. 종착은 성공이 아니다 — 그대로 쓴다.

```
 FINAL-REPORT(md 12개) ─▶ ① 01.5_outline.md(16장·5부) ─▶ ② 후보 24안 ─▶ ②.5 gallery.html
                                                                              │
      동결(재빌드 없음) ◀─ ④ check_contract FAIL 17건 ◀─ ③ deck.pptx(16장) ◀─ 02_deck-spec.json
                                                                              (계약 이행 0/8)
```

**입력.** K-IFRS 1115 프로젝트의 FINAL-REPORT — 이 리포 밖(`../k-ifrs-1115/FINAL-REPORT/`)에 있는 md 12개 묶음이다(`_workspace_kifrs/01.5_outline.md:5`가 "FINAL-REPORT/\*.md (12개)"로 지목). 이전 기록의 "6,098단어(A/B 실측에서 직접 추출 재료 6,246단어와 동급)"는 측정 시점·명령이 남아 있지 않아 `미검증`이다 — 2026-07-19에 같은 묶음을 `wc -w`로 재측정하면 35,130(공백 분리 어절)이 나와 당시 값과 자릿수가 다르다. 인터뷰에 실제로 들어간 분량은 재구성할 수 없으므로 이 장은 "md 12개"만 검증값으로 쓴다.

**① 목차 인터뷰 → `01.5_outline.md`.** 2026-07-14 사용자 확정. 16장·5부 구성, 병합 2건(S4+S5, S16+S17) 반영. 계약에 확정 이미지 3종이 배치되고(`screenshot_answer.png`→S6, `knowledge_graph_3d.png`+`screenshot_graph_node.png`→S9), 앵커 수치가 박혔다 — 용어 등재 423, 간선 2,694, 홀드아웃 78/92, 하드 재현율 59.1%, 리랭커 105건 중 103 탈락.

**② 구성 → `03_exhibit-candidates.json`.** 본문 8장(S4·6·8·9·10·11·13·15) × 3안 = **24후보**, 각 장 A안 preferred. 예: S6은 A안 `golden.screenshot`(screenshot_answer.png) 대 B·C안 네이티브 대안 구도.

**②.5 승인 → `02_deck-spec.json` 확정.** 목업 갤러리 실물이 `_workspace_kifrs/mockups/gallery.html`로 남아 있다. 확정된 spec은 16장, `meta.frame_style` "v3", part 5장 전부 total=5. `slides[12]`는 `type: "golden.validation"`에 kicker override `"4. 검증"`, `slides[14]`는 `type: "golden.ab_simulation"`에 kicker override `"5. 핵심 의사결정"`. 그러나 **본문 8타입 중 6장(slides[3]·[5]·[7]~[10])과 closing에는 content 키 자체가 없었다** — 주입은 kicker override 2건이 전부. 계약이 요구한 본문 콘텐츠 이행 0/8.

**③ 빌드 → `deck.pptx` 16장.** 당시 빌더에는 침묵 폴백(`_variant_k.py:208`의 `{**DEFAULT, **(c or {})}`)이 있어, 비어 있는 content가 골든 기본값 — 즉 골든 덱의 글과 숫자 — 로 조용히 메워진 16장 덱이 빌드됐다. 빈 장이 아니라 그럴듯하게 채워진 장이 나왔으므로 아무도 못 봤다.

**④ QA(게이트 수리 후 재검, 2026-07-15) → `04_contract-check.md`.** 계약·spec·pptx 3자 대조 결과: 장 수 16/16 PASS, 제목 **13/13** PASS(S1·S2·S16은 제목대조 열이 "—(역할)"이라 대조 대상에서 빠진다 — `04_contract-check.md:11~26`) — 그러나 **계약 위반 17건, exit 1, 판정 FAIL**:

| 위반 유형        | 건수 | 실값                                                                                             |
| ---------------- | ---- | ------------------------------------------------------------------------------------------------ |
| 골든오염         | 9장  | 장별 오염 문자열 S4 10 · S6 19 · S8 24(최다) · S9 17 · S10 17 · S11 20 · S13 17 · S15 17 · S16 2 |
| 계약 이미지 누락 | 2건  | 파일 3종(S6 1·S9 2)이 덱에 없음                                                                  |
| 플레이스홀더     | 4건  | "My Company"(S1·S16), "회사 로고"(S2·S16)                                                        |
| 앵커 수치 누락   | 2건  | S9 '2694' · S15 '1115'                                                                           |

같은 검증에서 이전 기록의 "PASS" 판정과 "501/501" 수치는 거짓으로 자백·정정됐다(스테일 산출물을 재측정 없이 복사한 수치 — 상세: [13_TROUBLESHOOTING.md](13_TROUBLESHOOTING.md)).

**종착 — 동결.** FAIL 17건 이후 재빌드·후속 QA 리포트가 없다. `02_deck-spec.json`은 미재작성 상태 그대로이고, `results/`의 최신 실물(`K-IFRS-1115-온톨로지-RAG-소개.pptx`, 16장·199KB, 2026-07-15)이 바로 FAIL 17건 판정을 받은 그 덱이다 — 그 뒤로 갱신이 없다. 현행 코드로 이 spec을 다시 빌드하면 content_contract가 **S4에서 즉시 ValueError로 중단**한다 — 이제는 오염된 덱이 나오는 대신 빌드 자체가 실패한다. 파일럿의 현재 상태 전수 목록은 [9_KIFRS-PILOT.md](9_KIFRS-PILOT.md).
