# 4. DECK-COMPOSE — ② deck-spec 구성 (선별·배치·레이아웃 매칭)

> 파이프라인 두 번째 단계. FINAL-REPORT 재료를 아웃라인 계약대로 B2B 셀링 골격에
> **선별·배치**해 `02_deck-spec.json`(빌더 입력)을 만든다. 창작이 아니라 배치가
> 정체성이다 — "콘텐츠는 사용자가 FINAL-REPORT로 이미 확정했다"
> (`.claude/agents/deck-composer.md`). 규칙 원문: `.claude/skills/deck-compose/SKILL.md`

## 4.1 파이프라인 내 위치

```
 FINAL-REPORT ─▶ ① GRILL(계약) ─▶ [② deck-compose] ─▶ ②.5 승인 ─▶ ③ build ─▶ ④ QA
                                   ████████████████
                                   이 장의 범위 — deck-composer 에이전트(opus) 수행
```

입력은 두 가지 — FINAL-REPORT(재료)와 `01.5_outline.md`(계약, `상세: 3_OUTLINE-GRILL.md`).
출력은 deck-spec.json이며, 그 전에 익스히빗 후보 사양서(`03_exhibit-candidates.json`)를
내고 ②.5 사용자 승인 게이트를 거친다. 스키마 계약은
`.claude/skills/pptx-build/references/deck-spec-schema.md`가 단일 출처다.

## 4.2 내부 흐름 — 재료가 spec이 되기까지

```
 01.5_outline.md(계약) + FINAL-REPORT(재료)
     │
     ▼
 [1] 장별 물성 한 줄 선언 ── "이 내용 = 비교 / 흐름 / 분해 / 경계 / 실물 화면 / …"
     │
     ▼
 [2] layout-matching 결정표 매칭(15타입 전수)
     │
     ├─ 구조 그대로 일치 ──▶ golden.<layout> + content(DEFAULT 키 전량 교체)
     ├─ 항목 수·구획 다름 ──▶ adapted.<layout> + 장 스크립트(goldenfab 부품 재사용)
     └─ 어느 행에도 안 맞음 ─▶ novel + 장 스크립트(골든 렌더 PNG 밀도 앵커)
     │
     ▼
 [3] 익스히빗 후보 사양서 03_exhibit-candidates.json(장당 3안)
     │   check_candidates.py 자기 심사 RESULT: PASS까지
     ▼
 [4] ②.5 승인 게이트 ── render_real_mockups.py 갤러리 → 사용자 "sNN=X" 회신
     │   (회신 전 deck-spec 확정 금지)
     ▼
 02_deck-spec.json ──▶ ③ pptx-build
```

[2]의 3갈래가 2026-07-16 개정의 핵심이다. 사용자가 "텍스트만 갈아 끼운 것"을 기각한
뒤, 골든은 복제 템플릿이 아니라 **변형 가능한 출발점**으로 재정의됐다
(`layout-matching.md:39-56`). 안 맞는 내용을 content에 우겨넣는 것이 기각된
"텍스트 스왑"이고, 그 대신 adapted(장 스크립트 변형)·novel(신규 설계) 출구가 열렸다.
adapted/novel 장은 build_pptx가 `build(prs, content)` 서명의 장 스크립트를 importlib로
로드해 실행하며(`build_pptx.py:1367,1439`), 검증은 스냅샷 회귀가 아니라
audit_deck(전역 오딧+밀도 밴드)+전 장 병렬 채점이 진다(`상세: 8_QA-GATES.md`).
novel 장의 골든 창고 편입은 덱 완성 후 **사용자가 고른 장만** 한다(골든 1장 =
유지보수 계약 1개).

## 4.3 구조 표 — 이 단계를 구성하는 것들

| 구성요소           | 개수                                         | 출처 파일                                                                        |
| ------------------ | -------------------------------------------- | -------------------------------------------------------------------------------- |
| 매칭 결정표        | 15타입 전수 + 3갈래 판정                     | `.claude/skills/deck-compose/references/layout-matching.md:8-24`                 |
| 골든 콘텐츠 계약표 | 10타입 150키(5종은 표 없음)                  | `.claude/skills/deck-compose/references/golden-content-contract.md`              |
| 구성 원칙          | 6칙                                          | `.claude/skills/deck-compose/SKILL.md`                                           |
| 익스히빗 후보      | 장당 3안(A/B/C)                              | 산출물 `03_exhibit-candidates.json`                                              |
| 후보 자기 심사     | 7종 검사(⑦에 조합 시뮬레이션 포함)                   | `.claude/skills/pptx-visuals/scripts/check_candidates.py`                        |
| 다양성 게이트      | 4종(쿨다운·최소 5종·박스 30%·동일 유형 ≤2회) | `audit_pptx.py:44-65` 구현, ②.5 갤러리 JS가 실시간 채점                          |
| 밀도 목표          | 슬라이드당 120~200단어                       | `.claude/skills/deck-compose/SKILL.md`; QA 하한 60단어                           |
| 레거시 견본        | 8장 표준 골격 1식                            | `.claude/skills/deck-compose/assets/standard-deck-spec.json` — 코드 소비 0(§4.6) |

매칭 결정표 15행의 물성→타입 요지(전문은 `layout-matching.md:8-24`):

| 물성(내용의 실제 형태)                   | 타입                   | 물성(내용의 실제 형태)            | 타입                    |
| ---------------------------------------- | ---------------------- | --------------------------------- | ----------------------- |
| 덱의 첫 인상(제목·가치 제안)             | `golden.cover`         | 메커니즘 + 실물 적용 서사         | `golden.tech_mechanism` |
| 부 목록과 이정표                         | `golden.toc`           | 실물 결과 화면 + 스키마·검증 보조 | `golden.tech_capture`   |
| 부의 시작 선언                           | `golden.part`          | 구 vs 신 — 같은 입력의 동작 대비  | `golden.ab_simulation`  |
| 결격·문제의 다구획 해부(원인 4~5개 병렬) | `golden.problem_grid`  | 시험 구성 + 최종 기록 + 실패 해부 | `golden.validation`     |
| 입력→출력의 분기 있는 실행 흐름          | `golden.exec_graph`    | 축별 좌우 대결(자사 vs 경쟁)      | `golden.mirror_matrix`  |
| 주장 + 실물 데이터 증거(표·JSON 카드)    | `golden.tech_evidence` | 경계 — 안(보장)과 밖(거절·유보)   | `golden.boundary`       |
| 위계 구조(트리·fan-out) + 구축 방법      | `golden.tech_tree`     | 덱을 닫는 마침 문장(수미상관)     | `golden.closing`        |
| 실물 화면(캡처)이 주인공                 | `golden.screenshot`    |                                   |                         |

판정 기준은 **주제가 아니라 물성**이다. 두 행에 걸치면 주 물성을 따르고 장을 쪼갠다.
mirror_matrix는 2026-07-16 개편으로 골든 덱(SLIDE_ORDER)에서는 제외됐지만 registry
타입으로는 유지된다(실전 덱 창고, `goldenfab/registry.py:20`).

## 4.4 골든 콘텐츠 계약 — content는 선택이 아니라 필수

`golden.<layout>` 배정 시 content dict가 골든 DEFAULT 키를 **전부** 덮지 않으면
빌더가 중단한다(`goldenfab/content_contract.py`, 2026-07-15 신설; `build_pptx.py:1344`가
assert_content 강제). 부분 주입도 FAIL이다 — 누락 키가 골든 기본값(다른 프로젝트 글)으로
조용히 메워지는 것이 "계약 0% 이행 덱이 전 게이트 통과" 사고의 본질이었기 때문이다
(`layout-matching.md:28-31`; 사고 전말은 `상세: 13_TROUBLESHOOTING.md`).

컴포저가 채울 키의 사전이 `golden-content-contract.md`(코드 DEFAULT dict 덤프)다.
현황 그대로 적으면 — **registry 15타입 중 10타입만 덤프돼 있다**(총 150키):

| 타입           | 키 수 | 타입          | 키 수 |
| -------------- | ----- | ------------- | ----- |
| problem_grid   | 18    | tech_capture  | 14    |
| exec_graph     | 14    | ab_simulation | 15    |
| tech_evidence  | 12    | validation    | 20    |
| tech_tree      | 17    | mirror_matrix | 8     |
| tech_mechanism | 19    | boundary      | 13    |

cover·toc·part·screenshot·closing 5종은 계약표에 표가 없다(DEFAULT dict 보유
10종만 덤프하는 생성기 구조 — `_workspace_kifrs/dump_contract.py:21-32`). 다만 필수
키의 **단일 출처는 계약표가 아니라** `goldenfab/content_contract.required_keys()`로,
이쪽은 15/15 전 타입을 해석한다(`layout-matching.md:34`). 계약표는 사람용 참조
전용이다. 낡은 표기 1건: 계약표가 안내하는 재생성 명령 경로 `_workspace/dump_contract.py`는
스테일이고 실물은 `_workspace_kifrs/dump_contract.py`다. cover·toc·part 3종의 명시
필수 키는 content_contract.py의 EXPLICIT_KEYS가 별도로 든다(cover 5키·toc 4키·part
4키 — `goldenfab/content_contract.py:22-30`).

## 4.5 구성 원칙·다양성 (세부)

deck-compose SKILL의 구성 원칙 6칙: ① 계약 준수(장 추가·삭제·강조 변경 금지)
② 창작 금지(재료 얇으면 부풀리지 않고 "재료 부족" 보고) ③ 숫자는 chart/metrics로
④ 한 메시지 + 120~200단어 ⑤ 개체만 덜렁 금지(commentary 필수) ⑥ 서술 문단 intro.
시각 타입 선택은 "주장 한 문장 먼저 → visual-selection.md 결정표" 순서이며 기본값
bar 금지·bullets는 최후 수단이다. 다양성은 게이트 4종(쿨다운 간격≥3장·본문 5장
이상 시 최소 5종·박스 30% 이하·동일 유형 ≤2회)이 기계 검증한다 —
"다양성 판단을 모델 디폴트에 맡기지 않는 장치다"(`deck-compose/SKILL.md`).
게이트 코드는 QA 층 소관이므로 `상세: 8_QA-GATES.md`, 시각 어휘 자체는
`상세: 6_VISUALS.md`.

②.5 승인 게이트: 후보 사양서는 check_candidates.py의 7종 검사(⑦이 조합 게이트 시뮬레이션까지 겸함)와 게이트 통과 조합
시뮬레이션을 RESULT: PASS까지 자기 심사한 뒤, render_real_mockups.py(실물 COM 렌더,
Windows+PowerPoint 전제)가 gallery.html을 만든다 — 폴백은 make_mockups.py(mpl 스케치).
사용자 회신("s04=A, s06=A, …") 전에는 deck-spec 확정 금지다. K-IFRS 파일럿의 실물
갤러리(`_workspace_kifrs/mockups/gallery.html`)는 본문 8장 × 3안 = 24후보였다.

## 4.6 standard-deck-spec.json — 코드 소비 0곳의 구세대 견본

`.claude/skills/deck-compose/assets/standard-deck-spec.json`은 "문제→솔루션→증거
8슬라이드" 표준 골격의 예시 spec이다(cover/toc/bullets/two_column/bullets/table/
metrics/cta). 실측 현황: **코드 소비자 0곳** — 참조는 deck-compose SKILL.md의 문서
언급 1건뿐이다(grep 1파일). 즉 어떤 스크립트도 이 파일을 읽지 않으며, "빌더가 이
견본을 사용한다"는 식의 현재형 서술은 성립하지 않는다. 내용도 골든 계열(golden.*)
도입 이전의 레거시 타입만 쓰는 구세대 견본이라, 골든 파이프라인(5부 17장)과 병존하는
채로 갱신 없이 남아 어느 쪽이 "표준"인지 문서 간 위상이 갈린 상태다. 이 보고서는
현행 표준을 골든 매칭 경로(§4.2)로 서술한다.

## 4.7 실증 예시 (walked example) — K-IFRS slides[12]가 validation에 앉기까지

실물 spec `_workspace_kifrs/02_deck-spec.json`(16장, meta.frame_style "v3")의
슬라이드 한 건을 추적한다. 이 워크스페이스는 2026-07-15 QA FAIL 후 동결된 스냅샷이며,
파일 내 경로 표기는 구 `_workspace/` 시절 것이다.

**[1] 계약 행** — 아웃라인 계약 장 목록 #13(부 Ⅳ 검증): "QNA 홀드아웃 + 미재현 분해.
leave-all-QNA-out 누수 차단(격리 0/92). 결론 재현 78/92, 하드 59.1%(84건), 에러 0.
미재현 14건 정직 분해"(`_workspace_kifrs/01.5_outline.md:25`). 사용자 확정 조정 ②로
"검증은 1장 유지"가 계약에 박혔다.

**[2] 물성 선언 → 결정표 매칭** — 이 내용의 물성은 "시험 구성(재료 분해) + 최종
기록(78/92) + 실패 해부(미재현 14건)"다. 결정표에서 판정 질문 "검증·테스트 결과를
보고하는 장인가"에 "예"인 행 — `golden.validation`(`layout-matching.md:21`)이
출발점이 된다. 구획 구성(시험 제작·최종 기록·실패 전수 귀속 3구획)이 골든 구조와
그대로 맞으므로 3갈래 중 **golden + content**(텍스트 전량 교체) 경로다. adapted나
novel로 갈 이유가 없는, 결정표의 가장 곧은 경로.

**[3] spec에 앉은 실물** — `02_deck-spec.json` slides[12]:

```json
{ "type": "golden.validation", "content": { "kicker": "4. 검증" } }
```

kicker override 1건만 주입돼 있다(덱 전체에서 kicker override는 이 장과 slides[14]
`golden.ab_simulation`의 `"5. 핵심 의사결정"` 2건뿐). ab_simulation 쪽 override는
골든 기본 kicker "4. 문제와 해결"(구 6부 구성 흔적)을 이 덱의 5부 번호로 덮은 것 —
계약의 부 번호가 골든과 다를 때 content 주입이 어떻게 쓰이는지 보여주는 실물이다.

**[4] 현황 그대로 — content 미주입 동결** — 그러나 validation의 계약 키는 20키다
(§4.4 표). kicker 1키 외 19키가 미주입이고, 본문 8타입 중 6장(slides[3]·[5]·[7]~[10])은
content 키 자체가 없으며 closing도 없다 — **계약 이행 0/8의 물증이 그대로 동결된
상태**다. 이 spec으로 빌드된 덱은 2026-07-15 QA에서 골든오염 9장(검증 장 S13=slides[12]에서만 17건,
ab_simulation S15에서도 17건) 등 계약 위반 17건 FAIL을 받았고, 이후 재작성·재빌드 기록이
없다. meta.note의 "기본값=K-IFRS 내용"이라는 자기변명은 QA 정정에서 기각됐다("골든
기본값은 근거가 아니다"). 같은 날 신설된 content_contract 때문에 **지금 이 spec을
그대로 빌드하면 첫 골든 본문 장인 S4(problem_grid)에서 즉시 중단**된다(gate-repair
스펙: "현 02_deck-spec.json으로 빌드하면 S4에서 즉시 실패"). 즉 이 파일은 "사고
이전의 spec이 사고 이후의 게이트에 막히는" 전환점을 박제한 사료다. 전말은
`상세: 9_KIFRS-PILOT.md`.

## 4.8 경계

이 장은 ② 배치와 ②.5 승인까지다. spec을 pptx로 굽는 빌더는 `상세: 5_PPTX-BUILD.md`,
골든 15타입의 실체와 adapted/novel 렌더 경로는 `상세: 7_GOLDEN-DECK.md`,
content_contract·audit_deck·다양성 게이트의 코드는 `상세: 8_QA-GATES.md`.
