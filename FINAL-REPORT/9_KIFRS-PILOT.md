# 9. K-IFRS 1115 파일럿 — 현재 상태

> **이 장의 결론을 먼저 쓴다: 게이트를 통과한 최종 산출물은 없다.** 첫 실전 파일럿인 K-IFRS 1115 회계기준 덱(16장)은 **2026-07-15 QA에서 계약 위반 17건으로 FAIL** 판정을 받았고, 이후 재빌드·후속 QA 리포트 없이 **동결** 상태다. 이 장은 그 동결 상태를 있는 그대로 기록한다.
>
> 다만 "산출물이 아예 없다"는 부정확하다 — `results/`에는 빌드 실물 4건이 남아 있다(실측 2026-07-19).
>
> | 파일 | 장수 | 크기 | 시점 | 성격 |
> | --- | --- | --- | --- | --- |
> | `results/k-ifrs-1115-소개.pptx` | 20 | 149KB | 2026-07-08 | v3~v4 시기 빌드(골든 체계 이전) |
> | `results/k-ifrs-1115-소개-fable.pptx` | 25 | 173KB | 2026-07-08 | 동상, 변형본 |
> | `results/k-ifrs1115/K-IFRS-1115-온톨로지-RAG-소개.pptx` | 16 | 200KB | 2026-07-14 | 파일럿 빌드 |
> | `results/K-IFRS-1115-온톨로지-RAG-소개.pptx` | 16 | 199KB | 2026-07-15 | **QA FAIL 17건 판정을 받은 실물** |
>
> 즉 파이프라인은 .pptx를 실제로 뽑아내는 데까지는 도달했고, 무너진 곳은 그 다음의 계약·품질 게이트다. 네 건 모두 `*.pptx`가 `.gitignore` L11-12로 비추적이라 클론에는 존재하지 않는다. **현행 골든 17장 체계·계약 게이트를 통과한 완주 PASS는 0건이다.**

## 9.1 파이프라인 내 위치

파일럿은 특정 단계가 아니라 파이프라인 ①~④ 전 구간을 실전 재료로 처음 통과시킨 시도다. 도달 지점은 ④ QA의 FAIL이다.

```
FINAL-REPORT ─▶ ① GRILL ─▶ ② compose ─▶ ②.5 승인 ─▶ ③ build ─▶ ④ QA ─▶ results/
                  │            │            │           │          │          │
                통과         통과(부분)    통과        통과       ▣ FAIL 17건  도달 못 함
                (07-14 계약) (content     (24후보)   (침묵 폴백  (2026-07-15) (완성본 0)
                             미주입)                  으로 빌드됨)    │
                                                                 ▼
                                                          ██ 동 결 ██
```

②의 "통과(부분)"는 deck-spec이 만들어지긴 했으나 본문 content가 미주입된 채였다는 뜻이고, ③의 빌드 성공은 당시 게이트의 침묵 폴백 덕에 가능했던 것이다(§9.4). 각 단계의 내부는 `상세: 3_OUTLINE-GRILL.md`~`8_QA-GATES.md`에 위임한다.

## 9.2 입력 재료 구조

### 9.2.1 `input/k-ifrs-1115/` 3파일 (v2 시절 A/B 보존본 — 전부 배선 0)

| 파일                                        | 규모          | 성격                                     | 배선                     | 비고                              |
| ------------------------------------------- | ------------- | ---------------------------------------- | ------------------------ | --------------------------------- |
| `input/k-ifrs-1115/claude-report-v2.md`     | 61KB · 332줄  | Claude 직접 추출 보고서(8요소 고정 골격) | grep 0건 — 미배선 보존본 | 전 수치에 `[파일명.md]` 출처 태그 |
| `input/k-ifrs-1115/notebooklm-report-1.md`  | 12.8KB · 90줄 | NotebookLM 슬라이드형 초안               | grep 0건 — 미배선 보존본 | 환각 후보 수치 다수(아래)         |
| `input/k-ifrs-1115/notebooklm-report-v2.md` | 43KB · 300줄  | NotebookLM 보고서형 재생성               | grep 0건 — 미배선 보존본 | 세일즈 과장 화법, 인용 각주 소실  |

3파일 합계 116,952바이트. 세 파일 모두 코드·문서 어디서도 소비되지 않는(grep 0건) v2 시절 A/B 실측용 보존본이며, 파일럿의 실제 재료 원천은 v5 계약대로 K-IFRS 프로젝트의 FINAL-REPORT다(재료 동급 실측 6,246단어 ≈ 6,098단어 — `docs/PIPELINE.md` L6-7).

NotebookLM 보조본의 다음 수치는 원 보고서(claude-report-v2.md)에 부재하거나 상충하는 **환각 후보로, 전부 `미검증`**이다 — "오매칭 0/31" · "18/18 활성화"(원 보고서엔 개선 전 수치만 존재) · "누락 문단 163개"(분모 부재) · "1,649 문서 풀"(원천 미상) · 재현율 84.7% vs 84.8%(두 보조본 간 반올림 불일치) · 모델명 3문서 3표기(Gemini Flash(thinking=high) / 1.5 Flash / 3 Flash). 이 중 어느 것이 CLAUDE.md v3 이력의 "환각 2건"인지도 `미검증`이다.

### 9.2.2 파일럿 워크스페이스 `_workspace_kifrs/` 주요 산출물

| 파일                                          | 역할                                           | 상태(2026-07-19 실측)                                        |
| --------------------------------------------- | ---------------------------------------------- | ------------------------------------------------------------ |
| `_workspace_kifrs/01.5_outline.md`            | 아웃라인 계약(주문서) — 사용자 확정 2026-07-14 | 유효 — check_contract가 인자로 소비                          |
| `_workspace_kifrs/02_deck-spec.json`          | 16장 deck-spec(빌더 입력)                      | FAIL 판정 당시 그대로 동결 — 본문 content 미주입             |
| `_workspace_kifrs/03_exhibit-candidates.json` | 본문 8장 × 3안 = 24후보 사양서                 | 산출 완료(②.5 재료)                                          |
| `_workspace_kifrs/mockups/gallery.html`       | ②.5 승인 갤러리(24카드 + 게이트 4종 채점)      | 산출 완료 — s09 기본 체크 B vs 계약 실이미지 불일치 `미검증` |
| `_workspace_kifrs/03_qa-report.md`            | QA 리포트(2026-07-15 정정판)                   | 최종 기록 = FAIL 17건, 이후 후속 리포트 없음                 |
| `_workspace_kifrs/04_contract-check.md`       | check_contract 기계 출력(행별 대조표)          | FAIL 17건의 물증                                             |
| `_workspace_kifrs/HANDOFF.md`                 | 핸드오프 문서(2026-07-15)                      | "다음 작업 5" 미완인 채 마지막 기록                          |
| `_workspace_kifrs/verify_s8s9_data.py`        | S8·S9 수치 기계 대조기(+selftest)              | 미배선 일회성 CLI — 출력 경로 `_workspace/` 스테일           |
| `_workspace_kifrs/golden_defaults.json`       | 골든 8타입 DEFAULT 스냅샷                      | 미배선 유물 — check_contract.py:100이 "스테일" 선언          |

디렉터리 자체의 git 지위도 미결이다 — 미추적(`??`)인데 `.gitignore`에는 `_workspace/`만 있고 `_workspace_kifrs/`는 없어(`.gitignore` L18-19), 커밋/방치 결정이 미뤄진 중간 상태다.

## 9.3 파일럿 진행 흐름 — 종착은 실패·동결

```
[2026-07-14] 아웃라인 계약 확정
  _workspace_kifrs/01.5_outline.md — 16장·5부, 병합 2건(S4+S5, S16+S17),
  확정 이미지 3종(screenshot_answer.png→S6, knowledge_graph_3d.png
  + screenshot_graph_node.png→S9), 앵커수치(등재 423·간선 2694·78/92·59.1% 등)
        │
        ▼
  ② deck-spec 구성 → _workspace_kifrs/02_deck-spec.json
  16장, meta.frame_style="v3" — 본문 8타입 중 6장 content 키 부재
        │
        ▼
  ②.5 승인 갤러리 — 본문 8장 × 3안 = 24후보 (mockups/gallery.html)
        │
        ▼
  ③ 빌드 → deck.pptx — 당시 침묵 폴백(_variant_k.py:99 패턴)으로
    골든 기본 텍스트가 그대로 유출된 채 빌드 성공
        │
        ▼
[2026-07-15] ④ QA(check_contract 계약 대조) ──▶ FAIL — 계약 위반 17건
        │
        ▼
  ██ 동 결 ██  재빌드 없음 · 후속 QA 리포트 없음 · 02_deck-spec 미재작성
              · .pptx 완성본 없음 (2026-07-19 현재)
```

주의: 이 흐름의 판정은 전부 **2026-07-15 당시 기록**이다. 이후 재실행된 QA는 없으므로, "FAIL 17건"은 마지막 기록이지 오늘 다시 돌린 결과가 아니다.

## 9.4 QA FAIL 17건 내역 (2026-07-15 판정, `_workspace_kifrs/04_contract-check.md`)

| 위반 유형          | 건수          | 상세                                                                                       |
| ------------------ | ------------- | ------------------------------------------------------------------------------------------ |
| 골든오염           | 9장           | 장별: S4 10 · S6 19 · S8 24 · S9 17 · S10 17 · S11 20 · S13 17 · S15 17 · S16 2            |
| 계약 이미지 미배치 | 2건(파일 3종) | S6 1종(screenshot_answer.png) · S9 2종(knowledge_graph_3d.png + screenshot_graph_node.png) |
| 플레이스홀더       | 4건           | "My Company" S1·S16, "회사 로고" S2·S16                                                    |
| 앵커수치 누락      | 2건           | S9 '2694' · S15 '1115'                                                                     |

같은 대조에서 장 수 16/16·제목 대조 **13/13**은 PASS였다 — 즉 겉 구조는 맞았고, **내용이 계약을 이행하지 않았다**. (제목 대조의 분모는 16장 중 역할 라벨 행 3개(S1 표지·S2 목차·S16 마무리, 표에 `—(역할)`)를 뺀 13행이다. 원천 리포트가 "14/14"로 적은 것은 역할 라벨을 2행으로 오산한 결과다 — `_workspace_kifrs/03_qa-report.md` L24, gate-repair 스펙 L122. 본 보고서는 `04_contract-check.md` L11~26 표의 재계수를 정본으로 쓴다.) 판정 원칙은 리포트에 명문화돼 있다: "정답은 계약이다. 골든 기본값·팩트시트 대응은 통과 근거가 아니다"(`.claude/skills/consistency-qa/scripts/check_contract.py:168`; `_workspace_kifrs/04_contract-check.md` L5).

두 가지를 함께 기록한다.

1. **이 FAIL 이전의 "PASS" 판정과 "501/501" 수치는 거짓으로 자백·정정됐다**(`_workspace_kifrs/03_qa-report.md` — 정정 경위는 `상세: 13_TROUBLESHOOTING.md`). 즉 파일럿의 신뢰 가능한 마지막 판정은 FAIL 하나뿐이다.
2. **이후 재빌드 리포트가 없다.** 되돌릴 대상 분배(deck-composer: 골든오염·앵커수치 / pptx-builder: 이미지·플레이스홀더)까지 리포트에 적혔으나, 그 되돌림이 실행된 기록은 워크스페이스에 없다.

"회사 로고" 플레이스홀더는 content·brand-kit으로 못 고치는 항목이다 — `goldenfab/layouts.py`(toc)·`goldenfab/s21_closing.py`의 하드코딩이라 코드 수정이 필요하며, check_contract FORBIDDEN 목록에 등록된 알려진 빚으로 남아 있다(`check_contract.py` L33~37).

## 9.5 동결 이후 골든 측 작업과 파일럿의 관계

동결 이후 진행된 작업은 전부 **골든 기준선 쪽**이며, 파일럿 덱 자체를 재빌드한 것은 하나도 없다.

| 시점                            | 골든 측 작업                                                                                                                                                                     | 파일럿에 미친 영향                                             |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------- |
| 2026-07-16 (커밋 56e1e2f)       | 골든 배선 재설계 — adapted/novel 2경로 + 밀도 밴드, 실전 검증은 audit_deck                                                                                                       | 파일럿 spec은 미갱신 — 신규 경로를 쓴 적 없음                  |
| 2026-07-16 (커밋 4834953, HEAD) | 5부 17장 개편 — S17 골든 제외·S18 Part Ⅳ 이동, 스냅샷 519도형                                                                                                                    | 파일럿 계약의 "골든 s17" 참조(구 19장 시절 클로징 지칭)가 낡음 |
| 2026-07-18 (미커밋)             | 골든 고밀도 개편 착수 — 레퍼런스 해부 R1~R12(`golden/_pilot_s06/ref_anatomy.md`), 명사구 재고 units.md v4 ~101단위, s06 파일럿 v1~v6, proto_f 실패 표본, brand-kit `compact: 19` | 파일럿과 무관한 골든 자체 실험 — 전부 골든 미편입              |

s06 파일럿 v6(`.claude/skills/pptx-build/scripts/goldenfab/s06_pilot.py`)은 registry 미등록·코드 import 0곳·git 미추적으로 **미배선 진행 중 실험**이고, `s06_proto_f.py`는 핸드오프가 "실패 표본"으로 명시해 보존한 파일이다(`docs/handoff-2026-07-18-golden-densify.md` L63). `compact: 19`도 미커밋 실험값이다(brand-kit.yaml git status M; 백업 `golden/backup/v5-monochrome-2026-07-18/brand-kit.yaml`과의 diff 실측 유일 차이).

관계를 한 줄로 요약하면: **파일럿이 멈춘 사이 골든이 두 번 개편돼, 동결된 spec은 이제 당시와 다른 공장을 마주한다.** 2026-07-15 게이트 수리(커밋 f6c1458) 이후의 content_contract는 현 `02_deck-spec.json`을 S4에서 즉시 빌드 중단시키며(gate-repair 스펙: "현 02_deck-spec.json으로 빌드하면 S4에서 즉시 실패"), 골든 기준선도 19장·490도형에서 17장·519도형으로 바뀌었다. 동결된 spec은 현행 공장에서 빌드조차 되지 않는다.

## 9.6 남은 일 목록 (전부 근거 경로 병기)

| #   | 남은 일                                              | 실측 상태                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            | 근거                                                      |
| --- | ---------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------- |
| 1   | 파일럿 재빌드 → .pptx 완성본 산출                    | `results/` 최종 산출물 0, 재빌드 기록 없음                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           | `.gitignore` L11-12; `_workspace_kifrs/03_qa-report.md`   |
| 2   | `02_deck-spec.json` content 주입                     | 본문 8타입 중 6장(slides[3]·[5]·[7]~[10])·closing에 content 키 자체가 없음 — 계약 이행 0/8의 물증 그대로                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             | `_workspace_kifrs/02_deck-spec.json`                      |
| 3   | HANDOFF 미완 항목                                    | "다음 작업 5: 나머지 7장 README 충실도"가 미완인 마지막 항목                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         | `_workspace_kifrs/HANDOFF.md`                             |
| 4   | ref_anatomy R1~R12 러너 등록                         | 고밀도 규칙 12종이 design-rules·오딧 러너에 미승격(s06 파일럿 통과 대기) — "러너에 안 물리면 없는 규칙" 상태                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         | `golden/_pilot_s06/ref_anatomy.md`                        |
| 5   | 경로 스테일 `_workspace/` → `_workspace_kifrs/` 해소 | 워크스페이스 개명/이동 후 하드코딩·문서 참조가 구 경로 그대로(2026-07-19 `grep -rn "_workspace/"` 재실측): `check_contract.py` L13·L14(사용법 도크스트링)·L160(이미지 디렉터리)·L272(출력 경로), `build_pptx.py` L84(mpl 익스히빗 PNG 출력처), `goldenfab/content_contract.py` L105(에러 안내), `_workspace_kifrs/dump_contract.py` L8·46, `_workspace_kifrs/verify_s8s9_data.py` L13~15·L305·L338(실행 시 산출물이 다른 폴더 `_workspace/`로 감), `.claude/skills/deck-compose/references/golden-content-contract.md` L5, `_workspace_kifrs/01.5_outline.md` 이미지 경로, `_workspace_kifrs/HANDOFF.md` 실행 커맨드 | 팩트시트 충돌 판정 `_workspace/` 행                       |
| 6   | `_workspace_kifrs/` git 지위 결정                    | 미추적인데 ignore 목록에도 없음 — 커밋/방치 미결                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | `.gitignore` L18-19 vs git status                         |
| 7   | 4단계(덱 재구성)                                     | gate-repair 스펙이 범위 외로 남긴 항목 — 이후 진행 여부 `미검증`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | `docs/superpowers/specs/2026-07-15-gate-repair-design.md` |
| 8   | gallery.html 계약 정합 확인                          | s09 기본 체크 B(표)인데 계약은 실이미지 확정 — 계약 확정 이전 산출물 또는 갱신 누락 `미검증`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         | `_workspace_kifrs/mockups/gallery.html`                   |

## 9.7 실증 예시 — slides[12]·[14]로 추적한 "지금 스펙에 무엇이 있고 무엇이 비어 있는가"

동결된 `_workspace_kifrs/02_deck-spec.json`(16장, meta.frame_style="v3", part 5장 전부 total=5)에서 content가 그나마 **부분 주입된** 두 장을 실값으로 추적한다. 이 두 장이 **본문 8장 중** 가장 채워진 장이다 — 본문 기준으로는 kicker override 2건이 content 주입의 전부이기 때문이다. (정형 장인 cover 1·toc 1·part 5는 이와 별개로 완전한 content를 갖고 있다 — 계약 이행 0/8이 본문만의 수치인 이유다.)

**slides[12]** — `type: "golden.validation"`, content는 kicker override `"4. 검증"` 단 하나. validation 타입의 content 계약은 20키다(`.claude/skills/deck-compose/references/golden-content-contract.md` — 10타입 150키 중 최다). 즉 이 장은 20키 중 1키(kicker)만 덮고 나머지 19키를 비워 둔 상태다. 2026-07-15 당시 빌더는 빈 키를 골든 DEFAULT로 침묵 폴백시켰으므로, 골든 검증 장의 기본 텍스트("78 / 92" 홀드아웃 수치 등 골든 콘텐츠 정본)가 K-IFRS 덱에 그대로 실려 나갔다 — 04_contract-check가 잡은 골든오염과 앵커수치 '1115' 누락(검증 장에 기준서 번호가 없음)이 정확히 이 구조의 결과다.

**slides[14]** — `type: "golden.ab_simulation"`, content는 kicker override `"5. 핵심 의사결정"` 하나. 골든 기본 kicker는 `"4. 문제와 해결"`(구 6부 구성 흔적, `_workspace_kifrs/golden_defaults.json` ab_simulation.kicker)인데 이를 파일럿 목차의 부 이름으로 덮었다 — **override 배선 자체는 동작했다**는 증거다. 그러나 ab_simulation 계약 15키 중 나머지 14키(골든 기본 left_note "105건 중 103건 탈락(98%)" 등)는 미주입 그대로였다.

정리하면: 이 스펙에 들어 있는 것은 타입 배정 16장 + 정형 장 7개(cover·toc·part 5)의 content + 본문 kicker 2건이고, 비어 있는 것은 본문 content 계약 전부다(slides[3]·[5]·[7]~[10]·closing은 content 키 자체가 없음). 그리고 현행 공장에서는 이 추적조차 재현할 수 없다 — 2026-07-15 게이트 수리 이후 content_contract가 이런 스펙을 S4에서 ValueError로 즉시 중단시키므로(`build_pptx.py:1344`의 assert_content 강제), slides[12]까지 빌드가 도달하지 않는다. 침묵 폴백으로 "그럴듯하게 채워진" 덱이 나오던 시대의 마지막 물증이 이 스펙이다.

## 9.8 교차 링크

- 게이트가 이 FAIL을 잡게 된 내부 구조(check_contract 6종 검사·content_contract): `상세: 8_QA-GATES.md`
- "계약 이행 0/8이 전 게이트 PASS" 사고의 근본원인과 수리: `상세: 13_TROUBLESHOOTING.md`
- 동결 이후 골든 개편(5부 17장·adapted/novel·고밀도 실험)의 본체: `상세: 7_GOLDEN-DECK.md`
