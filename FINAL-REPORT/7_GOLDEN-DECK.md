# 7. 골든 덱 — 5부 17장 레퍼런스와 goldenfab

> 골든 덱은 손으로 깎아 확정한 **5부 17장짜리 레퍼런스 프레젠테이션**이다. 실전 덱의 좌표·서식·밀도 기준선이자, 스냅샷 회귀 게이트(519도형 전수 대조)의 기준 산출물이며, 실전 장이 변형해 출발하는 원형(golden/adapted/novel)이다. 이 장은 goldenfab 빌드 구조, 변형 채택사, 2026-07-16 5부 개편(S17 제외·S18 이동), 스냅샷 회귀, adapted/novel 2경로를 다룬다. 오딧·게이트 규칙의 세부는 8장에 위임한다(상세: `8_QA-GATES.md`).

## 7.1 파이프라인 속 위치

```
FINAL-REPORT ─▶ [① GRILL] ─▶ [② compose] ─▶ [③ build] ─▶ [④ QA]
                                 │              │            │
                          레이아웃 매칭      golden.*      audit_deck
                          (15타입 결정표)    렌더 조회     밀도 밴드
                                 └──────┬───────┴────────────┘
                                        ▼
                        ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
                        ┃  골든 덱 (이 장)            ┃
                        ┃  goldenfab · 스냅샷 519도형 ┃
                        ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
```

골든 덱은 파이프라인의 한 단계가 아니라 ②·③·④가 공유하는 기준선 층이다. ②의 레이아웃 매칭 결정표(상세: `4_DECK-COMPOSE.md`)는 registry 15타입을 분모로 삼고, ③의 빌더는 `golden.*` 타입을 registry에서 조회해 렌더하며, ④의 실전 오딧은 골든 스냅샷에서 밀도 하한을 파생한다.

## 7.2 골든 덱 정본 — 5부 17장 · 스냅샷 519도형

현행 정본은 **5부 17장**(커밋 4834953 HEAD; `.claude/skills/pptx-build/scripts/goldenfab/reference.py:62~80` SLIDE_ORDER 17항목; `golden/build_golden.py:1` 도크스트링)이고, 스냅샷은 **17장 519도형**(`.claude/skills/pptx-build/assets/golden-snapshot.json` 실측)이다. 장별 구성:

| #   | 장 이름(스냅샷 names) | 레이아웃 타입  | 도형 수 |
| --- | --------------------- | -------------- | ------- |
| S1  | 표지                  | cover          | 8       |
| S2  | 목차                  | toc            | 31      |
| S3  | 간지1                 | part           | 11      |
| S4  | 문제정의              | problem_grid   | 48      |
| S5  | 간지2                 | part           | 11      |
| S6  | 실행그래프            | exec_graph     | 43      |
| S7  | 간지3                 | part           | 11      |
| S8  | 용어사전              | tech_evidence  | 42      |
| S9  | 지식그래프            | tech_tree      | 80      |
| S10 | 스크린샷              | screenshot     | 21      |
| S11 | 판단트리              | tech_mechanism | 46      |
| S12 | 구조화출력            | tech_capture   | 33      |
| S13 | 간지4                 | part           | 11      |
| S14 | 트러블슈팅            | ab_simulation  | 37      |
| S15 | 골든테스트            | validation     | 46      |
| S16 | 경계                  | boundary       | 26      |
| S17 | 클로징                | closing        | 14      |
| 합  | 17장                  | —              | **519** |

도형 종류 분포는 TEXT_BOX 226 · AUTO_SHAPE 216 · LINE 75 · PICTURE 2(golden-snapshot.json 실측). 목차(TOC)는 5부 구성(`reference.py:52~58`)이고 간지(part)는 스냅샷 실측 4장, 간지·클로징의 진행 도트 분모는 둘 다 5이지만 출처가 다르다: 간지는 `reference.py:98`이 `total=len(TOC)`로 파생하는 반면, 클로징은 `goldenfab/s21_closing.py:32` CLOSING_DEFAULTS의 `"total": 5` 리터럴 하드코딩이다(주석은 "완주 도트 수 = 목차 부 수 — 간지 도트와 같은 분모여야 한다"). 목차 부 수가 바뀌면 클로징만 수동 동기화가 필요한 이중 출처다.

콘텐츠 원천은 K-IFRS 1115 실프로젝트 실측 수치(`golden/00_factsheet.md` — "골든 스크립트는 이 수치만 사용, 창작 0")다. 단 이 팩트시트 문서 자체에는 낡은 기록이 남아 있다: 구성표가 구 6부 체계이고 간선 수치가 폐기값 2,697이다(정정 근거는 `_workspace_kifrs/07-s8-s9-verify.md` 기계 대조).

**미봉합 결함 — 골든 덱 내부 간선 수치 불일치.** 정정이 S9에만 반영됐다. 2026-07-19 실측(`grep -rn "2,69[47]"`):

| 파일:라인             | 출력 문자열                                 | 값            |
| --------------------- | ------------------------------------------- | ------------- |
| `s09_variants.py:61`  | "7종 2,629 + BC 근거층 65 = 2,694 · 고립 0" | **2,694**     |
| `s06_variants.py:394` | "노드 929 · 간선 2,697."                    | 2,697(폐기값) |
| `s18_variants.py:165` | `("간선", "2,697")`                         | 2,697(폐기값) |

즉 골든 덱 한 벌 안에서 S6·S18은 폐기값을, S9는 정정값을 찍고 있다. 스냅샷 회귀(compare_golden)는 "골든이 안 변했는지"만 보므로 이 불일치를 잡지 못하고, 앵커 수치 대조는 실전 덱 게이트(check_contract)에만 있어 골든 자신은 검사 대상이 아니다. 미해결 상태다. (같은 폐기값이 미배선 실패 표본 `s06_proto_f.py:41`에도 있으나 골든 빌드에 물리지 않는다.)

**문서 잔존 주의**: "19장·6부" "490도형"은 개편 전 수치로, `docs/user/00·03·05·07·08`, `pptx-build SKILL.md:78`, golden-snapshot.json의 `_comment`, `compare_golden.py:182`, `reference.py:111`(build_reference 도크스트링), s12~s21 변형 파일 도크스트링("골든 19장") 등에 아직 남아 있다. 문서에는 그렇게 남아 있으나 실측 정본은 17장·519도형이다. 전환 경위는 12장 소관(상세: `12_JOURNEY.md`).

## 7.3 goldenfab 빌드 흐름

```
golden/build_golden.py (43줄 CLI)
        │
        ▼
reference.py ── TOC 5부(L52~58) · SLIDE_ORDER 17항목(L62~80)
        │        cover/toc/part만 K-IFRS content 주입, 나머지 content=None
        ▼
registry.py ── LAYOUTS 15종: 타입명 → 채택 변형 함수(variant_*)
        │
        ▼
변형 함수 렌더 ── 좌표·색·도형 고정 (자체 audit은 이 경로에서 안 돎 — 아래 주 참조)
        │
        ├─▶ golden/golden-deck.pptx  (눈검용 렌더 산출물, .gitignore)
        │
        ▼
compare_golden.py ◀──── assets/golden-snapshot.json (기준선 519도형)
        │
        ▼
golden/variants/compare_full.md  (17/17 · 519/519 · 불일치 0 — PASS)
```

- `golden/build_golden.py`는 **43줄 CLI**다(2026-07-19 `wc -l` 실측 — 커밋 2e96215 메시지의 "46줄"은 당시 값). 2026-07-15 골든 단일화(golden/ 코드 사본 git 추적 **21파일 8,250줄 삭제** — `git show 2e96215 --numstat`에서 golden/ 하위 전삭제 파일만 집계한 값)로 395줄에서 얇아졌고, 그 경위를 자기 도크스트링에 기록하고 있다. 도크스트링의 "17장(5부)"은 이 묶음에서 드물게 현행 정합인 문서다.
- `reference.py`의 build_reference가 골든 기본값(content=None)을 의도적으로 쓰는 **유일한 경로**다. part_content는 `total=len(TOC)`를 명시한다 — "5부 덱에 6점은 조용한 오류"(`reference.py:98`; `layouts.py:194`의 코드 기본값 `c.get("total", 6)`은 잔존).
- 실전 빌드(공장 문)는 이 경로를 거치지 않는다. `build_pptx.py:1336` `_render_golden`이 registry를 조회하되 `content_contract.assert_content`(L1344)로 content 전량 주입을 강제한다 — 누락·빈 값이면 ValueError로 빌드 중단. "골든 기본값은 회귀 게이트 전용 기준점 — 콘텐츠 원천이 아니다"(`goldenfab/registry.py:5~7`), "골든의 글과 숫자는 공장 문 밖으로 한 글자도 못 나간다"(gate-repair 스펙).
- **변형 파일의 자체 `audit`은 골든 빌드 경로에서 돌지 않는다.** `audit(prs)` 정의는 10개 파일에 있으나(s06:603·s08:637·s09:477·s10:194·s11:979·s12:436·s14:621·s15:686·s17:613·s18:467), 호출은 전부 같은 파일의 `main()`(그 파일을 직접 실행하는 시안 렌더 CLI) 안이다 — 예 `s06_variants.py:654`, `s18_variants.py:516`. `build_golden.py`·`reference.py`·`compare_golden.py`·`build_pptx.py` 어디에도 호출이 없다(2026-07-19 `grep -rn "audit(prs)"` 실측). 즉 이 오딧들은 시안을 깎는 동안만 돌고, 골든 덱 빌드나 공장 문 빌드에는 물리지 않는다.

## 7.4 구조 표 — registry 15종 전수

`goldenfab/registry.py:24~40` LAYOUTS의 타입 → 채택 변형 매핑 전수(15/15 content dict 수용, `_workspace_kifrs/param-audit.md` 실측):

| #   | 타입           | 채택 변형          | 출처 파일                   | 골든 덱(SLIDE_ORDER) |
| --- | -------------- | ------------------ | --------------------------- | -------------------- |
| 1   | cover          | (정형 — 시안 아님) | goldenfab/layouts.py        | 포함(S1)             |
| 2   | toc            | (정형)             | goldenfab/layouts.py        | 포함(S2)             |
| 3   | part           | (정형)             | goldenfab/layouts.py        | 포함(간지 4장)       |
| 4   | problem_grid   | variant_k          | goldenfab/_variant_k.py     | 포함(S4)             |
| 5   | exec_graph     | variant_c          | goldenfab/s06_variants.py   | 포함(S6)             |
| 6   | tech_evidence  | variant_c          | goldenfab/s08_variants.py   | 포함(S8)             |
| 7   | tech_tree      | variant_a          | goldenfab/s09_variants.py   | 포함(S9)             |
| 8   | screenshot     | variant_a          | goldenfab/s10_screenshot.py | 포함(S10)            |
| 9   | tech_mechanism | variant_d          | goldenfab/s11_variants.py   | 포함(S11)            |
| 10  | tech_capture   | variant_b          | goldenfab/s12_variants.py   | 포함(S12)            |
| 11  | ab_simulation  | variant_c          | goldenfab/s14_variants.py   | 포함(S14)            |
| 12  | validation     | variant_c          | goldenfab/s15_variants.py   | 포함(S15)            |
| 13  | mirror_matrix  | variant_c          | goldenfab/s17_variants.py   | **제외**(창고 유지)  |
| 14  | boundary       | variant_b          | goldenfab/s18_variants.py   | 포함(S16)            |
| 15  | closing        | variant_a          | goldenfab/s21_closing.py    | 포함(S17)            |

경로 접두 `goldenfab/` = `.claude/skills/pptx-build/scripts/goldenfab/`. 15종 전부 `fn(prs, c)` 서명으로 content dict를 받고, 좌표·색·도형은 코드 고정·텍스트만 override하는 content/layout 분리 계약이다(param-audit 15/15 실측).

### 시안 체계 — 채택본·미채택 잔재·미배선 실험

각 장은 복수 시안(variant_a·b·c·d)을 렌더해 비교한 뒤 하나만 registry에 별칭 승격하는 구조다. 이 시안 이력에 붙은 수치는 검증 등급이 갈린다:

| 주장                 | 근거                                                              | 등급                                            |
| -------------------- | ----------------------------------------------------------------- | ----------------------------------------------- |
| 시안 렌더 누적 106장 | `docs/user/00-INDEX.md:35` "시안 렌더 106장은 `golden/variants/`" | **문서 표기값** — 코드로 재측정 불가            |
| S4 최종 48도형       | `assets/golden-snapshot.json` S4 실측                             | 실측                                            |
| S4 31→48도형         | `.claude/state/contracts/b43b2e72….md:195·233`(스냅샷 갱신 기록)  | 문서 표기값 — 구 31도형 산출물은 남아 있지 않음 |
| S4 시안 A~K 16버전   | `docs/user/03-golden-deck.md:12` "S4는 시안 A~K, 16버전"          | 문서 표기값                                     |
| S4 반려 4회          | 출처 없음. `design-rules.md:161` D-2 절 제목은 "**반려 5회** 끝"  | `미검증` — 문서끼리 충돌(4 vs 5)                |

즉 확실한 것은 최종 48도형뿐이고, 나머지 이력 수치는 문서 기록에만 있으며 반려 횟수는 두 문서가 서로 다른 값을 말한다. 채택되지 않은 코드의 현재 상태를 정직하게 구분하면:

| 구분             | 대상                                                                                        | 상태(실측)                                                                                                                                                                                                                                                                                                                                 |
| ---------------- | ------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 미채택 시안 잔재 | s06 a·b / s08 a·b / s11 a·b·c / s12 a / s14 a·b / s15 a·b / s17 a·b / s18 a (**15개 함수**) | registry 미등록 — 각 파일 main() 경유 시안 렌더만 가능, c override 미지원(골든 문자열 하드코딩). 이력 보존용                                                                                                                                                                                                                               |
| 죽은 레이아웃    | `goldenfab/_variant_h.py`의 variant_h(S4 시안 H)                                            | registry 미등록·호출 0곳. 파일은 `_shape_step` 헬퍼 공급원으로만 삶(s06_variants:15·s08_variants:16이 import)                                                                                                                                                                                                                              |
| 미배선 실험      | `goldenfab/s06_pilot.py`(파일럿 v6)                                                         | registry 미등록·**라이브 경로 import 0곳**(백업 스크립트 `golden/backup/s06_pilot_v3.py:17`만 `from goldenfab.s06_pilot import build`로 참조)·git 미추적(??) — 진행 중 실험. 골든 s06 원형+8개 교정(키커+14pt 헤드라인·하단 재서술 바 폐지→콘텐츠 패널 3·아이콘 ~30·동형 카드 4), 콘텐츠는 `golden/_pilot_s06/units.md` v4 재고 ~101단위만 |
| 실패 표본 보존   | `goldenfab/s06_proto_f.py`(프로토 F6)                                                       | registry 미등록·import 0곳·git 미추적. `docs/handoff-2026-07-18-golden-densify.md:63`이 "실패 표본(껍데기 먼저 → 속빈강정)"으로 명시. 브랜드킷 밖 NAVY "1B2A6B" 리터럴, NAVY_TINT 선언만 미사용. brand-kit `compact: 19`(미커밋)의 출처                                                                                                    |

s06_pilot v6·s06_proto_f는 채택본이 아니다 — 골든 빌드에 물리지 않으며, 편입 여부는 미결이다. 2026-07-18 고밀도 개편 실험(레퍼런스 해부 R1~R12 포함)은 전부 미커밋 워킹트리 상태다.

## 7.5 5부 17장 개편 — S17 제외 · S18 이동 (코드 실측 근거)

2026-07-16 커밋 4834953에서 차별점 부를 삭제하고 19장을 17장으로 재깎았다. 개편이 코드에 실제 반영됐는지의 실측:

**S17(7축 미러 매트릭스) 골든 제외 — 주장은 S14 결론 바 흡수, 타입은 창고 유지.**
- `reference.py:61` 주석과 SLIDE_ORDER(L62~80)에 mirror_matrix 없음 — "골든 덱에서 제외됐지만 레지스트리 타입으로는 유지(실전 덱 창고)".
- `registry.py:20`에 `from .s17_variants import variant_c as mirror_matrix` 유지 — 실전 덱이 골라 쓸 수 있는 창고 타입.
- `s14_variants.py:410` DEFAULT["bar"] 위 주석 — "구 S17(7축 비교) 삭제 후 그 주장을 이 바가 진다(2026-07-16)". 제외 사유는 셀 전수가 타 장 재탕(S14·S6·S4·S10)이라는 판정이었고, 그 전사(前史)로 미러 매트릭스 채움률 14~29% FAIL 9건이 러너 미등록 탓에 방치됐던 사고가 있다(과거형 — 이후 mirror_matrix·boundary를 오딧 러너에 등록해 봉합, 상세: `8_QA-GATES.md`).
- 잔재: `s17_variants.py:21` KICKER가 삭제된 부 이름 "5. 차별점" — 골든 밖이라 실해는 없으나 실전 편입 시 갱신이 필요하다.

**S18(정직한 한계 = 경계) Part Ⅳ 이동 · 경계 압축.**
- `s18_variants.py:23` KICKER = "4. 문제와 해결"(주석 "2026-07-16 5부 개편 — 한계 장은 검증 장 뒤 Part Ⅳ 소속"). 인용 수치(78/92·59.1%)의 원천인 검증 장(validation, S15) 뒤에 오도록 옮긴 것이다.
- 경계 안 상자는 내용 크기로 압축 — `s18_variants.py:279~281` 주석 "빈 상자의 원인은 상자가 내용보다 큰 것", 재깎기 좌표 L282~290(IN_W 4.6·IN_H 1.25 등).
- 한계 카드 3단(한계 bold/실측/→남은 과제) — DEFAULT limits 3종(헤지 2건 / 하드 인용 59.1% / 92건 비순수 홀드아웃, L243~276), 자체 audit이 한계 카드 ≠3이면 FAIL(L467~507).

## 7.6 골든 = 변형 가능한 출발점 — golden / adapted / novel

사용자가 "텍스트만 갈아 끼운 것"을 기각한 뒤(2026-07-16, 커밋 56e1e2f) 골든은 복제 템플릿이 아니라 **변형 가능한 출발점**으로 재정의됐다. deck-spec의 장 타입이 3갈래로 갈린다(`build_pptx.py` 디스패치; 3갈래 판정 기준은 `4_DECK-COMPOSE.md`의 매칭 결정표 소관):

```
deck-spec 장 type
 │
 ├─ golden.<layout> ──▶ registry 15종 조회 ──▶ content_contract.assert_content
 │                       (build_pptx.py:1336)     └─ 누락·빈 값 = ValueError 빌드 중단
 │                       골든 좌표·서식 그대로, 텍스트 전량 주입
 │
 ├─ adapted.<layout> ─▶ 장 스크립트 importlib 로드 (build_pptx.py:1367~1400)
 │                       sl["script"] 상대경로 필수 · build(prs, content) 필수
 │                       골든 장 스크립트를 내용에 맞게 변형한 사본
 │
 └─ novel ───────────▶ 장 스크립트 자유 형태 (같은 _render_scripted 경로, L1439 분기)
                        골든에 없는 새 레이아웃 — 창고 편입은 사용자 선택
 │
 └──(세 갈래 공통)──▶ audit_deck: 전역 오딧 + 밀도 밴드 (상세: 8_QA-GATES.md)
```

- **golden.\***: 내용이 골든 레이아웃에 그대로 맞는 장. content 전량 주입이 계약이고, 침묵 폴백은 content_contract가 막는다("키 존재는 주입이 아니다", `goldenfab/content_contract.py:94~95`).
- **adapted.\***: 내용이 형태를 바꿔야 하는 장 — 골든 장 스크립트를 변형한 스크립트를 spec이 지목하고, 빌더가 importlib로 로드해 `build(prs, content)`를 실행한다("장 스크립트가 레이아웃 본체다", `build_pptx.py:1381`). 골든 오염 검사(check_contract)는 adapted.\*에도 골든과 같은 잣대를 적용한다.
- **novel**: 골든 어휘에 없는 새 장. 형태는 자유지만 밀도는 골든에 앵커된다.

**밀도 밴드 — 스냅샷 파생 앵커.** `goldenfab/audit.py`의 density_band(L365 부근)가 golden-snapshot.json에서 fixed={cover,toc,part,closing} 제외 본문 장들의 도형 수·비어있지 않은 텍스트 프레임 수 **최솟값**을 파생한다 — 코드에 리터럴 0개, 골든이 바뀌면 하한도 따라 움직인다. 실전 장(golden.\* 정형 제외·adapted.\*·novel)은 `audit_deck.py`가 이 밴드+전역 오딧 8항목을 적용한다. 러너·규칙·병렬 채점의 세부는 8장 위임.

## 7.7 스냅샷 회귀 게이트 — compare_golden

goldenfab 자체의 보호 장치다(실전 장 검증은 audit_deck 몫 — 스냅샷 회귀는 goldenfab 보호 전용으로 축소, 2026-07-16 스펙). `compare_golden.py`가 지금 빌드한 골든 덱을 스냅샷과 전수 대조한다.

| 항목      | 값                                                                                                                                                                                            |
| --------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 비교 축   | 도형 수 / type / text / fill(솔리드 fore_color, enum 직접 대조 — str 비교 금지) / run(첫 런 pt·bold) / run_colors(전 런) / cells(표 셀) / l·t·w·h ±0.005" / names 순서(SLIDE_ORDER 변경 감지) |
| 최근 판정 | **17/17장 · 519/519도형 · 불일치 0 — PASS** + 파라미터 유효성 12/12 PASS 동봉 (`golden/variants/compare_full.md`, 5부 개편 후 재실행분)                                                       |
| 민감도    | MARGIN_L 0.6→0.61 변조 시 197건 검출 · brand-kit accent 변조 시 55건 검출(색 서명 확장 전 구 게이트는 0건) — gate-repair 스펙 실측                                                            |
| 기준선    | `--update-snapshot`으로만 재생성 — 스냅샷 손수정 금지(`_comment` 명시)                                                                                                                        |

파라미터 유효성 12종은 closing·screenshot 커스텀 content 반영 + 10종 레이아웃의 headline override 반영을 검사한다(cover·toc·part 3종은 param_efficacy 미포함 — 분모 12가 정본, "14/14"는 기억 오기로 정정된 기록). 알려진 공백: **폰트 이름은 비교하지 않는다**(크기·볼드만 — gate-repair 스펙·`docs/user/07-factory-port.md:4~5`, 미해결). 스냅샷 `_comment`의 "19장" 표기는 스테일이며 다음 `--update-snapshot` 때 갱신된다. 직전 FAIL 기록(19장 시절 493/507·불일치 검출)과의 대비는 13장 소관(상세: `13_TROUBLESHOOTING.md`).

또 하나의 정직한 제약: 골든 기준선 재생성은 **저자 머신에서만 가능**하다 — `ref/`가 gitignore이고 s10의 IMG_3D가 절대경로(`C:\...\k-ifrs-1115\images\knowledge_graph_3d.png`, `s10_screenshot.py:19`)라 클린 클론에서는 골든 빌드가 재현되지 않는다.

## 7.8 실증 예시 — S06 실행 그래프의 채택 흐름 (walked example)

S6 "실행그래프" 한 장이 시안에서 스냅샷까지 가는 여정을 실값으로 추적한다.

**① 변형 후보 3종.** `goldenfab/s06_variants.py`에 variant_a(사이드바형)·variant_b(풀와이드형)·variant_c(분기 포함 실행 그래프) 3안이 공존한다. main()이 3안을 모두 렌더해 `variants/s06_variants.pptx`로 저장 — 시안 비교·채점의 재료다(골든 전체로는 이런 시안 렌더가 누적 106장, 2026-07-11 시점).

**② 채택 → registry 등록.** variant_c가 채택돼 `registry.py:12`에서 `from .s06_variants import variant_c as exec_graph`로 별칭 승격됐다. 탈락한 variant_a·b는 삭제되지 않고 미채택 시안 잔재로 남는다(registry 미등록, c override 미지원).

**③ 채택본의 계약과 실값.** variant_c만 `c=None` 텍스트 override를 지원한다. 골든 기본 콘텐츠 DEFAULT_C(`s06_variants.py:375~411`)는 node_names 7개·details 4건이고, 좌표·색·도형은 코드 고정이다: LANE_Y 2.35·NODE_H 0.6, 노드 폭 [1.0, 1.45, 1.15, 1.45, 1.45, 1.45, 1.0], 다이아 높이 NODE_H+0.25, 거절 박스 2.0×0.45 @y3.55(L459~641). 자체 audit 규칙 3종 — bottom>6.35 초과·풀폭 우변≠12.733·accent≤4(`s06_variants.py:603~645`) — 은 정의돼 있으나, **호출은 같은 파일 `main()`의 `:654` 한 곳뿐이다.** 즉 `python s06_variants.py`로 시안을 렌더할 때만 돌고, 골든 덱 빌드(build_golden→reference)나 공장 문 빌드(build_pptx)에서는 돌지 않는다.

**④ 골든 덱 편입 → 스냅샷 박제.** `reference.py` SLIDE_ORDER의 6번째 항목("exec_graph", "S6 실행그래프")으로 골든 덱에 앉았고, 스냅샷에 **43도형**으로 박제됐다(7.2 표). 이후 매 빌드에서 compare_golden이 43개 도형의 type·text·fill·run·좌표(±0.005")를 대조하고, param_efficacy가 headline override 반영 여부까지 검사한다 — 최근 판정 17/17·519/519·불일치 0 PASS에 이 장도 포함이다.

**⑤ 계보는 닫히지 않았다.** 채택본 뒤에도 차세대 후보 `s06_pilot.py`(v6 — 골든 s06 원형+8개 교정, 재료는 units.md v4 ~101단위)와 실패 표본 `s06_proto_f.py`가 실험 잔재로 남아 있다. 둘 다 registry 미등록·미배선이며, 골든 편입 여부는 미결이다(7.4 표).

같은 패턴의 이웃 사례로 S08(tech_evidence, variant_c 채택)은 채택본에 방어 장치를 내장한다: BREAKDOWN 합=423 audit assert(`s08_variants.py:640`), 피치<칩 높이면 ValueError로 빌드 정지하는 `_pitch`(L397~416, "시끄럽게 죽인다"), 글자폭 파생 `chip_w`(한글 1자≈0.8×pt/72 — 오딧과 같은 자). 이 파생식은 두 장이 재사용하는 단일 출처가 됐는데 범위가 다르다 — `s09_variants.py:23`이 `_pitch`·`chip_w` 둘 다를, `s17_variants.py:442`가 함수 안 지역 import로 `_pitch`만 가져간다(`chip_w`는 미import).

## 7.9 정직한 현재 상태

- **골든 덱은 레퍼런스이지 실전 산출물이 아니다.** 실전 최종 .pptx 완성본은 없고, K-IFRS 파일럿은 2026-07-15 QA FAIL(계약 위반 17건) 후 동결 상태다(상세: `9_KIFRS-PILOT.md`).
- 2026-07-18 고밀도 개편(s06_pilot v6·proto_f·ref_anatomy R1~R12·brand-kit compact 19)은 전부 미커밋·미배선 — 진행 중 실험이다. R1~R12는 오딧 러너 미등록("러너에 안 물리면 없는 규칙" 상태).
- audit_golden의 레이아웃별 SPECS는 15종 중 7종만 등록 — exec_graph·tech_mechanism·ab_simulation·validation·closing은 레이아웃별 오딧 미등록이다(전역 규칙·자체 audit이 부분 보완, 상세: `8_QA-GATES.md`).
- `layouts.py`(toc)·`s21_closing.py`의 "회사 로고" 하드코딩은 content·brand-kit으로 못 고치는 알려진 빚(check_contract FORBIDDEN 등록).
- **골든 덱 내부 간선 수치 불일치 — 미해결.** S9는 정정값 2,694, S6·S18은 폐기값 2,697을 찍는다(§7.2 표). 어느 게이트도 골든 자신의 수치 정합을 검사하지 않는다.
- **변형 파일 자체 `audit`은 골든 빌드에 물리지 않는다.** 10개 파일에 정의돼 있으나 호출은 각 파일 `main()`(시안 렌더 CLI)뿐이다(§7.3). 시안을 깎을 때만 도는 오딧이다.
- **클로징 진행 도트 `total: 5`는 리터럴 하드코딩**이다(`s21_closing.py:32`). 간지는 `len(TOC)` 파생인데 클로징만 수동 동기화가 필요한 이중 출처다(§7.2).
- `s21_closing.py`의 SUMMARY 상수(결론 바 3문장)는 정의만 있고 미사용 — 도크스트링 서술과 코드 불일치(코드 정본). goldenfab 도크스트링 다수("골든 19장")도 스테일이다.
