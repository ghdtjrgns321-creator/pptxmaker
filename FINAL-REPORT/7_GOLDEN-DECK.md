# 7. 골든 덱 — 조합 사례이자 회귀 기준

> 골든 덱은 손으로 깎아 확정한 **5부 17장짜리 레퍼런스 프레젠테이션**이다. 역할은 둘이다 — 부품을 어떻게 조합·배치했나의 **조합 사례**, 그리고 goldenfab을 고쳤을 때 무엇이 깨졌는지 알려주는 **회귀 기준**(797도형 전수 대조). 오딧·게이트 규칙의 세부는 8장에 위임한다.

## 7.1 골든이 하는 일 — "정답지"가 아니다

```
FINAL-REPORT ─▶ [① GRILL] ─▶ [② compose] ─▶ [③ build] ─▶ [④ QA]
                                 │              │            │
                          부품·틀 선택      composed·golden  audit_deck
                                 └──────┬───────┴────────────┘
                                        ▼
                        ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
                        ┃  골든 덱 (이 장)            ┃
                        ┃  goldenfab · 스냅샷 797도형 ┃
                        ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
```

| 골든이 하는 일         | 어떻게                                                                     |
| ---------------------- | -------------------------------------------------------------------------- |
| **조합 사례 제공**     | 새 덱은 골든을 **참고**해 부품을 다시 조합한다(텍스트 교체가 아니다)       |
| **부품화 무손실 증명** | 장 함수를 부품 호출로 바꿔도 `compare_golden`이 **픽셀 동일**이어야 통과   |
| **밀도·기하 기준선**   | `audit_golden`(장별 정답지 SPECS) · `preflight_dense` · 배치 틀 6종의 좌표 |

**"이대로 글자만 갈아끼우는 정답지"가 아니다.** 그렇게 쓰던 시절의 실측이 있다 — 실전 덱의 타입 분포가 골든 `SLIDE_ORDER`와 그대로 겹쳤다. 그래서 2026-07-27에 골든의 역할을 "잣대"에서 "조합 사례 + 회귀 기준"으로 재정의하고, 도해를 부품으로 꺼내 `composed` 경로를 열었다(`상세: 5_PPTX-BUILD.md` §5.4.1).

빌드: `uv run python golden/build_golden.py` → `golden/golden-deck.pptx`.

## 7.2 골든 덱 정본 — 5부 17장 · 797도형

| #   | 장 이름    | 레이아웃 타입  | 도형 수 | 부품 배선                       |
| --- | ---------- | -------------- | ------- | ------------------------------- |
| S1  | 표지       | cover          | 8       | (프레임)                        |
| S2  | 목차       | toc            | 31      | (프레임)                        |
| S3  | 간지1      | part           | 11      | (프레임)                        |
| S4  | 문제정의   | problem_grid   | 62      | `gate_branch` + `fan_in`        |
| S5  | 간지2      | part           | 11      | (프레임)                        |
| S6  | 실행그래프 | exec_graph     | 43      | `routing_lane`                  |
| S7  | 간지3      | part           | 11      | (프레임)                        |
| S8  | 용어사전   | tech_evidence  | 97      | `bipartite_map` + `card_row`    |
| S9  | 지식그래프 | tech_tree      | 135     | `relation_catalog` + `card_row` |
| S10 | 스크린샷   | screenshot     | 22      | —                               |
| S11 | 판단트리   | tech_mechanism | 104     | `numbered_steps` + `n_branch`   |
| S12 | 구조화출력 | tech_capture   | 75      | —                               |
| S13 | 간지4      | part           | 11      | (프레임)                        |
| S14 | 트러블슈팅 | ab_simulation  | 37      | —                               |
| S15 | 골든테스트 | validation     | 60      | —                               |
| S16 | 경계       | boundary       | 65      | —                               |
| S17 | 클로징     | closing        | 14      | (프레임)                        |
| 합  | **17장**   | —              | **797** | 부품이 물린 장 **6장**          |

도형 종류 분포(스냅샷 실측): AUTO_SHAPE 330 · TEXT_BOX 318 · PICTURE 80 · LINE 68 · CHART 1. PICTURE 80은 대부분 Lucide 아이콘이다 — 고밀도 개편으로 아이콘 사용이 늘면서 도형 수가 519 → 797로 올랐다.

콘텐츠 원천은 K-IFRS 1115 실프로젝트 실측 수치다. **납품물이 아니다** — 공장 문(content_contract)이 골든 기본값의 유출을 막으므로 실제 프로젝트 덱에는 한 글자도 안 나간다. 여기 텍스트의 쓸모는 **분량**이다: 실물급 길이여야 레이아웃(줄바꿈·박스 높이)이 제대로 검증된다.

## 7.3 dense 승격 — 골든이 둘로 갈라져 있던 병

2026-07-26에 골든을 **하나로** 합쳤다. 그전에는 운영·열람용이 `golden-deck-operating.pptx`(dense)인데 registry·회귀 하네스는 sparse 변형을 지켰다. 그래서 도해마다 sparse 1벌 + dense 1벌이 존재했고(가로 좌표는 같고 세로만 달랐다 — s08 실측: 2.45/4.75/8.20 공통, 세로 3.88 vs 1.30), **장 하나 고치려면 두 벌을 고쳐야 했다.**

이것은 2026-07-15에 없앤 `golden/`↔`goldenfab/` 이중화와 **같은 병이 다른 축에서 재발한 것**이다. 승격으로 sparse 기준선은 폐기했고, 회귀 스냅샷과 장별 기하 오라클(`audit_golden.SPECS`)을 dense 좌표로 재수립했다.

registry 15종의 현행 매핑:

| #   | 타입           | 구현 모듈         | 골든 덱               |
| --- | -------------- | ----------------- | --------------------- |
| 1~3 | cover·toc·part | `layouts.py`      | S1·S2·간지 4장        |
| 4   | problem_grid   | `s04_dense.py`    | S4                    |
| 5   | exec_graph     | `s06_variants.py` | S6                    |
| 6   | tech_evidence  | `s08_dense.py`    | S8                    |
| 7   | tech_tree      | `s09_dense.py`    | S9                    |
| 8   | screenshot     | `s10_dense.py`    | S10                   |
| 9   | tech_mechanism | `s11_dense.py`    | S11                   |
| 10  | tech_capture   | `s12_dense.py`    | S12                   |
| 11  | ab_simulation  | `s14_dense.py`    | S14                   |
| 12  | validation     | `s15_dense.py`    | S15                   |
| 13  | mirror_matrix  | `s17_variants.py` | **제외**(실전 창고용) |
| 14  | boundary       | `s16_dense.py`    | S16                   |
| 15  | closing        | `s21_closing.py`  | S17                   |

**부채 1건 — exec_graph(S6)만 dense 기준작이 미승인이다.** sparse 렌더러(`s06_variants.py`)를 유지하고 있고, 기준작 후보 `s06_mid.py`(373줄)가 골든 미편입 상태로 남아 있다. 파일럿(밀도 최대)과 골든(공기 많음) 사이의 중간지점으로 사용자 확정 구성까지 나왔으나 승격되지 않았다.

## 7.4 부품화 — 정리가 곧 부수 효과였다

도해를 부품으로 꺼내는 작업이 **정리 작업이 됐다.** 정본 dense 장이 sparse **시안 파일**(`*_variants.py`)에 얹혀 글과 도해 함수를 빌려 쓰고 있어서 시안을 지울 수 없었다. 도해를 부품으로 꺼내고 글을 content로 밀어내자(계약 ③) 정본이 시안을 부를 이유가 사라졌고, 8개 파일 4,600여 줄이 삭제됐다:

`_variant_h` · `_variant_k` · `s06_variants`(시안 2종) · `s08_variants` · `s09_variants` · `s10_screenshot` · `s11_variants` · `s11_branch_snap` · `s12_variants` · `s14_variants` · `s15_variants` · `s18_variants`.

결과: `goldenfab` 최상위가 10,640 → **6,507줄**(−39%). 여기에 새로 생긴 `figures/` 3,103줄이 더해진다 — 줄 수가 줄었다는 것보다 **재사용 가능한 비율이 4.0%에서 도해 전체로 바뀌었다**는 것이 요점이다.

전 회차가 `compare_golden` 픽셀 동일로 통과했다. 한 번 걸렸던 것도 그 증거다 — S11에서 3도형 불일치가 났는데, 원인은 원래 데이터의 설명 문자열에 앞 공백이 있었고 부품이 하나를 더 붙여 이중 공백이 된 것이었다. 눈으로는 못 봤을 차이를 게이트가 잡았다.

**아직 장 함수에 남은 도해 5종** — 지울 시안이 없어 정리 이득이 없고, 개수 파생으로 바꾸면 그림이 달라져 픽셀 동일이 깨지는 자리다(설계 판단 필요):

- **S9 위계 트리 · S11 판정 트리** — 노드 좌표가 **손으로 맞춘 값**이다(글자 폭과 무관)
- **S15 격리 시험 · S16 경계 · S17 미러 매트릭스** — 장 전면을 쓰는 구성이라 `box` 분리가 애매

## 7.5 goldenfab 빌드 흐름

```
golden/build_golden.py (CLI)
        │
        ▼
reference.py ── TOC 5부 · SLIDE_ORDER 17항목
        │        cover/toc/part만 K-IFRS content 주입, 나머지 content=None
        ▼
registry.py ── LAYOUTS 15종: 타입명 → 구현 모듈
        │
        ▼
장 함수 렌더 ── 자리를 정하고, 그림은 figures 부품이 그린다
        │
        ├─▶ golden/golden-deck.pptx  (눈검용 렌더 산출물, .gitignore)
        │
        ▼
compare_golden.py ◀──── assets/golden-snapshot.json (기준선 797도형)
        │
        ▼
golden/variants/compare_full.md  (17/17 · 797/797 · 불일치 0 — PASS)
```

`reference.build_reference`가 골든 기본값(content=None)을 의도적으로 쓰는 **유일한 경로**다. 실전 빌드(공장 문)는 이 경로를 거치지 않는다 — `_render_golden`이 registry를 조회하되 `content_contract.assert_content`로 content 전량 주입을 강제한다.

**소스 이중화를 없앤 이력이 이 파일에 박혀 있다.** 원래 `golden/`(동결 원본)과 `goldenfab/`(공장)에 같은 레이아웃 코드가 두 벌 있었다. 이식이 충실한지 1회 검증하려던 비계였는데 안 걷었고, "수정은 goldenfab에서만" 규칙도 깨져 **1회성 검증이 영구 동기화 세금**이 됐다. 2026-07-15에 21파일 8,250줄을 삭제하고, 소스를 복사해 두는 대신 **결과물(도형 서명)을 스냅샷으로 박아** 비교하는 골든 파일 테스트의 정석으로 바꿨다. 복사본은 독립 검증력이 없고(그냥 사본) 동기화 비용만 든다.

## 7.6 스냅샷 회귀 게이트 — compare_golden

goldenfab 자체의 보호 장치다(실전 장 검증은 audit_deck 몫 — 스냅샷 회귀는 goldenfab 보호 전용).

| 항목      | 값                                                                                                                                                                                            |
| --------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 비교 축   | 도형 수 / type / text / fill(솔리드 fore_color, enum 직접 대조 — str 비교 금지) / run(첫 런 pt·bold) / run_colors(전 런) / cells(표 셀) / l·t·w·h ±0.005" / names 순서(SLIDE_ORDER 변경 감지) |
| 최근 판정 | **17/17장 · 797/797도형 · 불일치 0 — PASS** + 파라미터 유효성 12/12 PASS 동봉                                                                                                                 |
| 민감도    | MARGIN_L 0.6→0.61 변조 시 197건 검출 · brand-kit accent 변조 시 55건 검출(색을 못 보던 구 게이트는 같은 변조에 0건)                                                                           |
| 기준선    | `--update-snapshot`으로만 재생성 — 스냅샷 손수정 금지                                                                                                                                         |

파라미터 유효성 12종은 closing·screenshot 커스텀 content 반영 + 10종 레이아웃의 headline override 반영을 검사한다(cover·toc·part 3종은 미포함 — 분모 12가 정본).

**알려진 공백 2건.** ① **폰트 이름은 비교하지 않는다**(크기·볼드만). ② 스냅샷 `_comment`가 아직 "19장"으로 적혀 있다 — 다음 `--update-snapshot` 때 갱신된다.

또 하나의 정직한 제약: 골든 기준선 재생성은 **저자 머신에서만 가능하다** — `ref/`가 gitignore이고 S10의 이미지가 타 프로젝트 절대경로라 클린 클론에서는 골든 빌드가 재현되지 않는다.

## 7.7 실증 — S8 용어사전 한 장이 부품으로 갈라지는 과정

**① 장이 무엇을 소유하나.** `s08_dense.py`는 헤더·소제목·자리·출처선을 소유한다. 도해와 카드는 소유하지 않는다 — "이 장이 정하는 것은 **자리**뿐"이라고 파일 자신이 적어 둔다.

**② 도해를 부품에 넘긴다.** 실무어 집합과 기준서 개념 집합의 대응은 `figures.bipartite_map`이 그린다. 장은 `Box`(x, y, w, h)와 데이터만 넘기고, 항목이 몇 개인지에 따라 좌표를 파생하는 것은 부품의 일이다.

**③ 카드 띠도 부품에 넘긴다.** 하단 4장은 `figures.card_row`가 그린다. "장은 자리만 정하고 **장수 파생은 부품이 한다**" — 카드가 3장이 되든 5장이 되든 장 코드는 안 바뀐다.

**④ 픽셀 동일로 증명한다.** 이 교체 후 `compare_golden`이 S8 **97/97 일치**를 냈다. 부품으로 꺼내면서 그림이 한 점도 안 변했다는 뜻이고, 그것이 "부품화가 손실 없음"의 유일한 증명이다.

**⑤ 그리고 그 부품이 다른 덱에서 다시 쓰인다.** `bipartite_map`은 이제 `composed` 장이 임의 데이터로 부를 수 있다. 장 함수였을 때는 불가능했던 것이고, 그래서 골든에 없는 형태가 손코딩이나 범용 박스 폴백으로 샜다.

## 7.8 정직한 현재 상태

- **골든 덱은 레퍼런스이지 실전 산출물이 아니다.** 이 파이프라인으로 ④ QA까지 통과한 실전 .pptx는 아직 0건이다(`상세: 9_KIFRS-PILOT.md`).
- **exec_graph(S6) dense 기준작 미승인** — sparse 렌더러 유지, `s06_mid.py` 미편입(§7.3).
- **골든 도해 5종이 아직 장 함수 안에 있다**(§7.4). 손으로 맞춘 좌표라 개수 파생 전환은 부품화가 아니라 재설계다.
- **`hub_spoke`·`layer_stack`이 살 틀이 없다** — 창고에 있는데 배치 틀 6종이 전부 가로 밴드+하단 카드 구조라 앉힐 자리가 없다. 배치 틀 도구함을 따로 채워야 한다.
- **`audit_golden`의 레이아웃별 SPECS는 15종 중 7종만 등록**(`problem_grid`·`screenshot`·`tech_capture`·`tech_evidence`·`tech_tree`·`mirror_matrix`·`boundary`). 나머지는 전역 규칙만 받는다(`상세: 8_QA-GATES.md`).
- `layouts.py`(toc)·`s21_closing.py`의 "회사 로고" 하드코딩은 content·brand-kit으로 못 고치는 알려진 빚(check_contract FORBIDDEN 등록).
- **클로징 진행 도트 `total: 5`는 리터럴 하드코딩**이다. 간지는 `len(TOC)` 파생인데 클로징만 수동 동기화가 필요한 이중 출처다.
- 스냅샷 `_comment`·일부 도크스트링의 "19장" 표기는 스테일이다(실측 17장).

## 7.9 교차 링크

- 부품 계약 3조건·부품 10종의 실체: [6_VISUALS.md](6_VISUALS.md)
- 부품·틀을 고르는 규칙: [4_DECK-COMPOSE.md](4_DECK-COMPOSE.md)
- `composed` 렌더 경로·GRID·frames 좌표: [5_PPTX-BUILD.md](5_PPTX-BUILD.md)
- 오딧·회귀·배선 게이트 전 층: [8_QA-GATES.md](8_QA-GATES.md)
