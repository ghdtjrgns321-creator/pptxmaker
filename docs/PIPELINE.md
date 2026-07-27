# pptmaker v6 파이프라인 — 부품 도구함 · 골든 조합 사례 · FINAL-REPORT → 덱

> **이 문서가 "지금 어떻게 도나"의 단일 출처다.** 계획·근거·반려 이력은 `docs/` 안 날짜
> 문서들이 맡고, 여기에는 **현재 상태만** 현재형으로 적는다.

콘텐츠는 사용자가 프로젝트별 `FINAL-REPORT/`(보고서형 md 묶음)로 미리 정리해둔다. 이 하네스는
그 FINAL-REPORT를 목차·강조점 인터뷰(GRILL)로 셀링 골격에 앉히고 **일관된 골격·브랜드의
네이티브 .pptx로 빌드**한다. 추출·NotebookLM 단계는 없다 — FINAL-REPORT가 이미 사람이
작성·검증한 as-built 단일 재료 원천이라, 밀도·출처·검증이 확보돼 있다(같은 소스 A/B 실측:
직접 추출 재료 6,246단어 ≈ k-ifrs FINAL-REPORT 6,098단어로 동급). 서식은 변함없이
코드(`build_pptx.py` + `brand-kit.yaml`)만이 보장한다.

## 두 개의 흐름 — 덱을 찍는 축 · 부품을 벼리는 축

이 프로젝트에는 **서로 다른 두 흐름**이 있고, 읽을 때 이 둘을 섞으면 안 된다.

```
  [생산 축] 재료 → 덱                     [제작 축] 레퍼런스 → 부품
  ────────────────────────────           ────────────────────────────
  FINAL-REPORT                            실물 슬라이드(ref/ · 웹)
      │ ① GRILL                               │ part-design 스킬
      ▼                                       ▼
  아웃라인 계약                            1층 원소  elements.py
      │ ② composer                            │  (선·칩·카드 — 자리를 모른다)
      ▼                                       ▼
  deck-spec  ◀─ 부품을 **고른다** ──      2층 도해  figures/*.py
      │ ③ builder                             │  (draw(slide, box, data, kit))
      ▼                                       ▼
  deck.pptx                                3층 배치  골든덱 = 조합 사례
      │ ④ QA                                     (부품을 어떻게 앉혔나의 예시)
      ▼
  results/*.pptx
```

- **생산 축**은 사용자가 PPT를 요청할 때 돈다(`pptmaker` 오케스트레이터).
- **제작 축**은 창고를 채울 때만 돈다(`part-design` 스킬). 덱을 만들지 않는다.
- 둘이 만나는 지점은 **②**다 — composer가 창고에서 부품을 고른다. 창고에 없는 물성이 나오면
  즉흥하지 말고 제작 축으로 넘긴다(그 즉흥이 v4까지 "범용 박스 폴백"으로 샌 경로다).

### 3층이 각각 무엇인가

| 층           | 실물                            | 무엇을 소유하나                                   | 무엇을 소유하지 않나                |
| ------------ | ------------------------------- | ------------------------------------------------- | ----------------------------------- |
| **1층 원소** | `goldenfab/figures/elements.py` | 선·칩·카드·리본 같은 최소 단위. 색·크기는 kit에서 | 자기 좌표(부르는 쪽이 정한다)       |
| **2층 도해** | `goldenfab/figures/*.py` (부품) | `box` 안 **그래픽 하나**. 계약 3조건을 지킨다     | 페이지 구성(설명 패널·배너·열) · 글 |
| **3층 배치** | `goldenfab/reference.py` → 골든 | 어떤 부품을 어디에 얼마나 크게 앉히나             | 도해 자체(2층에서 꺼내 쓴다)        |

**부품 계약 3조건**(단일 출처 `figures/__init__.py`): ① 자리 독립 `draw(slide, box, data, kit)`
② 개수 파생(좌표 리터럴 zip 금지) ③ 글 무소유(문자열 상수 0).
시각을 정하는 절차는 `part-design` 스킬 — 지어내지 않고 실물을 찾아 옮긴다.

### 골든덱의 역할 — 잣대가 아니라 **조합 사례**

`golden/golden-deck.pptx`(17장)는 "이대로 글자만 갈아끼우는 정답지"가 **아니다.**
부품을 어떻게 조합·배치했나의 **예시**이고, 동시에 **회귀 기준**이다:

| 골든이 하는 일     | 어떻게                                                                   |
| ------------------ | ------------------------------------------------------------------------ |
| 조합 사례 제공     | 새 덱은 골든을 **참고**해 부품을 다시 조합한다(텍스트 교체가 아니다)     |
| 부품화 무손실 증명 | 장 함수를 부품 호출로 바꿔도 `compare_golden`이 **픽셀 동일**이어야 통과 |
| 밀도·기하 기준선   | `audit_golden`(장별 정답지 SPECS) · `preflight_dense`                    |

빌드: `uv run python golden/build_golden.py` → `golden/golden-deck.pptx`.
장 정의는 `goldenfab/registry.py`(LAYOUTS 15종) + `reference.py`(SLIDE_ORDER 17장).

### 현재 진척 (2026-07-27 실측)

| 항목                     | 수                                                           |
| ------------------------ | ------------------------------------------------------------ |
| 카탈로그 등재 (A) 문법   | 23종 (`archetype-catalog.md` G01~G23)                        |
| 부품으로 꺼낸 도해       | **9종** (아래 표)                                            |
| 골든 장에 부품이 물린 수 | **6장** — S4·S6·S8·S9·S11(2종)                               |
| 무손실 증명              | `compare_golden` 17/17 슬라이드 · **797/797 도형 픽셀 동일** |
| `goldenfab` 줄수         | 10,640 → **6,048** (−43%) — 시안 8파일 삭제                  |

| 부품               | 어휘   | 물성                                  | 항목 수     |
| ------------------ | ------ | ------------------------------------- | ----------- |
| `gate_branch`      | G01    | 판정 하나에서 결과가 갈린다           | 갈래 2 고정 |
| `fan_in`           | G02    | 원인 N개가 한 곳에서 만난다           | 파생        |
| `routing_lane`     | G04    | 한 줄로 흐르다 판정에서 일부가 빠진다 | 파생        |
| `bipartite_map`    | G05    | A 집합과 B 집합이 갈라지고 모인다     | 파생        |
| `relation_catalog` | G08    | 무엇이 무엇을 잇고 근거가 어디서 왔나 | 파생        |
| `numbered_steps`   | G10    | 순서가 메시지인 절차                  | 파생        |
| `n_branch`         | G11    | 갈래 수 자체가 메시지                 | 파생        |
| `hub_spoke`        | 미등재 | 하나에 무엇이 물려 있나(1:N)          | 파생        |
| `layer_stack`      | 미등재 | 층이 몇이고 각 층이 무엇인가          | 파생        |

**부품화가 곧 정리였다.** 정본 dense 장이 sparse **시안 파일**(`*_variants.py`)에 얹혀 글과
도해 함수를 빌려 쓰고 있어서 시안을 지울 수 없었다. 도해를 부품으로 꺼내고 글을 content로
밀어내자(계약 ③) 정본이 시안을 부를 이유가 사라졌고, 8개 파일 4,600여 줄이 삭제됐다:
`_variant_h` · `_variant_k` · `s06_variants`(시안 2종) · `s08_variants` · `s09_variants` ·
`s10_screenshot` · `s11_variants` · `s11_branch_snap` · `s12_variants` · `s14_variants`.

**아직 장 함수에 남은 도해** — 지울 시안이 없어 정리 이득이 없고, 개수 파생으로 바꾸면
그림이 달라져 픽셀 동일이 깨지는 자리다(설계 판단 필요):

- S9 위계 트리 · S11 판정 트리 — 노드 좌표가 **손으로 맞춘 값**이다(글자 폭과 무관)
- S15 격리 시험 · S16 경계 · S17 미러 매트릭스 — 장 전면을 쓰는 구성이라 box 분리가 애매

**부품에 번호를 붙이지 않는다**(2026-07-26). 한때 `G22`~`G28`을 부품 ID로 썼는데 카탈로그
G번호와 충돌했다 — 카탈로그 G22는 `accent 강조 패널`, G23은 `소제목 + 룰`로 이미 점유 중이었고
G24~G28은 카탈로그에 없는 번호였다. 그 6종(`arc_links`·`relation_matrix`·`flow_split`·
`orbit_rings`·`pillar_base`·`bracket_fan`)은 전부 눈검증 반려로 삭제됐다. 부품의 유일 키는
**모듈 이름(=파일명)**이고, 카탈로그 어휘와 대응되면 `META["어휘"]`에 그 ID를 적는다
(`bipartite_map` → `G05`, 나머지는 `미등재`). 근거·규칙: `figures/__init__.py`.
설계·근거: `docs/2026-07-26-부품-도구함-재설계.md`(계획·이력). **현재 상태는 이 문서가 정본.**

---

## 생산 축 상세 — 데이터 흐름 (전체)

```
[사용자]                         [Claude Code — pptmaker 오케스트레이터]
프로젝트별 FINAL-REPORT/     ┌──────────────────────────────────────────────────┐
(보고서형 md 묶음)           │ ① deck-outline-grill (목차·콘텐츠 인터뷰 게이트)   │
  │                         │    FINAL-REPORT 기반 목차 1안 → 장별 강조 채록      │
  ▼                         │    → 01.5_outline.md (아웃라인 계약)                │
<프로젝트>/FINAL-REPORT/ ──▶ │            ▼                                       │
  *.md 를 재료로 지정         │ ② deck-composer                                    │
                            │    FINAL-REPORT + 계약 → 골격 배치 → 02_deck-spec   │
                            │            ▼                                       │
                            │ ②.5 익스히빗 승인 게이트 (gallery.html)            │
                            │            ▼                                       │
                            │ ③ pptx-builder                                     │
                            │    build_pptx.py + brand-kit → deck.pptx           │
                            │            ▼                                       │
                            │ ④ consistency-qa                                   │
                            │    audit_pptx.py → PASS / FAIL(되돌림 1회)          │
                            └──────────────────────────────────────────────────┘
                                          ▼
                            results/<프로젝트명>-소개.pptx (최종 산출물)
```

## 단계별 역할

| 단계 | 에이전트/주체        | 스킬               | 입력                                    | 출력                                            |
| ---- | -------------------- | ------------------ | --------------------------------------- | ----------------------------------------------- |
| ①    | 오케스트레이터(대화) | deck-outline-grill | FINAL-REPORT/\*.md                      | 01.5_outline.md (아웃라인 계약)                 |
| ②    | deck-composer        | deck-compose       | FINAL-REPORT + 01.5_outline.md          | 03_exhibit-candidates.json(시각 후보)           |
| ②.5  | 사용자 승인 게이트   | pptx-visuals       | 후보 JSON → 갤러리                      | mockups/gallery.html → 회신 → 02_deck-spec.json |
| ③    | pptx-builder         | pptx-build         | 02_deck-spec.json                       | deck.pptx (하이브리드: 네이티브+mpl PNG)        |
| ④    | consistency-qa       | consistency-qa     | deck.pptx + 02_deck-spec + FINAL-REPORT | 03_qa-report.md (+다양성 게이트 4종)            |

## 원칙 (박제)

1. **재료는 FINAL-REPORT 단일 원천.** 사용자가 프로젝트별로 정리해둔 보고서형 md가 재료다.
   추출·NotebookLM 단계 없음 — FINAL-REPORT가 사람이 작성·검증한 as-built라 밀도·출처·검증이
   이미 확보돼 있다. FINAL-REPORT가 얇으면 ① GRILL 인터뷰에서 사용자에게 보강 재료(추가
   문서·스크린샷)를 요청한다(억지 창작 금지).
2. **구성은 빌드 전 인터뷰에서 확정.** GRILL이 FINAL-REPORT 기반 목차 1안을 던지고, 장별
   "가장 전하고 싶은 것"을 채록해 `01.5_outline.md`(계약)를 만든다. deck-composer는 이 계약을
   지킨다 — 장 추가·삭제·채록 메시지 변경 금지.
3. **수정은 항상 deck-spec 경유 재빌드.** 완성 pptx를 직접 뜯지 않는다 — deck-spec이 단일
   출처여야 "브랜드 바꿔 재빌드" 재현성이 유지된다. 차트·다이어그램 레시피는 `pptx-visuals`
   스킬이 단일 출처(`scripts/visuals.py`).
4. **브랜드 단일 출처는 `brand-kit.yaml`.** 색·폰트·크기·여백은 이 파일만 고친다.
5. **근거는 FINAL-REPORT에 대응.** 추출 단계가 없으므로 수치·고유명사의 진위는 ④
   consistency-qa가 "덱이 FINAL-REPORT와 일치하나"로 대조한다(창작 0 원칙).
6. **도해는 창고에서 꺼낸다 — 즉흥 금지.** 필요한 물성의 부품이 없으면 손으로 그리지 말고
   멈추고 제작 축(`part-design`)으로 넘긴다. 즉흥이 v4까지 "범용 박스 폴백"으로 샌 경로이고,
   그렇게 만든 도해는 다음 덱이 상속하지 못한다(한 번 쓰고 버려진다).
7. **부품은 시각을 지어내지 않는다.** 실물 레퍼런스를 찾아 옮긴다. 근거: 부품 1종에
   10라운드를 쓰고 자체 발명한 잉크 10종이 전부 폐기됐다(2026-07-26). 절차는 `part-design`.

## 다이어그램 DSL (deck-spec `diagram` 타입)

python-pptx에는 다이어그램 객체가 없으므로 도형+커넥터 조합으로 렌더한다. **살아 있는 레이아웃
9종**(정본은 `pptx-visuals/scripts/visuals.py:add_diagram` — 이 목록은 그 분기의 사본이다):

`matrix_2x2` · `spectrum` · `harvey_table` · `check_matrix` · `venn` · `timeline` ·
`icon_rows` · `stat_split` · `split_detail`

`layout`은 필수이고 기본값이 없다 — 누락·오타는 빌드 시점에 `ValueError`로 죽는다(조용한 폴백
금지. 그게 어휘가 3종으로 수렴한 기계적 원인이었다).

**박스+화살표 계열 9종은 2026-07-25 폐기**: `flow`·`process_band`·`layers`·`branch`·`cards`·
`from_to`·`band_table`·`pro_con`·`contrast_split`. 팔레트 전수 렌더에서 전면 반려됐고("그냥 도형
나열 / 고등학생 ppt") 실전 사용은 0건이었다. 이 물성들의 정본은 (A) 골든 도해다 — 분기·라우팅
`G01`·`G04`·`G09`·`G11` · 집합 대응 `G05` · 관계 전수 `G08` · 단계 `G10`·`G14` · 병렬 카드
`dense.hero_card` · 경계 `G19`. (B)에 남는 **정당한** 쓸모는 (A)에 어휘가 없는 물성뿐이다 —
시간축 · 좌표계 · 다축 충족도 · 교집합.

상세 스키마: `.claude/skills/pptx-build/references/deck-spec-schema.md`

## 목표 품질 기준 (완성본 수준)

- 참고: EY Price Point / Dallas 시정 보고서 급 — 글·분석 밀도가 높은 회사 발표용
- 슬라이드당 120~200 단어(BCG 실측), 헤드라인 1줄 + 거버닝 메시지 1줄
- YOUNG한 대학 발표 스타일 금지, 데이터·근거 중심

## 여정 (변경 이력)

- v1 (2026-07-06): project-analyst가 코드를 직접 읽어 사실 추출 → selling-curator가 창작 큐레이션.
  첫 실전에서 "쓸만한 내용 추출·문서화·시각화" 품질이 부족해 폐기.
- v2 (2026-07-08): 콘텐츠 생산을 NotebookLM으로 이관. project-analysis/project-analyst 폐기,
  selling-curation → deck-compose(통합·선별로 역할 축소), content-extract·pptx-visuals 신설.
- v3 (2026-07-08): Claude 직접 추출을 메인으로 회귀(A/B 실측 밀도·출처 우위), NotebookLM은 보조.
  PART 탭·간지·계층번호·시각 어휘·QA 게이트 도입.
- v4 (2026-07-08): 시각 다양성 파이프라인. 첫 실전 덱이 "막대 1·표 3·박스플로우" 3어휘로
  수렴한 문제를 ②.5 사용자 승인 게이트·mpl 하이브리드·다양성 게이트로 해결. 롤백 지점:
  git tag `v3-pre-visual-diversity`.
- v5 (2026-07-12): **전반부 완전제거.** content-extract 스킬·content-extractor 에이전트·NotebookLM
  경로(모드 B·notebooklm-prompt)를 폐기하고, 사용자가 프로젝트별로 정리해둔 `FINAL-REPORT/`를
  단일 재료 원천으로 채택. GRILL(deck-outline-grill)이 ① 첫 단계가 되고, deck-compose의 초안
  통합·중복제거·상충병기 기계는 단일 소스라 제거. 사라진 grep 수치검증은 ④ QA의 FINAL-REPORT
  대조로 흡수. 근거: FINAL-REPORT(k-ifrs 6,098단어)가 이미 추출 재료(6,246단어)와 동급 밀도라
  "보고서에서 보고서를 다시 뽑던" 중복 단계를 걷어냈다.
- v6 (2026-07-26): **부품 도구함 · 3층 분리.** 골든 11,867줄 중 재사용 가능한 부품이 479줄
  (4.0%)뿐이었고 도해가 전부 장 함수 안에 굳어 있어 **임의 데이터로 부를 수 없었다** — 그래서
  골든에 없는 형태는 손코딩 아니면 범용 박스 폴백으로 샜다(사용자 반려: "골든덱에 쓰인 적 없는
  쓰레기 목업"). 원소·도해·배치를 층으로 갈라 도해를 `figures/`로 꺼내고, 골든의 역할을
  "잣대"에서 **"조합 사례 + 회귀 기준"**으로 재정의. 같은 날 골든 단일화(dense 승격)로
  sparse/dense 이중 기준도 제거(도해마다 2벌 유지하던 것을 1벌로).
  시각 결정 절차는 `part-design` 스킬로 분리 — 부품 시각은 **지어내지 않고 실물에서 옮긴다**.
