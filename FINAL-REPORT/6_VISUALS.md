# 6. 시각 어휘층 — 차트·다이어그램·아이콘·목업 갤러리

> 이 장은 "슬라이드에 무엇을 그릴 수 있는가(어휘)"와 "그 어휘를 어떻게 고르고 승인받는가(추천→심사→갤러리)"를 다룬다.
> 구현 실물은 `.claude/skills/pptx-visuals/scripts/` 아래 7개 스크립트다 — pptx-visuals SKILL.md 서두에는 "구현은 3개"로 남아 있으나 실물은 7개(같은 문서 뒷절이 스스로 3개(render_real_mockups·check_candidates·make_icons)를 더 언급하고 recommend_archetypes는 아예 언급하지 않는 자기 불일치).

## 6.1 파이프라인 내 위치

```
① GRILL ──▶ ② deck-compose ──▶ ②.5 목업 승인 ──▶ ③ build_pptx ──▶ ④ QA
                 │                    │                 │              │
                 ▼                    ▼                 ▼              ▼
   ┌──────────────────────────────────────────────────────────────────────┐
   │ ★ 시각 어휘층 (.claude/skills/pptx-visuals/)                          │
   │   추천 엔진(②) · 후보 자기 심사(②) · 목업 갤러리(②.5)                 │
   │   렌더러 visuals·mpl_exhibits(③) · MPL_TYPES 집합은 QA 계수에도(④)    │
   └──────────────────────────────────────────────────────────────────────┘
```

시각 어휘층은 파이프라인의 한 단계가 아니라 ②~④를 관통하는 **가로층**이다. ②에서 후보를 추천·심사하고, ②.5에서 사용자 승인을 받고, ③에서 실제로 그리고, ④에서는 `MPL_TYPES` 집합이 "이미지가 정상인가"의 계수 기준으로 쓰인다. 절차 상세는 `상세: 4_DECK-COMPOSE.md`(②·②.5), 빌더 본체는 `상세: 5_PPTX-BUILD.md`, 게이트 전체는 `상세: 8_QA-GATES.md`.

## 6.2 어휘 인벤토리 (구조 표)

| 어휘층          | 종수                    | 목록                                                                                                                                                                          | 정본 출처                                                                                              | 낡은 서술(어디에 남아있나)                                                                                                  |
| --------------- | ----------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------- |
| 네이티브 차트   | **6종**                 | bar/hbar/line/pie/doughnut/stacked_bar (pie ≤5조각)                                                                                                                           | `.claude/skills/pptx-visuals/scripts/visuals.py:36~43` CHART_TYPES                                     | —                                                                                                                           |
| 도형 다이어그램 | **18종**                | flow·layers·branch·timeline·cards·from_to·process_band·band_table·matrix_2x2·spectrum·harvey_table·check_matrix·venn·pro_con·icon_rows·stat_split·contrast_split·split_detail | `visuals.py:181~220` add_diagram 디스패치                                                              | "다이어그램 2종 flow·layers(YAGNI)": pptx-visuals SKILL.md·visuals.py 도크스트링 L4·docs/PIPELINE.md                        |
| mpl 확장 차트   | **9종**                 | waterfall/heatmap/dumbbell/slope/funnel/annotated_scatter/histogram/**bubble/waffle** (220dpi PNG, 수치 편집 불가)                                                            | `.claude/skills/pptx-visuals/scripts/mpl_exhibits.py:25~35` MPL_TYPES                                  | "7종": mpl_exhibits 도크스트링 L3-4·pptx-build SKILL.md L29·CLAUDE.md v4 이력(bubble·waffle는 Deloitte p7·p24 실측 후 추가) |
| 아이콘          | **24종 × 3색 = 72 PNG** | Lucide(MIT) 큐레이션 24종, primary/accent/white 3색, 256px, CDN lucide-static@0.462.0 사전 박제                                                                               | `.claude/skills/pptx-visuals/scripts/make_icons.py:19~44(ICONS)·45~61`                                 | "Lucide 아이콘 72종": CLAUDE.md v3 이력(파일 수 기준 표기 — 종수는 24)                                                      |
| L-ID 카탈로그   | **38종(L01~L38)**       | 정량 네이티브 L01~06 · mpl L07~13 · 도형 L14~25 · 텍스트·표 L26~30 · v4.4 승격 L31~34 · v4.5 컴포지트 L35~38                                                                  | `.claude/skills/pptx-visuals/references/archetype-catalog.md`; `recommend_archetypes.py:24~63` LID_KEY | visual-selection.md B표 하단 3행은 카탈로그 밖(§6.9 갭 참조)                                                                |

네이티브 6종 + 도형 18종 = **PPT에서 직접 편집 가능한 어휘 24종**이 `visuals.py`(1,164행) 한 파일에 모여 있고, 여기에 mpl 9종(이미지)을 더한 것이 전체 렌더 어휘다. pptx-visuals SKILL.md의 선언 — "시각자료 코드는 이 모듈에만 존재한다" — 대로 빌더에는 차트 코드 중복이 없다.

카탈로그의 철칙은 "**모든 행은 렌더러가 실존한다**"(archetype-catalog.md) — 렌더러 없는 등재로 빌드 불가 spec이 나오는 것(hollow 카탈로그)을 막는 장치다. 이 철칙이 우회된 곳이 1군데 있다(§6.9).

## 6.3 시각 생성 흐름 — 후보에서 렌더까지

```
deck-spec 초안(장별 visual 후보)                        [② deck-composer]
   │
   ▼
recommend_archetypes.py ─ 실데이터 특성 7+1종 결정적 검출 ─▶ L-ID 랭킹
   │
   ▼
03_exhibit-candidates.json  (본문 장 × 3안, 조합장치 ≥1, 디폴트 어휘엔 why_not)
   │
   ▼
check_candidates.py ─ 검사 7종 + 게이트 통과 조합 시뮬레이션(≤200,000) ─ FAIL ─▶ 후보 재작성
   │ PASS
   ▼
render_real_mockups.py ─ 실물 pptx 빌드 + PowerPoint COM PNG 렌더
   │        └─ COM 불가 환경 ─▶ make_mockups.py (mpl 스케치 폴백)
   ▼
gallery.html ─ JS 게이트 4종 실시간 채점 + 그리디 기본선택 ─▶ 사용자 회신 "sNN=X"   [②.5 승인 게이트]
   │
   ▼
02_deck-spec.json 확정 ─▶ build_pptx.py ─┬─ CHART_TYPES 6종 ──▶ 네이티브 차트(수치 편집 가능)
                                          ├─ MPL_TYPES 9종 ───▶ 220dpi PNG(편집 불가)
                                          └─ add_diagram 18종 ─▶ 도형 다이어그램(편집 가능)
```

핵심 설계 의도: "다양하게·알맞게"라는 판단을 LLM 감각에 맡기지 않고 **결정적 규칙(추천)·exit code(심사)·사용자 눈(갤러리)**의 3중으로 기계화·외부화한 것이다. ①은 "무엇을 말할 것인가", ②.5는 "어떻게 보일 것인가"를 각각 승인받는다(pptmaker SKILL.md).

## 6.4 렌더러 세부 — visuals.py · mpl_exhibits.py (build_pptx 실배선)

### visuals.py — 편집 가능 어휘의 단일 출처

- **단일 소비자는 build_pptx다.** `build_pptx.py:28` import 후 add_icon(904·1011)·add_chart(1192·1224)·add_callouts(1227)·add_diagram(1266·1278·1285)에서 호출. 색·폰트·크기는 인자 brand dict(brand-kit.yaml)에서만 나온다 — hex 하드코딩 없음.
- **emphasis = BCG 규칙**: emphasis가 있으면 전 시리즈 회색(GRAY_BASE `#C7CBD1`, `visuals.py:95`) + 강조만 accent 1색(`visuals.py:98~131`; 단일 시리즈 bar/hbar는 포인트 단위 강조). 없으면 accent→primary→muted 순환. 자동 제목 제거, 데이터 레이블은 bar/hbar/line/stacked_bar만.
- **add_callouts**: 이미지에 굽지 않는 pptx 도형 콜아웃(3.0×0.55in, 영역 밖 탈출 방지 클램프 `visuals.py:144`) — 네이티브 경로에서는 주석도 편집 가능 객체다.
- **add_icon**: 사전 변환 PNG 삽입, 없는 이름은 ValueError — "조용한 누락 금지"(`visuals.py:29~32`).
- **add_diagram 18종 디스패치**(`visuals.py:181~220`), 미지 레이아웃 ValueError. stat_split/contrast_split/split_detail은 _icon_rows를 재사용하는 복합 골격이고 split_detail은 add_diagram을 재귀 호출한다(`visuals.py:419`). v4.5 플랫 룩: 노드 직각+무테(`visuals.py:995`), flow 화살촉은 XML 직접 부착(`visuals.py:982~986`).
- 구현 우회의 예: 하비볼 부분 채움은 python-pptx가 노출하지 않는 기능이라 PIE 도형 각도 조정(-90°+90°×score)으로 구현(`visuals.py:755~783`).
- 엣지: `_from_to` 열 머리 "From — 기존 방식"/"To — 본 시스템"이 하드코딩(`visuals.py:944`) — 전환 서사가 아닌 데이터에 쓰면 문구가 안 맞는다.

### mpl_exhibits.py — 하이브리드 이미지 경로

- 네이티브 차트가 표현 못 하는 9종을 matplotlib 220dpi(`mpl_exhibits.py:22`) PNG로 렌더한다. **MPL_TYPES 집합 자체가 라우팅 스위치** — build_pptx는 chart_type이 이 집합에 있으면 이미지 경로로 전환한다(`build_pptx.py:1256`). 배선은 3곳: 빌더(`build_pptx.py:27` import·`:1208` render), QA(`audit_pptx.py:22`), 목업(`make_mockups.py:25`).
- 스타일 단일 경로: 모든 렌더러가 `_fig()`(폰트·스파인·눈금 강제)로 시작해 `_save()`(annotations 콜아웃+note 각주+220dpi 저장)로 끝나고, 색은 회색 베이스+accent 1색만(`_emph_colors`, `mpl_exhibits.py:141~144`). "전부 회색, 강조 데이터만 브랜드색 1개 + 콜아웃 주석 + 하단 출처 각주가 BCG 룩의 최대 지렛대"(모듈 주석 L8~9).
- 렌더러별 자동 장치: waterfall은 증감분 입력에 합계 막대 자동 추가+단차 점선(`:147~178`), heatmap은 흰색→accent 단색 그라데이션(`:184`), funnel은 단계 간 전환율 ↓% 표기(`:279~281`), slope는 강조 시리즈만 굵은 선(`:225~234`). 미지 타입 ValueError(`:135`). 한글 폰트 우선순위 Pretendard→Malgun Gothic→NanumGothic(`:41~43`).
- 산출물은 이미지라 수치 편집 불가 — 수정은 deck-spec 경유 재빌드가 공식 경로다.

## 6.5 추천·심사·갤러리 — 목업 계열 4종 (파이썬 import 0곳, 스킬 절차서 CLI 배선)

이 절의 스크립트들은 §6.4와 배선 형태가 다르다. **build_pptx가 import하는 것은 visuals·mpl_exhibits뿐**이고, 목업 계열(check_candidates·render_real_mockups·make_icons는 파이썬 import 0곳, make_mockups는 render_real_mockups가 부품 CSS·edit_label만 import — `render_real_mockups.py:20`)은 스킬 절차서(deck-compose SKILL 등)가 실행 명령으로 배선한 **CLI 도구**다. 즉 "자동으로 실행된다"가 아니라 "절차서가 이 시점에 이 명령을 돌리라고 지정한다"가 정확한 현재형이다.

### recommend_archetypes.py — 결정적 추천 엔진

실데이터에서 특성 7+1종(has_numbers 수치≥3·time_axis·sum100 합 95~105·pair2·grouped·stagey·kpi·funnel_like 단조감소)을 **전부 결정적 규칙으로** 검출(`recommend_archetypes.py:105~146`)하고, classify 우선순위(`:149~167`)로 형상을 분류한 뒤 SHAPE_RANK(`:66~77`)로 L-ID를 랭킹한다. 소비자는 `check_candidates.py:38` 1곳 — 빌더·QA에는 미배선(후보 심사 단계 전용).

### check_candidates.py — 후보 사양서 자기 심사기

deck-composer가 사용자 보고 전에 스스로 돌려 PASS/FAIL(exit 0/1)을 받는 CLI(배선: deck-compose SKILL:96·108). 검사 7종 — ①본문 슬라이드 커버 전수(양방향 차집합) ②슬라이드당 후보 ≥3안 ③슬라이드 내 후보 유형 중복 금지 ④후보마다 조합장치 ≥1 ⑤shape 필드 필수 ⑥디폴트 어휘(bar·flow·cards) 후보의 why_not 의무 ⑦현행 재출품 금지. 추가로 추천군 밖 후보는 warn(FAIL 아님), 그리고 **다양성 게이트를 통과하는 선택 조합이 존재하는지를 product 전수 시뮬레이션(≤200,000 조합)으로 사전 확인**한다(`check_candidates.py:98~122`). 다양성을 문장이 아니라 exit code로 강제하는 장치다.

### render_real_mockups.py — ②.5 승인 게이트의 기본 경로

후보 전수를 **실제 build_pptx.Deck으로 빌드**해 candidates.pptx를 만들고 PowerPoint COM(SaveAs 포맷 18=PNG)으로 전장 PNG 렌더 후 인터랙티브 gallery.html을 생성한다(배선: deck-compose SKILL:99, 기본 경로). mpl 스케치와 달리 최종 산출물과 동일한 룩(마스터 틀·폰트·색)을 보여준다. 갤러리에는 다양성 게이트 4종(쿨다운 간격≥3·최소 5종·박스≤30%·동일 유형≤2회)을 JS로 실시간 재계산하는 세트 심사 패널이 있고(`render_real_mockups.py:195~202`), 기본 선택은 게이트 전부 통과 조합을 그리디로 사전 계산한다(페널티: 쿨다운 +100·동일 유형 ≥2회 +100·박스 30% 초과 +50·기유형 재사용 +3, `:108~118`). Windows+PowerPoint COM 전제, 실패 시 exit 1 후 make_mockups 폴백. 사용자는 "s07=B" 형식으로 회신하고, 회신 전 deck-spec 확정은 금지다.

주의: pptmaker SKILL.md는 ②.5 기본을 make_mockups.py로 서술하나 **현행 기본은 render_real_mockups.py**(deck-compose·pptx-visuals SKILL) — 오케스트레이터 문서 쪽이 낡았다.

### make_mockups.py — COM 불가 환경 폴백

mpl 스케치 PNG(5.2×2.9in)+gallery.html 생성기(배선: deck-compose SKILL:102 폴백). 원칙은 "차트는 실데이터, 다이어그램은 스케치"(`make_mockups.py:203`) — 룩 미리보기와 데이터 검증을 동시에 한다. 후보마다 편집 가능성 3라벨(native=PPT 수치편집가능 / image=재빌드 수정 / shape=편집가능)을 판정해 카드에 명시한다(`:35~43`).

## 6.6 아이콘 파이프라인 — make_icons.py

```
Lucide CDN(lucide-static@0.462.0) ─▶ SVG 캐시(icons/_svg) ─▶ 브랜드 hex 치환 ─▶ pymupdf 256px PNG
                                                                                  │
                          visuals.add_icon(없는 이름 ValueError) ◀─ assets/icons/ 박제 ◀─┘
                                    │
                                    ▼
                               build_pptx 렌더
```

일회성 CLI(파이썬 import 0곳, 배선: pptx-visuals SKILL:69~70 재실행 절차)로 24종×3색(primary·accent는 brand-kit에서, white 고정)=72 PNG를 `pptx-build/assets/icons/`에 사전 박제한다 — 빌드 시 네트워크 의존 제거. **브랜드 색이 바뀌면 이 스크립트만 재실행하는 ripple 단일 경로**다. 실행에는 네트워크(unpkg CDN)가 필요하고, 다운로드 실패 이름은 모아 출력 후 exit 1(`make_icons.py:85~87`). 산출물 기준으로는 `visuals.py:24`(ICON_DIR)→add_icon→build_pptx로 이어지는 실배선이다.

## 6.7 배선 지도 (7 스크립트 전수)

| 스크립트                | 배선 형태                                                      | 근거                                                                            |
| ----------------------- | -------------------------------------------------------------- | ------------------------------------------------------------------------------- |
| visuals.py              | **build_pptx 실배선** (단일 소비자)                            | `build_pptx.py:28` import, add_chart/add_diagram/add_icon/add_callouts 호출 8곳 |
| mpl_exhibits.py         | **실배선 3곳** (빌더+QA+목업)                                  | `build_pptx.py:27·1208·1256`, `audit_pptx.py:22`, `make_mockups.py:25`          |
| recommend_archetypes.py | 라이브러리 — 소비자 1곳(후보 심사 전용)                        | `check_candidates.py:38`                                                        |
| check_candidates.py     | 파이썬 import 0곳 — CLI(스킬 절차서 배선)                      | deck-compose SKILL:96·108                                                       |
| render_real_mockups.py  | 파이썬 import 0곳 — CLI(②.5 기본 경로)                         | deck-compose SKILL:99                                                           |
| make_mockups.py         | render_real_mockups가 부품(CSS·edit_label)만 import — CLI 폴백 | `render_real_mockups.py:20`; deck-compose SKILL:102                             |
| make_icons.py           | 파이썬 import 0곳 — 일회성 CLI(산출물만 실배선)                | pptx-visuals SKILL:69~70; `visuals.py:24`가 산출물 소비                         |

## 6.8 실증 예시 (walked example) — K-IFRS S6 익스히빗 후보의 여정

K-IFRS 1115 파일럿(상세: `9_KIFRS-PILOT.md`)에서 시각 어휘층이 실제로 굴러간 궤적을 S6(실행 그래프 장) 한 건으로 추적한다.

1. **계약(①)**: 2026-07-14 사용자가 아웃라인 계약 `_workspace_kifrs/01.5_outline.md`를 확정 — 확정 이미지 3종 배치를 계약에 박았다(screenshot_answer.png→S6, knowledge_graph_3d.png+screenshot_graph_node.png→S9).
2. **후보 사양서(②)**: deck-composer가 `_workspace_kifrs/03_exhibit-candidates.json`에 본문 8장(S4·6·8·9·10·11·13·15) × 3안 = **24후보**를 작성, 각 장 A안 preferred. S6는 A안 `golden.screenshot`(screenshot_answer.png 실물 이미지) 대 B·C 네이티브 대안 구도.
3. **갤러리(②.5)**: 목업 갤러리 산출물 실물이 `_workspace_kifrs/mockups/gallery.html`로 남아 있다 — 갤러리 단계까지 실행된 물증. 단 gallery.html의 s09 기본 체크는 B(표)인데 계약은 실이미지 확정 — 계약 확정 이전 산출물이거나 갱신 누락(`미검증`).
4. **확정·빌드(③)**: `_workspace_kifrs/02_deck-spec.json` 16장으로 확정·빌드. (이 spec은 본문 content 미주입 등 별개 결함을 안고 있었다 — 상세: `9_KIFRS-PILOT.md`.)
5. **결말(④)**: 2026-07-15 check_contract가 계약·spec·pptx 3자 대조에서 **계약 이미지 FAIL 2건(파일 3종: S6 1·S9 2)** — 계약이 S6에 박은 screenshot_answer.png가 최종 pptx에 실리지 않았음을 md5 대조로 검출. QA 총판정 FAIL 17건의 일부로 기록됐고 이후 재빌드 없이 동결 상태다.

교훈: 시각 어휘층의 승인 게이트(②.5)는 "어떤 그림을 쓸지"까지는 확정하지만, 그 채택이 최종 pptx에 관철됐는지는 ④의 계약 대조(check_contract)가 잡는다 — 어휘층과 QA층의 역할 경계가 이 실패 사례에 그대로 드러난다.

## 6.9 알려진 갭 (정직 목록)

| #   | 갭                                                                                                                                                                                                                                                          | 실측 근거                                                                                             |
| --- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| 1   | **visual-selection.md B표 하단 3행(process_grid·category_spine_table·mapping)은 렌더러 미실존** — visuals.py grep 0건·L-ID 없음. 이 표를 따른 composer는 빌드 불가 spec을 산출한다. 카탈로그 철칙("모든 행은 렌더러 실존")이 이웃 문서에서 우회된 유일 지점 | `.claude/skills/pptx-visuals/references/visual-selection.md` B표 vs `visuals.py`·archetype-catalog.md |
| 2   | **SHAPE_RANK 추천 커버리지 35/38** — L15(layers)·L26(table)·L30(bullets)은 어떤 데이터 형상에도 추천되지 않는다(카탈로그에는 있으나 추천 엔진이 밀지 않는 폴백 취급)                                                                                        | `recommend_archetypes.py:66~77` SHAPE_RANK 합집합                                                     |
| 3   | **다양성 게이트 정의 3곳 중복** — audit_pptx.diversity_checks / render_real_mockups JS(`:195~202`) / 그리디 페널티(`:108~118`)가 별도 구현. 드리프트 위험. 게다가 visual-selection.md C-2는 "3조항·연속 2장 금지"로 낡음(코드는 4종·간격≥3으로 더 엄격)     | notes 실측; `audit_pptx.py:44~65`(코드 정본 4종)                                                      |
| 4   | **check_candidates SHAPES 죽은 상수** — 형상 10종 정의만 있고 참조 0(shape 값 검증 미구현). 또 시뮬레이션 L100이 `sl["visual_candidates"]` 직접 접근이라 fixed 슬라이드에 키가 없으면 KeyError — 미완 흔적                                                  | `check_candidates.py:22~33`·`:100`                                                                    |
| 5   | 문서 지연 다발 — SKILL "구현은 3개"(실물 7개)·"다이어그램 2종 YAGNI"(실측 18종), mpl 도크스트링 "7종"(실측 9종), pptmaker SKILL ②.5 기본 도구 서술 낡음(§6.5)                                                                                               | 각 §의 낡은 서술 열 참조                                                                              |
| 6   | make_mockups 갤러리 크롬 CSS의 accent(#D66E3A)가 brand-kit이 아닌 하드코딩 — 브랜드 변경 시 목업 이미지는 따라가지만 갤러리 UI 색은 안 따라간다                                                                                                             | `make_mockups.py:220~226`                                                                             |

---

*이전: `5_PPTX-BUILD.md`(빌더 본체) · 다음: `7_GOLDEN-DECK.md`(골든 덱). 시각 다양성 게이트의 판정 주체는 `8_QA-GATES.md`.*
