# 10. 차별점 — "AI에게 PPT 만들어줘"와 무엇이 다른가

> **요지:** 이 프로젝트의 차별점은 슬라이드를 "생성"하는 능력이 아니라, 완성도 기준(골든 덱)·서식 단일 출처(brand-kit)·콘텐츠 계약(content_contract)·기계 게이트(오딧 5계열)를 **코드로 박제해 재현 가능하게 만든 것**이다. 단, 이 프로젝트는 **미완성**이며 실전 파일럿 완주 산출물은 0건이다 — §3 한계가 그 목록이다.

## 0. 파이프라인 내 위치

이 장은 특정 단계가 아니라 파이프라인 전 구간에 걸친 **차별 장치들의 층**을 다룬다.

```
 FINAL-REPORT ─▶ [①GRILL] ─▶ [②compose] ─▶ [②.5승인] ─▶ [③build] ─▶ [④QA] ─▶ (results/…pptx)
      │             │            │             │            │           │           │
      │        ═════╪════════════╪═════════════╪════════════╪═══════════╪═════      │
      │        ★ 이 장: 각 단계에 박힌 차별 장치(계약·앵커·SSOT·게이트)를 횡단 정리   │
      └─ 재료 단일화                └ 골든 매칭 3갈래         └ brand-kit   └ 기계 게이트
```

단계별 상세는 `3_OUTLINE-GRILL.md`~`8_QA-GATES.md`, 실전 파일럿의 현재 상태(FAIL·동결)는 `9_KIFRS-PILOT.md`.

## 1. 비교 표 — 통상적 접근과의 차이

| 관점          | 통상적 접근(AI에게 "PPT 만들어줘")                        | 본 프로젝트                                                                                                               | 차이가 만드는 효과                                                                                                 |
| ------------- | --------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| 완성도 기준   | "예쁘게·전문적으로" 같은 산문 지시 — 기준이 세션마다 증발 | 손으로 깎은 골든 덱 **5부 17장·519도형**을 스냅샷(`golden-snapshot.json`)으로 박제, 모든 변경은 전수 대조                 | 완성도가 취향이 아니라 재현 가능한 기준점이 된다(17/17장·519/519 불일치 0 실측)                                    |
| 산출물 형태   | 이미지·HTML 슬라이드 — 받은 뒤 수치 수정 불가             | python-pptx 기반 **네이티브 .pptx**(표·차트가 편집 가능한 개체)                                                           | 덱을 받은 쪽이 PowerPoint에서 데이터만 고칠 수 있다                                                                |
| 서식 일관성   | 프롬프트마다 색·폰트가 흔들림                             | `brand-kit.yaml` **단일 출처(SSOT, Single Source of Truth)** — 코드 소비자 9종이 전부 이 파일에서 색·폰트·크기를 받음     | 파일 하나 수정 = 전 덱 일괄 반영; accent 변조 실험에서 게이트가 55건 검출(§4 실증)                                 |
| 품질 검증     | 육안 확인                                                 | **기계 게이트 5계열**(compare_golden·check_contract·audit_golden·audit_deck·audit_pptx) + 제3자 채점                      | 좌표 0.01" 변조도 197건 검출(아래 §2-3); "빌드됐다"와 "계약대로다"가 분리됨                                        |
| 콘텐츠 진실성 | LLM 창작이 수치에 혼입                                    | 재료는 사용자가 정리한 FINAL-REPORT **단일 원천**, 창작 금지 + `content_contract`가 골든 기본값 유출을 빌드 중단으로 차단 | 골든 글자가 실전 덱으로 새는 사고를 코드가 막는다(단, 이 강제는 2026-07-15 사고 후 신설 — `13_TROUBLESHOOTING.md`) |
| 구성 결정     | 모델이 목차를 임의 구성                                   | ①GRILL 인터뷰 계약(`01.5_outline.md`) — 장 추가·삭제·채록 변경 금지, QA가 계약 대조                                       | 구성 반려가 빌드 후가 아니라 텍스트 단계에서 끝난다                                                                |
| 수정 방식     | 완성 파일을 직접 뜯어 고침                                | 모든 수정은 deck-spec/brand-kit 경유 **재빌드** — 같은 spec = 같은 결과                                                   | 수정이 누적 드리프트를 만들지 않는다                                                                               |

## 2. 특장점 — 근거 있는 항목만

각 항목에 근거(코드 경로·실측·상술 장)를 붙였다. 근거 없는 항목은 싣지 않았다.

### 2-1. 골든 덱 앵커 (상세: `7_GOLDEN-DECK.md`)
- 골든 덱 5부 17장을 도형 단위로 스냅샷 박제: 519도형(TEXT_BOX 226·AUTO_SHAPE 216·LINE 75·PICTURE 2), `.claude/skills/pptx-build/assets/golden-snapshot.json`.
- 최근 회귀 판정 **17/17장·519/519도형·불일치 0 PASS** + 파라미터 유효성 12종 PASS — `golden/variants/compare_full.md`.
- 골든은 복제 템플릿이 아니라 **변형 가능한 출발점**: 매칭 결과에 따라 `golden.<layout>`(content 전량 교체) / `adapted.<layout>`(장 스크립트 변형) / `novel`(골든 밀도 앵커) 3갈래 — `.claude/skills/deck-compose/references/layout-matching.md`, `build_pptx.py:1336`·`:1367`.

### 2-2. 네이티브 pptx — 이미지가 아니다 (상세: `5_PPTX-BUILD.md`, `6_VISUALS.md`)
- `build_pptx.py` RENDERERS 레거시 11종(L1322~1334) + 골든 계열 3종 디스패치(L1336·L1367·L1439).
- 시각 어휘: 네이티브 차트 6종(`visuals.py` CHART_TYPES L36~43) + 도형 다이어그램 18종(add_diagram 디스패치 L181~220) + mpl 확장 9종(220dpi 이미지 — 이 9종만 수치 편집 불가, `mpl_exhibits.py` L25~35) + Lucide 아이콘 24종×3색=72 PNG(`make_icons.py`).
- 카탈로그 L01~L38 전 행에 "렌더러 실존" 철칙 — `archetype-catalog.md`.

### 2-3. 기계 게이트 — 육안이 아니라 코드가 판정 (상세: `8_QA-GATES.md`)
- `compare_golden.py`: 스냅샷 전수 대조(도형 수·type·text·fill·런 pt/bold·전 런 색·표 셀·좌표 ±0.005").
- 민감도 실측: MARGIN_L 0.6→0.61 변조 시 **197건 검출**, brand-kit accent 변조 시 **55건 검출**(구 게이트는 0건) — `docs/superpowers/specs/2026-07-15-gate-repair-design.md` L189·L211-213.
- `check_contract.py` 검사 6종(장 수·제목·골든 오염·계약 이미지 md5·금지어 5종·앵커 수치) — 파일럿 FAIL 17건을 실제로 검출한 게이트.
- `audit_golden.py` 검사 규칙 11종 + `--selftest` **11케이스**(검사기가 결함을 정말 잡는지 자기 검증) — 모듈 도크스트링 표기는 "2초", 2026-07-19 재실행 실측은 1.16초.
- `audit_deck.py` generic_checks 8항목 + **밀도 밴드**(리터럴 0 — 골든 스냅샷에서 최솟값 파생, 골든이 바뀌면 기준도 따라 이동), `goldenfab/audit.py` density_band.
- `audit_pptx.py` 검사 13항(다양성 게이트 4종 포함).

### 2-4. 브랜드 단일 출처 (상세: `5_PPTX-BUILD.md`)
- `.claude/skills/pptx-build/assets/brand-kit.yaml` — 색 6종(primary #15171B·accent #D66E3A 등)·폰트 3종·크기 **8단**(작업 트리에는 미커밋 실험값 `compact: 19`가 더해져 9단).
- 코드 소비자 **9종** 실측(2026-07-19 `grep -rln "load_kit("` + brand-kit 경로 참조): build_pptx·goldenfab/kit.py(유일 로더)·visuals·render_real_mockups·make_mockups·make_icons·audit_pptx·audit_golden·audit_deck. `check_contract.py`는 소비자가 아니다 — brand-kit을 로드하지 않고 위반 안내 문자열에서 이름만 언급한다(L32·L35).
- `goldenfab/kit.py`가 hex·pt 리터럴 금지 하에 파생색(mix)·한글 어절 줄바꿈 강제(eaLnBrk="0"+KOREAN)·fit_picture(실비율 읽기)를 제공.

### 2-5. 재료 단일화 — FINAL-REPORT (상세: `2_PIPELINE.md`, `3_OUTLINE-GRILL.md`)
- 추출·NotebookLM 전반부를 제거하고 사용자 유지 FINAL-REPORT를 단일 원천으로. 근거 실측: 직접 추출 6,246단어 ≈ K-IFRS FINAL-REPORT 6,098단어 동급(`docs/PIPELINE.md` L6-7) — "보고서에서 보고서를 다시 뽑는" 중복 제거.

### 2-6. 콘텐츠 계약 강제 (상세: `4_DECK-COMPOSE.md`, `8_QA-GATES.md`)
- `goldenfab/content_contract.py`: golden.* 렌더 시 content 전량 주입 강제 — 누락·빈 값이면 ValueError로 빌드 중단(`build_pptx.py:1344`). "키 존재는 주입이 아니다"(content_contract.py L94~95).
- 골든 기본값은 회귀 게이트 전용 — 공장 문(build_pptx)으로는 못 나간다(`goldenfab/registry.py` L5~7).

### 2-7. 검사기의 자기 검증 (상세: `8_QA-GATES.md`, `11_TEST-DECISIONS.md`)
- audit_golden `--selftest` 11케이스, verify_s8s9 결함 주입 5/5 검출·오탐 0, verify_content_contract 15타입×5케이스 exit 0 — "항상 PASS를 내는 검사기"를 걸러내는 층이 별도로 있다.

## 3. 한계·트레이드오프 — 정직 목록

### 3-1. 이 프로젝트는 미완성이다 — 실전 검증(파일럿 완주) 0건
- **게이트를 완주한 최종 산출물이 없다.** `results/`의 빌드 실물 4건은 전부 골든 체계 이전(2026-07-08)이거나 QA FAIL 판정분(07-14·07-15) — 완주 PASS 0건.
- 유일한 실전 시도인 K-IFRS 1115 파일럿(16장)은 **2026-07-15 QA FAIL — 계약 위반 17건**(골든 오염 9장 + 계약 이미지 2건 + 플레이스홀더 4건 + 앵커 수치 누락 2건) 판정 후 재빌드 기록 없이 **동결** 상태다 — `_workspace_kifrs/03_qa-report.md`·`04_contract-check.md`. 상세: `9_KIFRS-PILOT.md`.
- 따라서 §1~§2의 장치들은 "골든 덱 자체 검증"까지는 실측 PASS지만, **실전 덱 완주로는 아직 한 번도 증명되지 않았다.**

### 3-2. 게이트 커버리지 갭
- audit_golden SPECS는 레이아웃 15종 중 **7종만 등록**(problem_grid·screenshot·tech_capture·tech_evidence·tech_tree·mirror_matrix·boundary) — 미등록은 **8종**: 본문 계열 exec_graph·tech_mechanism·ab_simulation·validation·closing + 정형 장 cover·toc·part. `audit_golden.py` L123~200.
- golden-content-contract 계약표는 15타입 중 **10타입만** 덤프(cover/toc/part/screenshot/closing 표 없음).
- compare_golden은 폰트 **이름**을 비교하지 않는다(크기·볼드만 — 알려진 미해결 공백), `docs/user/07-factory-port.md` L4-5.
- verify_content_contract 케이스⑤ 정규식은 `c.get("x") or`·dict 리터럴 폴백을 못 잡는 탐지 갭.
- 게이트는 전부 CLI 수동 실행 — 자동 후크가 아니어서 "실행을 건너뛰면 없는 게이트"가 된다.

### 3-3. 문서 스테일
- `docs/user` **9종**(00-INDEX~08-system-overview)은 **2026-07-15 이후 미갱신** — 17장 개편·adapted/novel·audit_deck 미반영(19장·490도형 서술이 6파일에 잔존). 08 문서는 "현행(정본)" 라벨인데 최신이 아니다. **2026-07-20 형상관리에서 제외**했다 — 같은 이야기의 낡은 판본이 최신판(본 보고서)과 함께 공개되면 어느 쪽이 맞는지 알 수 없기 때문이다. 파일은 로컬에 남아 있고 본 보고서의 인용 근거로만 쓴다.
- `pptx-build/SKILL.md`도 "19장·490도형"으로 낡음(실측 17장·519도형). 본 보고서의 수치는 코드·스냅샷 실측을 정본으로 채택했다.

### 3-4. 증거 소실 — gitignore 구조
- `ref/`(레퍼런스 PNG 75장)·`data/`(McKinsey PDF 등)·`results/`(빌드 실물 4건 — 20·25·16·16장)·`_workspace/`(A/B 실측·v4 검증 원본)·`docs/user/`(구 사용자 문서 9종)·골든 렌더 PNG가 전부 gitignore — 진단·실측의 1차 증거가 클론에 없다.
- 따라서 본 보고서가 `docs/user`·핸드오프를 근거로 인용한 대목은 **저자 머신에서만 검증 가능**하다. 코드·스냅샷을 근거로 든 대목은 클론에서도 재현된다.
- **클린 클론 빌드 불가**: ref/ 소실 + `s10_screenshot.py:19`의 타 프로젝트 절대경로 하드코딩 — 골든 기준선 재생성이 저자 머신에서만 가능(gate-repair 스펙 알려진 공백).

### 3-5. 구조적 트레이드오프
- **골든 1장 = 유지보수 계약 1개** — 창고를 넓힐수록 스냅샷·오딧·계약표 유지 비용이 정비례로 는다(novel의 창고 편입을 사용자 선택으로 둔 이유).
- 폰트(Pretendard·D2Coding)는 뷰어 환경 의존 — 미설치 환경에서 대체 폰트로 렌더.
- 미커밋 워킹트리: brand-kit compact 19pt(실험값)·s06 파일럿·`_workspace_kifrs/` 등이 커밋/방치 미결 상태로 남아 있다(git status 실측).

## 4. 실증 예시(walked example) — accent 한 글자 변조가 게이트에 잡히기까지

"단일 출처 + 기계 게이트" 조합의 효과를 변조 실험 실측(2026-07-15, 당시 골든 19장·490도형 시점)으로 추적한다. 근거: `docs/superpowers/specs/2026-07-15-gate-repair-design.md` L189·L211-213.

1. **입력**: `brand-kit.yaml`의 accent `#D66E3A`(Muted Ember)를 `#00FF00`(형광 초록)으로 1개 값만 변조한다. SSOT 구조라 이 한 줄이 골든 덱 전 장의 강조색을 바꾼다.
2. **구 게이트(색 서명이 없던 시절)**: compare_golden이 490/490 "불일치 0"으로 **통과** — 도형 수·좌표·텍스트만 보고 색은 장님이었다. 형광 초록 덱이 회귀 게이트를 무사 통과하는 상태.
3. **신 게이트(fill·run_colors 서명 확장 후)**: 같은 변조에 **55건 검출** — accent가 실제로 칠해진 도형·런이 전부 불일치로 잡힌다.
4. **같은 방식의 좌표 검증**: `goldenfab/grid.py`의 MARGIN_L 0.6을 0.61로(0.01인치) 변조하면 **197건 검출** — 좌측 정렬 도형 전체가 스냅샷과 어긋나기 때문.
5. **종착**: 변조를 되돌리면 다시 불일치 0. 이 실험이 "게이트가 무엇을 보는지"를 수치로 증명했고, 색 서명 확장 시점에 숨어 있던 실제 색 불일치 14건(HANDOFF 기록)도 함께 드러났다 — 검증 로그 전체는 `11_TEST-DECISIONS.md`.

통상적 접근에는 이 실험 자체가 성립하지 않는다 — 기준(스냅샷)도, 그것을 대조하는 게이트도 없기 때문이다.
