# pptmaker v5 파이프라인 — FINAL-REPORT(사용자 정리) → 목차 인터뷰 → 일관 서식 공장

콘텐츠는 사용자가 프로젝트별 `FINAL-REPORT/`(보고서형 md 묶음)로 미리 정리해둔다. 이 하네스는
그 FINAL-REPORT를 목차·강조점 인터뷰(GRILL)로 셀링 골격에 앉히고 **일관된 골격·브랜드의
네이티브 .pptx로 빌드**한다. 추출·NotebookLM 단계는 없다 — FINAL-REPORT가 이미 사람이
작성·검증한 as-built 단일 재료 원천이라, 밀도·출처·검증이 확보돼 있다(같은 소스 A/B 실측:
직접 추출 재료 6,246단어 ≈ k-ifrs FINAL-REPORT 6,098단어로 동급). 서식은 변함없이
코드(`build_pptx.py` + `brand-kit.yaml`)만이 보장한다.

## 데이터 흐름 (전체)

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

## 다이어그램 DSL (deck-spec `diagram` 타입)

python-pptx에는 다이어그램 객체가 없으므로 도형+커넥터 조합으로 렌더한다. 시작 레이아웃 2종:

- `flow`: 노드 좌→우 배치 + 화살표 (프로세스 흐름)
- `layers`: 노드 상→하 적층 (아키텍처 레이어)

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
