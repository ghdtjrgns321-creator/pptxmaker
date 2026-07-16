# pptmaker

내가 만든 프로젝트를 소개하는 B2B 셀링 PPT를 일관 서식으로 찍어내는 서식 공장.

## 하네스: PPT 서식 공장 (v5)

**재료(v5):** 콘텐츠는 사용자가 프로젝트별 `FINAL-REPORT/`(보고서형 md 묶음)로 미리
정리해둔다 — 추출·NotebookLM 단계 없음. 이 하네스가 그 FINAL-REPORT를 목차 인터뷰로 셀링
골격에 앉히고 **일관된 골격·브랜드의 네이티브 .pptx로 빌드**한다. 전체 설계·원칙은
`docs/PIPELINE.md`.

**트리거:** "프로젝트 PPT 만들어줘 / 소개 자료 생성 / 발표자료 / pptx 생성 / 이 프로젝트로 덱 만들어" 및 후속
"다시 만들어 / 슬라이드 수정 / 브랜드 바꿔 재빌드 / 스토리 다시" 요청 시 `pptmaker` 오케스트레이터
스킬을 사용하라. 단순 질문은 직접 응답 가능.

**구성:** ① deck-outline-grill(목차 인터뷰 — 오케스트레이터가 메인 대화에서 직접 수행) → ②
deck-composer → ③ pptx-builder → ④ consistency-qa. 브랜드 일관성 단일 출처는
`.claude/skills/pptx-build/assets/brand-kit.yaml`(이 파일만 고치면 전체 반영). 차트·다이어그램
레시피 단일 출처는 `pptx-visuals` 스킬. **수정은 항상 deck-spec/brand-kit 경유 재빌드** —
완성 pptx를 직접 뜯지 않는다.

## 환경·의존성 (uv)

- 패키지·가상환경은 **uv**로 관리. `.venv`·`uv.lock`으로 재현 가능. Python 3.13 고정(`.python-version`).
- 스크립트 실행은 `uv run python <경로>` — venv 자동 활성, 별도 activate 불필요.
- 의존성 추가: `uv add <pkg>`(런타임) / `uv add --dev <pkg>`(개발). 런타임=`python-pptx`·`pyyaml`, 개발=`ruff`.
- `.venv`·`*.pptx`(빌드 산출물)·`.claude/state`는 커밋하지 않음(`.gitignore`).

## git 운영

- 브랜치 흐름 `feature→develop→main`. **main 직접 커밋 금지**, 기본 작업 브랜치는 `develop`.
- 커밋은 사용자 요청 시만. Conventional Commits, 1커밋=1논리변경, **AI/Claude 서명 금지**.

## 변경 이력

| 날짜       | 변경 내용                  | 대상                                                                                                                                               | 사유                                                                                                                                                                                                                                                                                                                                                                                  |
| ---------- | -------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 2026-07-06 | 초기 구성                  | 전체                                                                                                                                               | 프로젝트 소개용 B2B 셀링 PPT 자동생성 하네스 신규 구축                                                                                                                                                                                                                                                                                                                                |
| 2026-07-06 | 공장 골격 확정             | selling-curation, deck-spec-schema                                                                                                                 | 표준 골격을 "문제→솔루션→증거 8슬라이드"(B2B 표준)로 박제, 앞뒤 고정·본문 유연                                                                                                                                                                                                                                                                                                        |
| 2026-07-06 | 브랜드킷 흑백+포인트 확정  | pptx-build/assets/brand-kit.yaml                                                                                                                   | 무채색 베이스(primary #15171B) + 포인트 Muted Ember(accent #D66E3A) 단일 포인트 컨셉                                                                                                                                                                                                                                                                                                  |
| 2026-07-06 | 마스터 틀 구현             | pptx-build/scripts/build_pptx.py                                                                                                                   | 컨설팅 표준 틀(헤더 구분선·푸터 회사명·우하단 페이지번호·출처선) 전 장 자동 스탬프                                                                                                                                                                                                                                                                                                    |
| 2026-07-06 | AI티 제거·문서형 개편      | build_pptx.py, deck-spec-schema, selling-curation                                                                                                  | BCG 레퍼런스 기준: 세로 rail→헤더 이중룰, hero카드→인라인 **강조**, 문서형 푸터(중앙 워드마크·세로 저작권)                                                                                                                                                                                                                                                                            |
| 2026-07-06 | 실측 반영(글자·밀도·구성)  | brand-kit, build_pptx, reference-metrics, selling-curation                                                                                         | NYCHA BCG 111슬라이드 실측: 제목24·본문13·캡션9pt, 부제·중첩불릿(•/–)·각주 지원, 밀도 120~200단어                                                                                                                                                                                                                                                                                     |
| 2026-07-06 | git·uv 관리 전환           | pyproject.toml, uv.lock, .gitignore, CLAUDE.md                                                                                                     | 재현 가능 환경(uv venv·lock)과 버전관리(main/develop) 도입, Python 3.13 고정                                                                                                                                                                                                                                                                                                          |
| 2026-07-08 | v2 파이프라인 전환         | docs/PIPELINE.md, content-extract·pptx-visuals 신설, deck-compose 개편, project-analysis 폐기                                                      | 콘텐츠 생산을 NotebookLM으로 이관(초안 3~5개 통합), 이 프로젝트는 서식·템플릿 일관성에 집중. diagram 타입(flow/layers) 추가                                                                                                                                                                                                                                                           |
| 2026-07-08 | v3 슬라이드 문법·시각 어휘 | build_pptx.py, visuals.py, audit_pptx.py, visual-selection.md, 스킬 4종, PIPELINE.md                                                               | PART 탭·간지·계층번호(우석진 실측), 시각 13종(cards·branch·timeline·from_to·panels·banner 등, Dallas/BCG 실측), Lucide 아이콘 72종, intro·commentary 문서형 밀도, QA 게이트 8종. 재료는 Claude 직접 추출 메인 + NotebookLM 보조(A/B 실측 근거)                                                                                                                                        |
| 2026-07-08 | v4 시각 다양성 파이프라인  | mpl_exhibits.py·make_mockups.py 신설, visuals.py·build_pptx.py·audit_pptx.py, 스킬 5종, ref/                                                       | 어휘 3종 수렴 문제 해결: ②.5 사용자 승인 게이트(목업 갤러리), mpl 하이브리드 7종(워터폴·히트맵·덤벨·슬로프·퍼널·주석산점도·히스토그램), 다양성 게이트 3종(연속 금지·최소 5종·박스 30%), 회색+accent 1색 스타일 박제, Dallas 레퍼런스 앵커. 롤백 태그 v3-pre-visual-diversity                                                                                                          |
| 2026-07-12 | v5 전반부 완전제거         | content-extract·content-extractor·notebooklm-prompt 삭제, pptmaker·deck-outline-grill·deck-compose·PIPELINE·docs/user 재배선                       | 사용자가 프로젝트별 `FINAL-REPORT/`(보고서형 md)를 이미 유지 → 추출·NotebookLM 전반부가 "보고서에서 보고서를 다시 뽑는" 중복이라 제거. GRILL이 ① 첫 단계, 재료 원천을 FINAL-REPORT로 단일화. deck-compose의 초안 통합·중복제거 기계는 단일 소스라 삭제, grep 수치검증은 QA의 FINAL-REPORT 대조로 흡수                                                                                 |
| 2026-07-16 | 골든 = 변형 가능한 출발점  | build_pptx(_render_scripted)·goldenfab/audit(밀도 밴드)·audit_deck 신설·check_contract, design-rules(P5)·layout-matching·스킬 2종·deck-spec-schema | 사용자가 "텍스트만 갈아 끼운 것" 기각 — 매칭 장도 내용이 형태를 정하게 adapted(장 스크립트 변형)·novel(골든 렌더 밀도 앵커) 경로 신설. 스냅샷 회귀는 goldenfab 보호 전용으로 축소, 실전 장은 audit_deck(전역 오딧+스냅샷 파생 밀도 밴드)+전 장 병렬 채점(FAIL 장만 재채점). novel의 창고 편입은 사용자 선택. 스펙: docs/superpowers/specs/2026-07-16-golden-adaptive-wiring-design.md |
