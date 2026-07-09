# pptmaker v3 파이프라인 — Claude 직접 추출(메인) + NotebookLM(보조) → 일관 서식 공장

재료 추출부터 서식까지 Claude Code가 메인으로 수행하고, NotebookLM은 보조(관점 변주·
비로컬 소스)로 쓴다. 2026-07-08 같은 소스 A/B 실측: **직접 추출이 밀도 1.3배·출처 인용
16배 우위 + NotebookLM 재료에서 고유명사 환각 2건 검출**(_workspace/08_material-comparison.md).
재료 형식은 **보고서형 8요소**(슬라이드 요약형 금지 — 재료가 마름). 서식은 변함없이
코드(`build_pptx.py` + `brand-kit.yaml`)만이 보장한다.

## 데이터 흐름 (전체)

```
[사용자]                         [Claude Code — pptmaker 오케스트레이터]
NotebookLM에                ┌──────────────────────────────────────────────────┐
프롬프트 템플릿 투입          │ ① content-extractor                              │
  │                         │    pptx 3~5개 파싱 → 01_extracted.md              │
  ▼                         │    + charts/*.json + images/                      │
pptx 초안 3~5개 다운로드      │            ▼                                     │
  │                         │ ② deck-composer                                   │
  ▼                         │    통합·중복제거·골격 배치 → 02_deck-spec.json     │
input/<프로젝트명>/ 에 저장 ──▶│            ▼                                     │
                            │ ③ pptx-builder                                    │
                            │    build_pptx.py + brand-kit → deck.pptx          │
                            │            ▼                                      │
                            │ ④ consistency-qa                                  │
                            │    audit_pptx.py → PASS / FAIL(되돌림 1회)         │
                            └──────────────────────────────────────────────────┘
                                          ▼
                            results/<프로젝트명>-소개.pptx (최종 산출물)
```

## 단계별 역할

| 단계 | 에이전트           | 스킬            | 입력                         | 출력                                            |
| ---- | ------------------ | --------------- | ---------------------------- | ----------------------------------------------- |
| ①    | content-extractor  | content-extract | 로컬 소스(기본)·input/(보조) | 01_extracted.md, charts, images                 |
| ②    | deck-composer      | deck-compose    | 01_extracted.md              | 03_exhibit-candidates.json(시각 후보)           |
| ②.5  | 사용자 승인 게이트 | pptx-visuals    | 후보 JSON → make_mockups.py  | mockups/gallery.html → 회신 → 02_deck-spec.json |
| ③    | pptx-builder       | pptx-build      | 02_deck-spec.json            | deck.pptx (하이브리드: 네이티브+mpl PNG)        |
| ④    | consistency-qa     | consistency-qa  | deck.pptx + 02_deck-spec     | 03_qa-report.md (+다양성 게이트 3종)            |

## 원칙 (박제)

1. **재료는 직접 추출이 기본.** 로컬 소스가 있으면 content-extract 모드 A(보고서형 8요소
   직접 작성 + 수치·고유명사 grep 검증)로 뽑는다. NotebookLM은 모드 B(보조) — 보고서형
   텍스트만 받고(pptx 금지), 수입 재료는 고유명사까지 검증한다.
   프롬프트 템플릿: `docs/notebooklm-prompt.md` (보조 경로용)
2. **차트는 데이터 추출 → 네이티브 재생성(A안).** 단, NotebookLM pptx는 **슬라이드당
   래스터(PNG) 1장**으로 구워져 나온다(2026-07-08 실측: 15/15 슬라이드 텍스트 상자 0개) —
   본문·수치 추출은 에이전트의 **이미지 판독**(멀티모달 Read)이 기본 경로다. 폴백:
   네이티브 차트 수치 직접 추출 → 이미지 판독 복원 → 실패 시 이미지 크롭 + "폴백" 표기.
   텍스트 보고서(md)를 함께 받으면 그것을 주 재료로, pptx 이미지는 보조로 쓴다.
3. **수정은 항상 deck-spec 경유 재빌드.** 완성 pptx를 직접 뜯지 않는다 — deck-spec이 단일
   출처여야 "브랜드 바꿔 재빌드" 재현성이 유지된다. 차트·다이어그램 레시피는 `pptx-visuals`
   스킬이 단일 출처(`scripts/visuals.py`).
4. **브랜드 단일 출처는 `brand-kit.yaml`.** 색·폰트·크기·여백은 이 파일만 고친다.

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
- 샘플 NotebookLM pptx 미확보 상태로 설계 — 첫 실전 투입 때 추출 로직(특히 차트 형태)을 보정한다.
- v4 (2026-07-08): 시각 다양성 파이프라인. 첫 실전 덱이 "막대 1·표 3·박스플로우" 3어휘로
  수렴한 문제를 4방법론 조합으로 해결 — ①익스히빗 사양서 선행 패스(②.5 사용자 승인 게이트,
  make_mockups.py 갤러리) ②확장 매핑·다양성 강제 3조항(visual-selection C-2, audit 게이트)
  ③mpl 하이브리드 7종(mpl_exhibits.py — 네이티브 편집성은 기본 6종 유지, 확장만 이미지)
  ④레퍼런스 앵커링(ref/catalog.md, Dallas 6페이지). 스타일 지렛대 박제: 전부 회색 + 강조만
  accent 1색 + 콜아웃 + 출처 각주. 롤백 지점: git tag `v3-pre-visual-diversity`.
