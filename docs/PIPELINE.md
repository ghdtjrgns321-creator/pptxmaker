# pptmaker v2 파이프라인 — NotebookLM 재료 → 일관 서식 공장

콘텐츠 생산(리서치·초안)은 NotebookLM이, 서식·템플릿 일관성은 이 프로젝트(Claude Code)가
맡는다. 역할 분리의 근거: LLM 슬라이드 생성은 좌표·폰트·색을 보장하지 못하므로, 서식은
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

| 단계 | 에이전트          | 스킬            | 입력                     | 출력                            |
| ---- | ----------------- | --------------- | ------------------------ | ------------------------------- |
| ①    | content-extractor | content-extract | input/ 의 pptx 3~5개     | 01_extracted.md, charts, images |
| ②    | deck-composer     | deck-compose    | 01_extracted.md          | 02_deck-spec.json               |
| ③    | pptx-builder      | pptx-build      | 02_deck-spec.json        | deck.pptx                       |
| ④    | consistency-qa    | consistency-qa  | deck.pptx + 02_deck-spec | 03_qa-report.md                 |

## 원칙 (박제)

1. **골격·서식은 100% Claude Code 쪽.** NotebookLM 프롬프트에는 콘텐츠 요구만 담는다
   (헤드라인+거버닝 메시지, 마침표 금지, 수치 중심). 서식 지시는 재빌드 시 버려지므로 넣지 않는다.
   스토리 골격(문제→솔루션→증거)은 프롬프트에 **힌트**로만 주고, 강제는 ② deck-composer가 한다.
   프롬프트 템플릿: `.claude/skills/pptmaker/assets/notebooklm-prompt.md`
2. **차트는 데이터 추출 → 네이티브 재생성(A안).** 추출 3단계 폴백:
   네이티브 차트 수치 직접 추출 → 도형·텍스트에서 수치 복원 → 실패 시 이미지 크롭 + "폴백" 표기.
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
