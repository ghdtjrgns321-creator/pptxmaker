# 전체 파이프라인 — 현재 상태 (v4 기준)

소스에서 최종 pptx까지 데이터가 지나는 길을 현재형으로 기록한다. 폐기된 구조는 [01-journey.md](01-journey.md)에만 있다.
개발자용 상세는 [../PIPELINE.md](../PIPELINE.md).

## 데이터 흐름 (시작 → 종착)

```
[재료]                        [Claude Code — pptmaker 오케스트레이터]
로컬 프로젝트 소스(기본) ──▶ ① content-extractor
NotebookLM 보고서(보조)      │   보고서형 8요소로 직접 추출 + 수치·고유명사 검증
                             │   → _workspace/01_extracted.md (+charts/*.json, images/)
                             ▼
                          ①.5 목차 공동설계 게이트 (deck-outline-grill, 신설 예정)
                             │   초안 목차 제안 → grill식 인터뷰 → 사용자 "확정"
                             │   → _workspace/01.5_outline.md (아웃라인 계약)
                             ▼
                          ② deck-composer
                             │   통합·중복제거·B2B 셀링 골격 배치
                             │   → 03_exhibit-candidates.json (시각 후보)
                             ▼
                          ②.5 사용자 승인 게이트
                             │   make_mockups.py → mockups/gallery.html
                             │   사용자 회신 반영 → 02_deck-spec.json
                             ▼
                          ③ pptx-builder
                             │   build_pptx.py + brand-kit.yaml
                             │   → deck.pptx (네이티브 표·차트 + mpl PNG 하이브리드)
                             ▼
                          ④ consistency-qa
                                audit_pptx.py → PASS ─▶ results/<프로젝트명>-소개.pptx (종착)
                                             └ FAIL ─▶ 원인 단계로 되돌림 (1회)
```

## 단계별 역할

| 단계 | 에이전트             | 스킬                           | 입력                              | 출력                            |
| ---- | -------------------- | ------------------------------ | --------------------------------- | ------------------------------- |
| ①    | content-extractor    | content-extract                | 로컬 소스(기본) · input/(보조)    | 01_extracted.md, charts, images |
| ①.5  | 목차 공동설계 게이트 | deck-outline-grill (신설 예정) | 01_extracted.md                   | 01.5_outline.md (아웃라인 계약) |
| ②    | deck-composer        | deck-compose                   | 01_extracted.md + 01.5_outline.md | 03_exhibit-candidates.json      |
| ②.5  | 사용자 승인 게이트   | pptx-visuals                   | 후보 JSON                         | 02_deck-spec.json               |
| ③    | pptx-builder         | pptx-build                     | 02_deck-spec.json                 | deck.pptx                       |
| ④    | consistency-qa       | consistency-qa                 | deck.pptx + 02_deck-spec.json     | 03_qa-report.md                 |

## 단일 출처 (여기만 고치면 전체 반영)

| 대상                   | 단일 출처                                                                           |
| ---------------------- | ----------------------------------------------------------------------------------- |
| 색·폰트·크기·여백      | `.claude/skills/pptx-build/assets/brand-kit.yaml`                                   |
| 차트·다이어그램 레시피 | `pptx-visuals` 스킬 (`scripts/visuals.py`)                                          |
| 슬라이드 구성 데이터   | `02_deck-spec.json` — 수정은 항상 이 파일 경유 재빌드, 완성 pptx를 직접 뜯지 않는다 |

## 원칙 (요약)

1. 재료는 Claude 직접 추출이 기본, NotebookLM은 보조(보고서형 텍스트만, 고유명사 검증 필수).
2. 차트는 데이터 추출 → 네이티브 재생성. NotebookLM pptx는 이미지 판독이 기본 경로.
3. 수정은 deck-spec 경유 재빌드 — 재현성 유지.
4. 브랜드는 brand-kit.yaml 하나만 고친다.

## 골든 덱 전환 후 달라지는 것

[03-golden-deck.md](03-golden-deck.md) 결정이 반영되면 이 그림에서 바뀌는 곳은 ③·④다:

- ③ build_pptx의 성공 기준이 "게이트 통과" → "골든 덱 8장 재현"으로 바뀐다.
- ④ QA 게이트는 마감 결함(각주 겹침·잘림 등) 담당으로 강등되고, 시각 품질 판정은 골든 덱과의 일치로 대체된다.
- ①·② (콘텐츠 트랙)는 그대로 유지하되, 내용 고도 상향(임원 언어 큐레이션)이 ②에 추가된다.
