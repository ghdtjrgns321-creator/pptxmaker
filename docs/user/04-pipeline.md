# 04. 전체 파이프라인 — 현재 상태

소스에서 최종 pptx까지 데이터가 지나는 길을 현재형으로 기록한다. 폐기된 구조는 [01-journey.md](01-journey.md)에만 있다.
개발자용 상세는 [../PIPELINE.md](../PIPELINE.md).

## 데이터 흐름 (시작 → 종착)

```
[재료]                             [Claude Code — pptmaker 오케스트레이터]

<프로젝트>/FINAL-REPORT/*.md ──▶  ① deck-outline-grill (목차 공동설계 게이트)
  (사용자가 미리 정리한               FINAL-REPORT 기반 목차 1안 → grill식 인터뷰 → "확정"
   보고서형 md 묶음.                  → _workspace/01.5_outline.md (아웃라인 계약)
   추출·NotebookLM 단계 없음)                  │
                                              ▼
                                  ② deck-composer
                                     FINAL-REPORT + 01.5_outline 계약 → B2B 셀링 골격 배치
                                     → 03_exhibit-candidates.json (시각 후보)
                                              │
                                              ▼
                                  ②.5 사용자 승인 게이트
                                     make_mockups.py → mockups/gallery.html
                                     사용자 회신 반영 → 02_deck-spec.json
                                              │
                                              ▼
                                  ③ pptx-builder
                                     build_pptx.py + brand-kit.yaml
                                     type "golden.*" → goldenfab registry(15타입) / 그 외 → 레거시
                                     → deck.pptx (네이티브 표·차트 + mpl PNG 하이브리드)
                                              │
                                              ▼
                                  ④ consistency-qa
                                     audit_pptx.py + FINAL-REPORT 근거 대조
                                     → PASS ─▶ results/<프로젝트명>-소개.pptx (종착)
                                     → FAIL ─▶ 원인 단계로 되돌림 (1회)
```

## 단계별 역할

| 단계 | 에이전트/주체        | 스킬               | 입력                           | 출력                            |
| ---- | -------------------- | ------------------ | ------------------------------ | ------------------------------- |
| ①    | 오케스트레이터(대화) | deck-outline-grill | FINAL-REPORT/\*.md             | 01.5_outline.md (아웃라인 계약) |
| ②    | deck-composer        | deck-compose       | FINAL-REPORT + 01.5_outline.md | 03_exhibit-candidates.json      |
| ②.5  | 사용자 승인 게이트   | pptx-visuals       | 후보 JSON                      | 02_deck-spec.json               |
| ③    | pptx-builder         | pptx-build         | 02_deck-spec.json              | deck.pptx                       |
| ④    | consistency-qa       | consistency-qa     | deck.pptx + 02_deck-spec.json  | 03_qa-report.md                 |

## 단일 출처 (여기만 고치면 전체 반영)

| 대상                   | 단일 출처                                                                           |
| ---------------------- | ----------------------------------------------------------------------------------- |
| 색·폰트·크기·여백      | `.claude/skills/pptx-build/assets/brand-kit.yaml`                                   |
| 차트·다이어그램 레시피 | `pptx-visuals` 스킬 (`scripts/visuals.py`)                                          |
| 슬라이드 구성 데이터   | `02_deck-spec.json` — 수정은 항상 이 파일 경유 재빌드, 완성 pptx를 직접 뜯지 않는다 |

## 원칙 (요약)

1. 재료는 사용자가 정리한 FINAL-REPORT 단일 원천 — 추출·NotebookLM 단계 없음.
2. 구성은 빌드 전 GRILL 인터뷰에서 확정(아웃라인 계약) — deck-composer는 계약을 지킨다.
3. 수정은 deck-spec 경유 재빌드 — 재현성 유지.
4. 브랜드는 brand-kit.yaml 하나만 고친다.

## 골든 덱 전환 반영 현황 (2026-07-11)

[03-golden-deck.md](03-golden-deck.md) 결정 중 ③은 반영 완료, ④와 ②.5는 Phase 4 대기:

- ③ **반영 완료** — build_pptx가 `golden.*` 타입을 goldenfab registry(15타입)로 디스패치하고,
  성공 기준은 compare_golden.py의 골든 덱 도형 전수 일치(490/490)다 → [07-factory-port.md](07-factory-port.md).
- ④ QA 게이트의 마감 결함 전담 강등, ②.5 목업 갤러리 제거·다양성 게이트의 골든 어휘 검사 교체는
  Phase 4에서 진행한다(그 전까지 본문 그림이 현행).
- ①(GRILL)·②(콘텐츠 트랙)는 유지하되, 내용 고도 상향(임원 언어 큐레이션)이 ②에 추가될 예정이다.
