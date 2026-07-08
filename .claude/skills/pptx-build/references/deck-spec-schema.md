# deck-spec.json 스키마 (에이전트 간 데이터 계약)

`deck-composer`가 산출하고 `pptx-build`가 소비하는 단일 계약. 이 스키마를 벗어난 키는
빌더가 무시하거나 실패한다. 슬라이드 **골격**(앞뒤 고정, 본문 유연)은 여기서 강제된다.

## 최상위 구조

```json
{
  "meta": { "project": "프로젝트명", "tagline": "한 줄 소개", "audience": "대상(선택)" },
  "slides": [ { "type": "...", ... }, ... ]
}
```

## 골격 규칙 (오케스트레이터가 검증)

- **맨 앞 고정:** `cover` → `toc` 로 시작
- **맨 뒤 고정:** `cta` 로 종료
- **본문(유연):** 그 사이는 프로젝트 성격에 맞게 `bullets/two_column/table/chart/diagram/metrics/section`을 자유 배치
- `toc`는 items를 비워두면 본문 섹션 제목으로 자동 채워진다 → 목차/본문 순서가 항상 일치

## 슬라이드 타입별 필드

| type         | 필수 필드                                                                    | 용도                           |
| ------------ | ---------------------------------------------------------------------------- | ------------------------------ |
| `cover`      | `title`, `subtitle`                                                          | 표지                           |
| `toc`        | (없음) `items` 선택                                                          | 목차. 비우면 본문 제목 자동    |
| `section`    | `title`                                                                      | 섹션 구분(전환)                |
| `bullets`    | `title`, `bullets[]`                                                         | 불릿 3~5개. Pain/특징 등       |
| `two_column` | `title`, `left{heading,items[]}`, `right{heading,items[]}`                   | 문제/솔루션·Before/After 대비  |
| `metrics`    | `title`, `items[{value,label}]`                                              | 큰 숫자 KPI 최대 4개           |
| `table`      | `title`, `headers[]`, `rows[[]]`                                             | 네이티브 표(기술스택·비교표)   |
| `chart`      | `title`, `chart_type`(bar/line/pie), `categories[]`, `series{name:[values]}` | 네이티브 차트(성장·비중)       |
| `diagram`    | `title`, `layout`(flow/layers), `nodes[{label,sub?}]`                        | 도형 다이어그램(흐름·아키텍처) |
| `cta`        | `title`, `contact{email,site,...}`                                           | 마무리·연락처                  |

### `diagram` 상세 (pptx-visuals 스킬이 렌더)

- `layout: "flow"` — 노드 좌→우 배치 + accent 화살표. 프로세스·파이프라인용. 노드 3~5개 권장.
- `layout: "layers"` — 노드 상→하 적층. 아키텍처 레이어용. 노드 2~5개 권장.
- 노드: `{"label": "굵은 제목", "sub": "설명 한 줄(선택)"}`. `sub`는 생략 가능.

```json
{"type":"diagram","title":"처리 흐름","layout":"flow",
 "nodes":[{"label":"수집","sub":"pptx 파싱"},{"label":"통합"},{"label":"빌드","sub":"brand-kit 적용"}]}
```

## 공장 표준 골격 (박제 — 문제→솔루션→증거)

확정된 본문 배치. 앞뒤 고정(cover→toc … cta), 본문 5단은 이 순서가 기본:

문제(bullets) → 솔루션(two_column) → 핵심기능(bullets/table) → 차별점(table/bullets)
→ 성과(metrics/chart). 순서 자체가 설득 장치다.

성과 숫자가 없으면 성과 슬라이드는 정성 강점으로 대체하거나 뺀다(문제→솔루션 순서는 불변).

## 밀도·구성 (NYCHA BCG 실측 반영 — references/reference-metrics.md)

기관 문서 밀도를 목표로 한다(슬라이드당 120~200 단어). 지원 필드:

- `subtitle`: 제목 아래 설명 한 줄(15pt). 대부분의 본문 슬라이드 권장.
- **중첩 불릿**: `bullets`/`two_column` items를 문자열 대신 객체로 — `{"text":"리드 문장","sub":["근거1","근거2"]}`.
  최상위 `•` + 하위 `–`(들여쓰기). 그룹 블록으로 밀도를 만든다.
- `footnotes`: `["단서·참조", ...]` → 하단에 7~8pt 각주.
- `source`: 정량 슬라이드 좌하단 출처.

```json
{"type":"bullets","title":"문제","subtitle":"수기 프로세스가 만드는 비용",
 "bullets":[
   {"text":"주문 접수가 **분산**돼 있다","sub":["이메일·전화 **40%**","폼·팩스 잔존"]},
   "오입력 반품 **월 12건**"],
 "footnotes":["2026-Q2 내부 집계 기준"]}
```

## 인라인 강조 (핵심 데이터 — BCG식)

`bullets`와 `two_column` items 안에서 `**...**`로 감싼 부분은 **볼드 + accent색**으로 렌더된다.
핵심 수치·차별 키워드만 감싼다(문장 전체 금지 — 강조가 강조가 아니게 됨). 실제 기관 덱은
숫자를 본문 안에 인라인으로 강조하지, 큰 색 숫자 카드를 쓰지 않는다.

```json
{"type":"bullets","title":"문제","bullets":["전체 주문의 **40%**가 수기 입력","반품 **월 12건**"]}
```

## 마스터 틀 (자동 스탬프 — 컨설팅 표준)

모든 본문 슬라이드에 코드가 동일 위치로 자동 삽입한다(spec에 안 적어도 됨):
- **헤더:** 제목 아래 얇은 구분선
- **푸터:** 좌 `회사명 · 프로젝트명`, 우하단 페이지번호 `NN / MM`, 위에 헤어라인
- **표지(cover)·마무리(cta)는 페이지번호·푸터 생략** (브랜드 마크만)

선택 필드 `"source"`: 표/차트/metrics 등 정량 슬라이드에 넣으면 좌하단에 `출처: …`로 렌더된다.
정량 주장에는 출처를 붙이는 것이 셀링 신뢰의 기본이다.

```json
{"type":"chart","title":"성과","source":"내부 로그 2026-Q2", "...":"..."}
```

## 주의

- 표·차트는 **네이티브 객체**로 생성된다(이미지 아님). 값은 spec의 숫자를 그대로 넣어라 —
  이미지로 굽지 말 것. 데이터만 고치면 PPT에서 자동 반영되는 게 핵심 가치다.
- `bullets`는 슬라이드당 5개 이하. 넘치면 슬라이드를 나눠라(정보 밀도 과다는 셀링 역효과).
