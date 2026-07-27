---
name: consistency-qa
description: 빌드된 .pptx가 브랜드 일관성(색·폰트·여백)과 셀링 품질(흐름·오탈자·정보밀도)과 네이티브 표/차트 정상 여부를 만족하는지 검증한다. 실패 항목은 앞 단계로 되돌릴 지시를 낸다. consistency-qa 에이전트가 산출물을 검수할 때, 또는 "PPT 검증/일관성 확인/품질 점검 다시" 요청 시 반드시 사용.
---

# consistency-qa — 일관성·품질 검증

"공장 일관성 + 셀링 전달력"이 지켜졌는지 확인하는 마지막 방어선. **존재 확인이 아니라 교차
검증** — spec이 의도한 것과 실제 pptx가 담은 것을 대조한다. 프로그램으로 검증 가능한 것은
스크립트로, 판단이 필요한 것은 기준에 따라 정성 평가한다.

## 0. 계약 대조 — **정답은 계약이다** (스텝 1, 필수)

> 왜 이게 1번인가(2026-07-15 실증): 계약에 8장 전부 "README §번호"와 실제 이미지 4종이
> **확정**으로 적혔는데 반영 0장인 덱이 이 QA를 **PASS**로 통과했다. 근거로 "본문=골든
> 기본값(00_factsheet 실측, 창작 0)"이라 적었다. 골든은 이 프로젝트로 만들어졌으니 골든 글은
> 팩트시트에 **대응한다** — 즉 옛 §4 기준("FINAL-REPORT에 대응")으로는 골든 글이 정답이 된다.
> 게이트가 주문서를 안 읽고 자기들끼리 일관성만 봤다. 그래서 기준을 계약으로 바꾼다.

가장 먼저, 다른 어떤 검사보다 앞서 실행한다:

```bash
uv run python .claude/skills/consistency-qa/scripts/check_contract.py \
    _workspace/01.5_outline.md _workspace/02_deck-spec.json _workspace/deck.pptx
```

검사: 장 수·순서 · 제목 · **골든 기본값 오염** · 계약 확정 이미지 실존(md5) · 템플릿
플레이스홀더 · 앵커 수치. 출력 `_workspace/04_contract-check.md`, 종료코드 1 = FAIL.

**FAIL이면 다른 항목이 전부 PASS여도 리포트 판정은 FAIL이다.** 되돌림 대상은 리포트 지시대로.

### 이 절의 금지 사항 (전부 실제로 저지른 것)

- **확정 격하 금지** — 계약에 "확정"인 항목을 "사용자 선택 대기"·"원하면 추가 가능"으로
  적지 않는다. 사용자가 이미 지시했고 계약에 채록됐다. 미이행은 미결이 아니라 **FAIL**이다.
- **골든 기본값을 정답으로 쓰지 않는다** — "골든과 같아서 PASS"는 근거가 아니다. 골든은 서식
  견본이지 콘텐츠 원천이 아니다.
- **이미지 0을 미덕으로 쓰지 않는다** — "전 슬라이드 네이티브 도형, PICTURE 0"은 계약이
  이미지를 요구하지 **않은** 장에서만 칭찬거리다. 계약이 요구했는데 0이면 FAIL이다.
- **스테일 수치 인용 금지** — 게이트 수치는 **이번 실행 결과만** 인용한다. 기억·이전 리포트에서
  베끼지 않는다(실제 사고: 재실행값 490을 "501"로 적었고, 그때 그 게이트는 색을 0개 보고 있었다).

## 1. 네이티브 객체 검증 (스크립트)

`scripts/audit_pptx.py`로 pptx를 재파싱해 확인한다:
- deck-spec의 `table` 슬라이드 수 == pptx의 `has_table` 수 (이미지로 새지 않았는지)
- deck-spec의 **네이티브** `chart` 수 == pptx의 `has_chart` 수 — 확장 7종(waterfall 등)·
  `render:"image"`는 mpl PNG가 정상 경로이므로 Picture로 계수(image_exhibits 검사)
- `diagram` 슬라이드는 도형(자동도형) 렌더가 정상 — 그림(PICTURE)으로 새지 않았는지
- 슬라이드 수 == spec의 slides 길이

불일치 시 → **pptx-builder로 되돌림**.

## 1-2. 다양성 게이트 4종 (v4.2 — audit_pptx.py가 spec 기준 기계 검증)

- `diversity_cooldown`: 동일 시각 유형(chart:종류/diagram:레이아웃/table…) 간격 3장 미만 = FAIL
- `diversity_min_kinds`: 본문 5장 이상인데 유형 5종 미만 = FAIL
- (`diversity_box_ratio`는 2026-07-25 삭제 — 대상 어휘인 박스+화살표 9종이 폐기돼 계수가 항상 0이었다.
  빈 박스 위험은 (A) 골든 쪽으로 옮겨가 `goldenfab.audit.check_adhoc_card`(§8)가 진다)
- `diversity_max_repeat`: 동일 유형 덱 전체 3회 이상 = FAIL

위반 시 → **deck-composer로 되돌림**(아키타입 재배정 — archetype-catalog.md 세트 제약).

## 1-3. 골든·변형·신규 장 전역 오딧 (audit_deck.py — 밀도·재탕·경계·accent)

spec에 `golden.*`(정형 제외)·`adapted.*`·`novel` 장이 하나라도 있으면 **반드시** 실행한다.
이 장들은 §1-2(레거시 다양성 게이트)의 대상이 아니라 goldenfab 전역 규칙으로 채점된다:

```bash
uv run python .claude/skills/pptx-build/scripts/audit_deck.py \
    _workspace/02_deck-spec.json _workspace/deck.pptx
```

검사(장별): **§6-D 밀도 밴드(글자수 주지표** — 골든 본문 최저 텍스트량 파생, 스크린샷 장은 예외)
· §2 accent 상한 · P4③ 노드 재탕 · P4④ 채움률 · P4⑩ 노드 클래스 · §F 그림 침범 · P2 경계.
종료코드 1 = FAIL. **밀도 FAIL = "허전·탑헤비" 결함이 정량으로 잡힌 것** — 도형 수가 아니라
글자수로 재므로 "카드 4장으로 도형 채우기"로는 통과 못 한다(그 반사가 재작업 루프의 원인이었다).

- **밀도 미달** → **deck-composer/장 스크립트로 되돌림**: 카드로 메우지 말고 그 장 물성의 근거
  텍스트를 보강한다(layout-matching §검색 플라이휠·design-rules §8 R0). 재드로·4카드 강제 금지.
- 단일 dense 모듈 개발 중 사전 점검은 `preflight_dense.py <mod>`(같은 audit.py 지표, green 후 제시).

## 2. 브랜드 일관성 (스크립트 + 대조)

- 텍스트 런의 폰트가 brand-kit의 head/body/mono 집합에만 속하는지 (외부 폰트 혼입 0)
- 사용된 색이 brand-kit 팔레트 근방인지 (임의 색 혼입 검사)
- 골격 준수: 첫 두 장 cover→toc, 마지막 장 cta

불일치 시 → **pptx-builder 또는 brand-kit 확인**.

## 3. 렌더 눈검증 (필수 — 기계 검사가 못 잡는 시각 결함)

audit PASS 후 반드시 전 장을 PNG로 렌더해 **직접 본다**(겹침·잘림·여백 과다는 파싱으로 안 잡힘):

```powershell
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open((Resolve-Path "_workspace\deck.pptx").Path, $true, $false, $false)
New-Item -ItemType Directory -Force _workspace\render | Out-Null
$pres.SaveAs((Resolve-Path "_workspace\render").Path + "\slide", 18); $pres.Close(); $ppt.Quit()
```

- PNG 수 == 슬라이드 수 확인(전수), 최소 표본: 간지 1 + 차트 1 + 표 1 + 다이어그램 각 레이아웃 1 + 최다 텍스트 슬라이드 1을 Read로 판독
- 체크: 텍스트 겹침/잘림 0, 개체가 본문 영역 밖으로 나가지 않음, 슬라이드 하단 40% 공백이면 밀도 결함
- 발견 결함은 [되돌릴 대상]과 함께 리포트에 좌표·슬라이드 번호로 특정
- 렌더 후 산출물 복사가 "다른 프로세스가 사용 중"으로 잠기면 유령 POWERPNT 프로세스가
  원인 — `Get-Process POWERPNT`로 확인 후 종료하고 재시도(COM Quit이 즉시 안 풀릴 수 있음)

## 4. 셀링 품질 (정성 기준)

| 항목        | 합격 기준                                                      | 실패 시       |
| ----------- | -------------------------------------------------------------- | ------------- |
| 표지 명료성 | 표지만으로 "무엇을 파는지" 이해                                | deck-composer |
| 흐름        | **계약이 확정한 목차 순서** 유지(골격 표준은 계약이 없을 때만) | deck-composer |
| 정보 밀도   | 슬라이드당 불릿 5개 이하, 한 메시지                            | deck-composer |
| 근거        | 숫자·고유명사가 **계약이 그 장에 지정한 출처**에 대응(창작 0)  | deck-composer |
| 오탈자      | 명백한 맞춤법·깨진 글자 0                                      | pptx-builder  |

> 근거 기준 주의: "FINAL-REPORT 어딘가에 있음"은 합격 근거가 **아니다**. 계약이 S8에 "README
> §3"을 지정했으면 S8의 근거는 README §3이어야 한다. 팩트시트에 있다는 이유로 통과시키면
> 골든 기본값이 전부 통과한다(2026-07-15 사고의 원인).

## 출력: `_workspace/03_qa-report.md`

각 항목 PASS/FAIL + FAIL이면 [되돌릴 대상 에이전트]와 구체적 수정 지시. 집합 주장은
분모를 명시한다(예: "표 3/3 네이티브 확인", "폰트 위반 0/전체 런 128").

## 원칙

- **정답은 계약이다.** 골든·팩트시트·이전 리포트가 아니라 사용자가 확정한 계약(`01.5_outline.md`)이
  기준. 계약과 대조하지 않은 PASS는 PASS가 아니다.
- FAIL을 관대하게 넘기지 않는다. 일관성은 이 하네스의 핵심 가치다.
- 단, 근거 없이 트집잡지 않는다 — 기준표에 없는 주관적 취향은 지적하지 않는다.
- 되돌림은 1회 재시도 후에도 실패하면 리포트에 명시하고 진행(무한 루프 금지).
