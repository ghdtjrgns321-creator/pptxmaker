---
name: consistency-qa
description: 빌드된 .pptx가 브랜드 일관성(색·폰트·여백)과 셀링 품질(흐름·오탈자·정보밀도)과 네이티브 표/차트 정상 여부를 만족하는지 검증한다. 실패 항목은 앞 단계로 되돌릴 지시를 낸다. consistency-qa 에이전트가 산출물을 검수할 때, 또는 "PPT 검증/일관성 확인/품질 점검 다시" 요청 시 반드시 사용.
---

# consistency-qa — 일관성·품질 검증

"공장 일관성 + 셀링 전달력"이 지켜졌는지 확인하는 마지막 방어선. **존재 확인이 아니라 교차
검증** — spec이 의도한 것과 실제 pptx가 담은 것을 대조한다. 프로그램으로 검증 가능한 것은
스크립트로, 판단이 필요한 것은 기준에 따라 정성 평가한다.

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
- `diversity_box_ratio`: 박스 다이어그램(flow/layers/cards/branch/from_to) 본문의 30% 초과 = FAIL
- `diversity_max_repeat`: 동일 유형 덱 전체 3회 이상 = FAIL

위반 시 → **deck-composer로 되돌림**(아키타입 재배정 — archetype-catalog.md 세트 제약).

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

| 항목        | 합격 기준                           | 실패 시       |
| ----------- | ----------------------------------- | ------------- |
| 표지 명료성 | 표지만으로 "무엇을 파는지" 이해     | deck-composer |
| 흐름        | 문제→솔루션→가치 순서 유지          | deck-composer |
| 정보 밀도   | 슬라이드당 불릿 5개 이하, 한 메시지 | deck-composer |
| 근거        | 숫자에 출처(01_extracted) 대응      | deck-composer |
| 오탈자      | 명백한 맞춤법·깨진 글자 0           | pptx-builder  |

## 출력: `_workspace/03_qa-report.md`

각 항목 PASS/FAIL + FAIL이면 [되돌릴 대상 에이전트]와 구체적 수정 지시. 집합 주장은
분모를 명시한다(예: "표 3/3 네이티브 확인", "폰트 위반 0/전체 런 128").

## 원칙

- FAIL을 관대하게 넘기지 않는다. 일관성은 이 하네스의 핵심 가치다.
- 단, 근거 없이 트집잡지 않는다 — 기준표에 없는 주관적 취향은 지적하지 않는다.
- 되돌림은 1회 재시도 후에도 실패하면 리포트에 명시하고 진행(무한 루프 금지).
