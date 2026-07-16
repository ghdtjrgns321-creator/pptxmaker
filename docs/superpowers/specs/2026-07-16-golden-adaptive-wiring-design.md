# 골든 배선 재설계 — 변형 가능한 출발점

- 날짜: 2026-07-16
- 상태: 사용자 승인(2026-07-16)
- 범위: 배선(코드·문서) 재설계. 골든 깎기 자체는 이 배선 위에서 계속.

## 배경 — 사용자가 기각한 것

기존 배선은 물성이 매칭되면 골든 레이아웃에 **텍스트만 갈아 끼운 것**(content 스왑, 좌표·색·
도형 고정)이 나가는 구조였다. 사용자가 원하는 UX는 다르다:

> 골든덱을 기초로, 프로젝트 docs에 비슷한 게 있으면 골든의 **완성도를 참고해 그 정도의
> 시각화·밀집화**를 하고, 골든에 없는 새 스타일 장이 필요하면 골든을 기준으로 저만치
> 그 내용에 맞는 시각화·밀집화가 되었으면 좋겠다.

즉 골든의 역할은 복제 템플릿이 아니라 **완성도 기준점(캘리브레이션 앵커)**이다.

## 확정 결정 3건

1. **골든 = 변형 가능한 출발점.** 구획 문법·도형 어휘·색 규율·타이포·GRID는 상속,
   항목 수·구획 수·피치·배치는 내용에서 파생. 정형 장(표지·목차·간지·클로징)만 고정 템플릿 유지.
2. **전 장 병렬 채점.** 채점 범위는 본문 전 장 그대로 두되 장당 1개 에이전트를 동시 파견 —
   순차 2시간+가 병렬로 최장 1장 시간(~15분). FAIL 장만 수정 후 그 장만 재채점.
3. **창고 편입은 사용자 선택.** novel 장은 기본 그 덱 전용. 덱 완성 후 사용자가 고른 장만
   goldenfab 승격 + design-rules 박제 + 스냅샷 케이스 추가(골든 1장 = 유지보수 계약 1개).

## 빌드 경로 (재배선)

```
장 후보 → 물성 선언 → layout-matching 결정표
  ├─ 매칭 + 내용 구조 일치     → golden.<layout>  (content 전량 교체 — content_contract)
  ├─ 매칭 + 카디널리티·구획 상이 → adapted.<layout> (골든 코드·렌더 Read 후 장 스크립트 변형,
  │                                                goldenfab 부품 kit·grid 재사용)
  └─ 매칭 없음                 → novel            (가장 가까운 골든 렌더 2~3장을 밀도 앵커로
                                                   Read, P1 물성 선언부터 신규 설계)
```

구현은 파라미터 격자가 아니라 **장 스크립트**(`build(prs, content)`) — 골든 깎기에서 이미 쓰는
variants 스크립트 방식을 실전 덱으로 확장한 것. `build_pptx._render_scripted`가 디스패치한다.

## 품질 보증 — 스냅샷의 빈자리 3겹

스냅샷 회귀(compare_golden)는 **goldenfab 보호 전용**으로 물러난다. 실전 변형·신규 장은:

| 겹         | 무엇                                                                                         | 어디                           |
| ---------- | -------------------------------------------------------------------------------------------- | ------------------------------ |
| 기계(2초)  | 전역 규칙(generic_checks) + 밀도 밴드                                                        | `scripts/audit_deck.py`        |
| 밀도 밴드  | 골든 스냅샷 본문 11장의 최저 도형·텍스트 프레임 수에서 파생(리터럴 0) — 골든보다 성기면 FAIL | `goldenfab/audit.density_band` |
| 채점(병렬) | P4 10문항 + 골든 앵커 PNG 대비 완성도, FAIL 장만 재채점                                      | design-rules P4·P5             |

장별 예외는 spec `"audit": {"dup_allow": N, "known": {...}}`로 선언(코드 수정 없이 실측을 박음).
adapted 장의 골든 글 유출은 check_contract의 골든오염 검사가 `adapted.<key>`까지 커버.

## 변경 파일

- 코드: `build_pptx.py`(_render_scripted + 디스패치), `goldenfab/audit.py`(density_band ·
  check_density · generic_checks), `audit_deck.py`(신설), `audit_golden.py`(selftest 밀도 2건),
  `check_contract.py`(layout_key — adapted 접두)
- 문서: `design-rules.md`(골든셋 원칙 개정 · P4 병렬 채점 · P5 신설), `layout-matching.md`
  (판정 절차 3갈래), `deck-compose/SKILL.md`, `pptx-build/SKILL.md`, `deck-spec-schema.md`
  (골든 계열 타입 표)

## 검증

- `audit_golden.py --selftest` PASS(밀도 밴드: 희소 검출 + 골든 오탐 0 포함)
- `compare_golden.py` PASS — goldenfab 무변(490도형 동일)
- 배선 검증 러너(계약 VERIFY): adapted·novel 빌드, script 누락 중단, 밴드 = 스냅샷 파생값,
  골든 본문 밀도 11/11, audit_deck exit 0/1 양 케이스, 유출 검출, 문서 5/5
