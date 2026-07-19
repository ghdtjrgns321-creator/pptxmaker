# 핸드오프 — 골든 덱 고밀도 개편 (2026-07-18 세션 종료)

> 다음 세션 첫 프롬프트에 이 파일을 지목할 것. 이 문서가 유일한 인수인계 소스다.

## 0. 목표 한 줄

**골든 덱을 기반으로**, 레퍼런스 장표(`data/KakaoTalk_20260718_160537227.png`)처럼
섹션을 나눠 밀도를 높인 스타일을 덧붙인다. 브랜드는 골든 그대로(흑백 primary
`#15171B` + Muted Ember accent `#D66E3A`) — **레퍼런스의 네이비를 베끼는 게 아니다.**

## 1. 확정된 방향 (재논의 금지)

1. 골든 = 출발점. 전면 재작성 아님 — 골든의 이상 포인트(§2)를 레퍼런스 문법으로 교정.
2. 재료 절차: 장마다 FINAL-REPORT(`C:/Users/ghdtj/workspace/portfolio/k-ifrs-1115/FINAL-REPORT/`)
   재채굴 → **명사구 재고 파일 선행**(실례: `golden/_pilot_s06/units.md`). 문장 단위 재료 금지.
3. 스토리 골격(5부 17장 순서·주장)과 검증된 수치 대장은 유지.
4. 골든의 관계 도해 어휘(팬인·판정 다이아몬드·경계 다이어그램·간선 매핑·체인)는
   **보존 자산** — 섹션 골격이 밀도를 잡고, 존 하나는 도해가 주인공으로 차지한다.
5. 사용자 시각 취향(메모리에도 있음): 숫자 나열 기각, 비율·관계 인코딩 + 고밀도.

## 2. 골든 덱의 이상 포인트 8 — 그대로 체크리스트로 박제할 것

| #   | 결함                                    | 교정 방향                                                        |
| --- | --------------------------------------- | ---------------------------------------------------------------- |
| 1   | 헤더가 너무 크다 (대학생 티)            | 키커+28pt 헤드라인+구분선 의식 → 압축 (파일럿은 19pt 한 줄 사용) |
| 2   | 맨 밑 한 줄 결론 바 — 아마추어적        | 재서술 바 폐지 → 그 지면을 콘텐츠 존(스탯 밴드·사례 스트립)으로  |
| 3   | 아이콘 전무                             | Lucide 24종×3색 보유(`pptx-build/assets/icons/`) — 행·칩에 배선  |
| 4   | **재료 알갱이가 문장 단위 (근본원인)**  | 2~6단어 명사구 40~60단위/장으로 재분해 — 밀도는 재료 속성        |
| 5   | 리드 문단(14pt)이 상단 20% 공범         | 헤더와 함께 상단 의식 통째 압축                                  |
| 6   | 요소에 경계 없음 (여백·룰만의 BCG 문법) | 모든 요소를 박스·카드·칩에 — 빈 공간이 구조적으로 못 생기게      |
| 7   | 동형 반복 구조 부재                     | 동일 내부 해부의 카드 N개 반복 (배지→제목→태그→리스트→칩)        |
| 8   | 타이포 대비 3:1 문서형 (28 vs 12pt)     | 전 텍스트 8~14pt 좁은 대역 — 정보가 면을 덮는 인상               |

## 3. 작업 절차 (지난 세션 실패에서 도출 — 순서 엄수)

1. **해부 체크리스트 먼저**: 레퍼런스 이미지를 요소 단위로 해부해 기계 검사 가능한
   수치로 박제 — 존별 단위 수·아이콘 수·박스 수·타이포 대역·위계 규칙(어두운 바 개수 등).
   이 저장소 원칙 그대로: *러너에 안 물리면 없는 규칙*.
2. **시안 수렴은 빠른 매체에서**: HTML 목업 등으로 만족할 때까지 반복 (pptx 좌표
   깎기로 시안 탐색 금지 — 한 사이클에 흠 2개 고치는 속도라 수렴 불가).
3. **pptx 이식은 1회**: 승인된 목업을 python-pptx로 옮기고 체크리스트 오딧으로 판정.
4. 한 장(파일럿) 통과 후에 문법 승격(design-rules.md + 오딧 러너 등록) → 나머지 장 확산.

## 4. 하지 말 것 (이번 세션에서 실측된 실패 패턴)

- **완성 정의 없이 최신 불만 하나씩 고치는 직렬 깎기** — 밀도↔초점 시소만 탄다.
  (v1 밀도 좋음→도해 요청, v2 도해 넣다 초점 상실, v3 초점 잡다 아이콘 0개로 밀도 후퇴)
- 껍데기(3존 프레임)부터 그리기 — proto_f가 그렇게 죽었다. 재고량이 레이아웃을 결정한다.
- 골든 도해를 좁은 레일에 미니어처로 구겨 넣기 — 팬인은 수렴선 길이가, 플로우는
  노드 크기가 본체다. 도해엔 제 비례의 지면을 줄 것.
- 한 장에 주인공 둘 (v2에서 경계 다이어그램+플로우 경쟁 → 시선 분산).
- 네이비 리터럴 사용 — brand-kit 토큰만.

## 5. 파일 지도

| 경로                                                         | 내용                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| `data/KakaoTalk_20260718_160537227.png`                      | **레퍼런스 장표** (해부 대상)                                |
| `golden/render_v6/s01~s17.png`                               | 현행 골든 렌더 (교정 대상)                                   |
| `golden/_pilot_s06/units.md`                                 | s06 명사구 재고 76단위 + § 출처 앵커 + v3 변경 기록          |
| `golden/_pilot_s06/pilot.pptx` · `render/s01.png`            | 파일럿 v3 산출물 (v1 밀도 최고·v3 위계 최고, 합본 없음)      |
| `.claude/skills/pptx-build/scripts/goldenfab/s06_pilot.py`   | 파일럿 코드 — 존 바·동형 카드·팬인·체인·다이아몬드 구현 참고 |
| `.claude/skills/pptx-build/scripts/goldenfab/s06_proto_f.py` | 실패 표본 (껍데기 먼저 → 속빈강정)                           |
| `.claude/skills/pptx-build/scripts/render_deck.ps1`          | pptx → PNG 렌더 (PowerPoint COM)                             |
| `.claude/skills/pptx-build/assets/brand-kit.yaml`            | 색·폰트·크기 단일 출처 (`compact: 19` 이번에 추가됨)         |
| `.claude/skills/pptx-build/assets/icons/`                    | Lucide 24종 × primary/accent/white                           |
| `.claude/skills/pptx-build/references/design-rules.md`       | 시각 규칙 단일 출처 — 승격 시 여기에 추가                    |

## 6. 수치 주의

- 파일럿의 수치는 전부 `units.md`에 FINAL-REPORT § 앵커로 추적 가능.
- 예외: 경계 다이어그램 수치(강제 OUT 4건·진입 누락 2건·IAS 목록)는 골든 경계 장에서
  승계한 것 — 원문 장(1_OVERVIEW·11_COVERAGE 추정) 대조 전까지 `미검증` 취급.

## 7. 빌드·렌더 원라이너

```powershell
uv run python -c "
import sys; sys.path.insert(0, '.claude/skills/pptx-build/scripts')
from goldenfab.reference import new_presentation
from goldenfab.s06_pilot import build
prs = new_presentation(); build(prs); prs.save('golden/_pilot_s06/pilot.pptx')"
pwsh -File .claude/skills/pptx-build/scripts/render_deck.ps1 -Pptx golden/_pilot_s06/pilot.pptx -OutDir golden/_pilot_s06/render
```
