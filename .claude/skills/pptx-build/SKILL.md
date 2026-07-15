---
name: pptx-build
description: deck-spec.json과 brand-kit.yaml로 진짜 네이티브 .pptx를 빌드한다. 표는 네이티브 표, 차트는 네이티브 차트로 생성(이미지 아님, 데이터만 고치면 PPT에서 자동 반영). pptx-builder 에이전트가 슬라이드를 실제 파일로 만들 때, 또는 "PPT 빌드/다시 빌드/deck 재생성/브랜드킷 수정 후 재빌드" 요청 시 반드시 사용. **골든 레이아웃·아키타입 코드(goldenfab/의 _variant_*·s0N_variants·kit·grid·reference)를 재설계·수정할 때, 그리고 슬라이드의 시각 구조를 바꾸는 모든 작업 — "이 장 다시 만들어 / 레이아웃 바꿔 / 시각화가 내용과 안 맞아 / 구획 나눠 / 2칸으로 / 밀도 높여 / 포인트색 넣어 / 이 슬라이드 이상해" — 에도 반드시 사용**: 디자인 단일 출처 design-rules(물성→형식, 오딧 5종, P1 물성선언, P4 셀프반려 8문항)가 이 스킬 경유로만 로드되므로, 거치지 않으면 규칙이 적용되지 않는다.
---

# pptx-build — 네이티브 PPTX 빌더

deck-spec을 실제 `.pptx`로 만든다. 핵심 원칙은 **일관성은 코드와 brand-kit에 박제**한다는 것 —
같은 spec은 항상 같은 결과를 낸다. 손으로 색·폰트를 넣지 않는다.

## 입력

1. `deck-spec.json` — `references/deck-spec-schema.md`의 계약을 따르는 슬라이드 명세
2. `assets/brand-kit.yaml` — 색·폰트·여백 단일 출처(기본 제공, 사용자가 값만 교체)

## 실행

```bash
python scripts/build_pptx.py <deck-spec.json> <out.pptx> [--brand assets/brand-kit.yaml]
```

`--brand`를 생략하면 스킬 번들의 `assets/brand-kit.yaml`을 자동으로 쓴다.

## 하는 일 / 하지 않는 일

- **한다:** 골격(cover→toc→본문→cta) 렌더, 네이티브 표/차트 생성, brand-kit 색·폰트 적용,
  목차 자동 채움(본문 섹션 제목).
- **하이브리드(v4):** `chart_type`이 확장 7종(waterfall·heatmap·dumbbell·slope·funnel·
  annotated_scatter·histogram)이거나 `render:"image"`면 pptx-visuals의 `mpl_exhibits`로
  220dpi PNG를 `<out>/render/`에 생성해 삽입한다 — 이것만이 허용된 이미지 차트 경로.
  네이티브 차트의 `emphasis`(회색+강조 1색)와 `annotations`(도형 콜아웃)도 렌더한다.
- **하지 않는다:** 기본 6종 차트·표를 이미지로 굽지 않는다. 색·폰트를 하드코딩하지 않는다
  (brand-kit에서만). spec에 없는 슬라이드 타입을 임의 발명하지 않는다.

## 빌드 후 자체 확인

빌드가 끝나면 반드시 재파싱으로 네이티브 여부를 확인한다(이미지 폴백이 아닌지):

```python
from pptx import Presentation
prs = Presentation("out.pptx")
for s in prs.slides:
    for sh in s.shapes:
        assert not (sh.shape_type and sh.shape_type == 13 and False)  # 표/차트가 그림으로 새지 않았는지
# has_table / has_chart 로 네이티브 객체 카운트 → 0이면 spec 확인
```

표가 있어야 하는데 `has_table` 카운트가 0이거나, 차트가 그림(PICTURE)으로 잡히면 실패로 보고한다.

## brand-kit 수정 시

색·폰트를 바꾸라는 요청은 코드가 아니라 `assets/brand-kit.yaml`만 수정하면 된다.
수정 후 기존 deck-spec으로 재빌드하면 전체가 새 브랜드로 일괄 반영된다.

## 스키마

슬라이드 타입·필수 필드·골격 규칙은 `references/deck-spec-schema.md` 참조.

## 골든 엔진 (Phase 2 배선 — 우선 사용)

- deck-spec 슬라이드에 `"type": "golden.<layout>"`을 쓰면 build_pptx가
  **goldenfab registry**(`scripts/goldenfab/registry.py`, 15타입)로 디스패치한다 —
  골든 장은 헤더·결론바·출처를 자체 포함하므로 레거시 푸터 스탬프가 붙지 않는다.
- 콘텐츠: **15타입 전부 `"content": {...}` dict 지원**(파라미터화 완료 — 하드코딩 0). 주면
  **텍스트만** 교체(좌표·색·도형은 코드 고정 — override 불가).
  **content는 선택이 아니라 필수다** — 골든 DEFAULT 키를 전부 덮지 않으면 빌드가 중단된다
  (`goldenfab/content_contract.py`). 안 그러면 골든 글(다른 프로젝트 내용)이 조용히 나간다.
  타입별 content 키는 `deck-compose/references/golden-content-contract.md`(코드 DEFAULT 덤프) 참조.
- 타입 배정 판정은 deck-compose의 `references/layout-matching.md` 결정표.
- **회귀 게이트**: goldenfab 수정 후 `uv run python scripts/compare_golden.py` —
  `goldenfab/reference.py`가 조립한 레퍼런스 덱 19장을 **기준선 스냅샷**
  (`assets/golden-snapshot.json`, 현재 490도형)과 도형 전수 대조. 불일치 1건이라도 FAIL,
  통과 없이 반영 금지.
  - **goldenfab이 유일 소스다.** 레이아웃을 **의도적으로** 바꿨으면
    `compare_golden.py --update-snapshot`으로 기준선을 재생성하고 **git diff로 무엇이 바뀌었는지
    리뷰**한다. 스냅샷을 손으로 편집하지 말 것.
  - 게이트가 보는 것: 도형 수·종류·좌표(±0.005")·텍스트·채움색·**런 색**·**표 셀**·폰트 pt/bold.
  - 게이트가 **못** 보는 것: **"형태가 내용에 맞나"** — 그 자리를 P1 물성선언 + P4 셀프반려가
    메운다(아래 디자인 규칙).

## 디자인 규칙 (골든 확정 — 필수)

레이아웃 GRID·오딧 5종·색 규율·타이포 위계·텍스트 리듬·도형 어휘·구성 문법은
**`references/design-rules.md`가 단일 출처**다. 슬라이드를 렌더하기 전에 반드시 읽고,
오딧을 통과하지 못한 산출물은 내보내지 않는다. 사용자 승인으로 확정된 규칙이므로
임의 변경 금지 — 변경은 골든 덱(golden/) 재승인 경유.

### 계약 필수 2항목 (P1·P4 — 산문으로 두면 잊힌다, 계약에 박아라)

레이아웃·시각을 만들거나 고치는 작업이면 **코드를 열기 전에** 완료 계약에 아래 둘을 넣는다.
design-rules 안에만 있으면 이 스킬을 안 거친 경로에서 0이 된다 — 그래서 여기 올린다.

1. **물성 선언(P1)** — 계약 **1번 항목**: `물성 선언 — <내용> = <형태>` 한 줄.
   내용의 동사(차단·분해·분기·도달)는 그 동작이 보이는 형태로, 명사는 실물 형태로.
   **회색 텍스트 카드 나열은 물성이 진짜 "목록"일 때만.** 완성 시안이 선언과 다르면 그 항목 FAIL.
2. **반려 채점 8/8(P4) — 자기채점 금지, 반드시 제3자 에이전트** (2026-07-14 실증으로 확정)
   사용자에게 **제시하기 전에** 렌더 PNG를 **별도 에이전트**(Agent 도구, 설계 맥락 미제공)에게
   넘겨 8문항 + 아래 2문항을 채점시킨다. **오케스트레이터 자가 PASS 금지** — 1건이라도 FAIL이면
   제시하지 말고 고친 뒤 재채점. 계약에 "제3자 채점 8/8 통과(에이전트 ID)" 항목 필수.
   - **왜 자기채점이 금지인가(실증):** 오케스트레이터가 8/8 PASS를 준 슬라이드를 제3자가 채점하자
     **4건 FAIL**이 나왔다 — ①텍스트 덤프, ②메커니즘 미형태화, ⑨셰브런으로 XOR을 그림,
     ⑩**같은 장에서 검정의 의미가 뒤집힘**(상단 검정=안전 / 하단 검정=치명). ⑩은 자가채점이
     끝내 못 본 결함이다. 규칙을 다 읽어도 **자기 판정을 자기가 검증하면 오판이 통과한다.**
   - **추가 2문항(필수):** ⑨ **형식-내용 적합** — 이 형태가 내용의 물성(진행/분기/대비)에 맞나.
     ⑩ **색 의미 일관** — 같은 색이 한 장 안에서 다른 뜻으로 쓰이지 않나.
   - 채점 프롬프트는 설계 의도·변명을 빼고 **보이는 것만**으로 판정시킬 것("애매하면 FAIL").

자주 걸리는 함정: ①텍스트 덤프 ②메커니즘이 형태로 안 보임 ④빈 반쪽 ⑥라벨-도형 붙음.
§5 "대비 분기는 공통 뿌리 → 엘보 → 두 경로 — **나열형 2레인 금지**", §6 "**2×2 등 축 매트릭스보다
시나리오/스토리 구도 우선**", §6 "장 내 시각 어휘 중복 금지"는 실제 반려 이력에서 나온 조항이다.
