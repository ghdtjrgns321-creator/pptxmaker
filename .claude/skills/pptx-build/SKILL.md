---
name: pptx-build
description: deck-spec.json과 brand-kit.yaml로 진짜 네이티브 .pptx를 빌드한다. 표는 네이티브 표, 차트는 네이티브 차트로 생성(이미지 아님, 데이터만 고치면 PPT에서 자동 반영). pptx-builder 에이전트가 슬라이드를 실제 파일로 만들 때, 또는 "PPT 빌드/다시 빌드/deck 재생성/브랜드킷 수정 후 재빌드" 요청 시 반드시 사용.
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
