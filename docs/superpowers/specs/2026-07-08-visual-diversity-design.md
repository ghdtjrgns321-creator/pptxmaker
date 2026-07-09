# v4 시각 다양성 파이프라인 설계 (2026-07-08 승인)

## 문제

첫 실전 덱의 시각 어휘가 3종(가로막대 1·표 3·박스플로우)으로 수렴. 원인 3개:
"다양하게"는 스펙이 아님(모델은 최빈값으로 수렴) · 도구 천장(python-pptx 네이티브 차트는
막대·선·원 계열뿐) · 원패스 생성(설계 판단과 제작을 한 번에 시켜 디폴트 수렴).

## 해법 — 4방법론 조합 (사용자 결정 4건 반영)

1. **익스히빗 사양서 선행 패스 + 사용자 승인 게이트(②.5)** — deck-composer가 슬라이드별
   시각 후보 2~3안(서로 다른 유형)을 `03_exhibit-candidates.json`으로 내고,
   `make_mockups.py`가 실데이터 목업 갤러리(HTML)를 생성. 사용자가 "s07=B" 회신해야
   deck-spec 확정. 각 후보에 편집 가능성 라벨(네이티브=수치편집가능 / 이미지=수치편집불가) 명시.
2. **매핑표·강제 조항** — visual-selection.md A-2(확장 유형 7종 매핑) + C-2(연속 2장 금지·
   최소 5종·박스 30% 이하). consistency-qa의 audit_pptx.py 다양성 게이트 3종이 기계 강제,
   위반 시 deck-composer로 되돌림.
3. **하이브리드 렌더** — 기본 6종(bar/hbar/line/pie/doughnut/stacked_bar)은 네이티브 유지,
   확장 7종(waterfall/heatmap/dumbbell/slope/funnel/annotated_scatter/histogram)은
   `mpl_exhibits.py`가 220dpi PNG 생성(빌더 `_is_image_exhibit` 자동 판별, `render:"image"`
   강제 가능). 네이티브에도 `emphasis`(회색+강조 1색)·`annotations`(도형 콜아웃, 편집가능) 지원.
4. **레퍼런스 앵커링** — Dallas PDF 6페이지를 `ref/*.png`로 자산화(`ref/catalog.md`),
   deck-spec `ref` 필드로 지정하면 갤러리에 표시.

**스타일 지렛대 박제:** 전부 회색(C7CBD1) + 강조 데이터만 accent 1색 + 콜아웃 주석 + 하단
출처 각주 — mpl은 `_fig()→_save()` 단일 경로, 네이티브는 `_apply_series_colors`가 강제.

## 산출물

- 신규: `pptx-visuals/scripts/mpl_exhibits.py`, `make_mockups.py`, `ref/`(PNG 6+catalog)
- 수정: `visuals.py`(emphasis·add_callouts), `build_pptx.py`(하이브리드), `audit_pptx.py`
  (다양성 게이트 3종·이미지 계수 분리), 스킬 문서 5종, deck-spec-schema.md(필드 4종),
  visual-selection.md, PIPELINE.md
- 의존성: matplotlib(런타임), pypdfium2(dev)
- 롤백: git tag `v3-pre-visual-diversity` (c244316)

## 검증 (완료)

스모크 7/7 렌더 · 회귀 빌드 25/25 · 하이브리드 Picture 1/1 · emphasis XML accent 1/1·회색 3/3 ·
게이트 위반/통과 7케이스 전건 일치 · 갤러리 라벨 grep 4/2건. 대조표: `_workspace/v4/`.
