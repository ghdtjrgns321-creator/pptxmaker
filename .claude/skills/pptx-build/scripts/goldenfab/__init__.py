"""goldenfab — 골든 덱 레이아웃의 공장 이식판 (Phase 2).

골든(golden/*.py)은 동결된 기준 원본이고, 이후 수정은 이 패키지에서만 한다(05-plan 결정 #9).
콘텐츠는 전부 인자(dict)로 받는다 — 레이아웃 좌표·서식은 design-rules.md의 확정값 그대로.
검증: scripts/compare_golden.py 가 골든 원본과 도형 전수 비교(회귀 게이트).
"""
