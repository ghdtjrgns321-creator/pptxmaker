# 골든 타입 content 계약 (코드 DEFAULT 자동 덤프)

`goldenfab` 각 타입 함수의 DEFAULT dict를 그대로 덤프한 것. deck-spec의 `content`에
여기 키를 넣으면 텍스트가 교체된다(키 생략 시 골든 기본값). **좌표·색·도형은 override 불가**.
재생성: `uv run python _workspace/dump_contract.py`.

## `golden.problem_grid`  (_variant_k.variant_k)
content 키 (14개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `kicker` | str | '1. 문제 정의' |
| `headline` | str | '일반 LLM을 회계에 못 쓰는 이유는 틀려서가 아니라, 틀리는 방향 때문이다' |
| `left_eyebrow` | str | '핵심 리스크' |
| `left_big` | str | '2종 오류' |
| `left_big2` | str | '= 허위 확정' |
| `left_desc` | str | '근거 없는 확신은 부실감사와 재무제표 왜곡이라는 실제 손해로 이어진다.' |
| `left_nums` | list | [('01', '검증 불가', '그럴듯한 답을 즉시 내놓지만, 근거 문단을 확인할 방법이 없어 감사 조서에 쓸 수 없다.... |
| `right1_head` | str | '같은 질문, 두 개의 결말' |
| `question` | str | '"이 계약, 수익을\n지금 인식해도\n됩니까?"' |
| `lanes` | list | [{'name': '일반 LLM', 'steps': [('즉시 답변', '근거 조문 없음'), ('확신형 어조', '판단... |
| `right2_head` | str | '오류의 두 방향 — 왜 2종만 치명적인가' |
| `panels` | list | [{'head': '1종 · 놓침 — 안전한 실패', 'desc': '근거가 있는데 못 찾아 "모른다"고 답하는 실패는 ... |
| `bar` | str | '확실할 때만 확정하고, 애매하면 근거를 보여주며 유보한다' |
| `source` | str | '출처: 1_OVERVIEW.md · PROJECT_OVERVIEW.md · 5_INTERFACE.md (00_facts... |

## `golden.exec_graph`  (s06_variants.variant_c)
content 키 (14개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `kicker` | str | '2. 파이프라인' |
| `headline` | str | '질문에서 응답까지 — 분기까지 전부 결정적인 실행 그래프' |
| `subtitle` | str | '진입부 임베딩 유사도 0 — LLM이 고르는 것은 35토픽 목록뿐, 경로는 그래프가 결정한다.' |
| `node_names` | list | ['질문', 'Analyze', '판정', 'Retrieve', 'Generate', 'Format', '응답'] |
| `node_tags` | list | ['용어사전 매칭', '그래프 1홉 탐색', '판단트리 주입', '경고·꼬리질문'] |
| `in_label` | str | 'IN' |
| `reject_box` | str | '거절 메시지' |
| `out_label` | str | 'OUT' |
| `reject_desc` | str | '범위 밖 질문은 본선에 진입하지 못하고 즉시 거절된다' |
| `detail_head` | str | '핵심 기술 — 결정성을 어떻게 만들었나' |
| `details` | list | [('01  용어사전 — 후보 진입점', '등재 423 · AI 신규 창작 0.', ' 사람이 만든 자료 3종(질의 매핑... |
| `rerank_note` | str | '초기 설계에 있던 유사도 재정렬(rerank) 단계는 제거 — 그래프가 이미 결정적으로 선별하므로 재정렬할 것이 없다' |
| `bar` | str | '질문에서 답까지 4개 노드 전부가 결정적으로 동작한다' |
| `source` | str | '출처: 4_SEARCH-PIPELINE.md (00_factsheet.md §C)' |

## `golden.tech_evidence`  (s08_variants.variant_c)
content 키 (10개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `headline` | str | '용어사전 — 실무 언어를 기준서 개념에 잇는 진입 색인' |
| `kicker` | str | '3. 기술 설명 — TECH 01 · Analyze' |
| `narratives` | list | [('왜 필요한가', "실무는 '리베이트'라 말하고 기준서는 '고객에게 지급할 대가'라 쓴다. 이 언어 간극을 잇지 않으... |
| `table_rows` | list | [('용어', '원천', '등급', '연결 개념 (어디로 진입하나)'), ('리베이트', '질의 매핑', '자동', '고... |
| `table_caption` | str | '등재 423 중 발췌 5건 — 자동 316 · 위임판단 86 · 검토 18 · 확정 1 · 제외 2' |
| `json_title` | str | 'aliases.json — 실제 엔트리' |
| `json_lines` | list | ['{ "term": "상품권",', '  "sources": ["query-mapping"],', '  "grade":... |
| `json_caption` | str | '모든 엔트리가 결정 로그(누가·왜)를 갖는다 — 전건 추적.' |
| `bar` | str | 'AI가 만드는 것은 색인 하나 — 틀려도 1종(놓침)으로 드러나는 자리에만 둔다' |
| `source` | str | '출처: 2_DATA-TAXONOMY.md §2.6 · 4_SEARCH-PIPELINE.md (00_factsheet.m... |

## `golden.tech_tree`  (s09_variants.variant_a)
content 키 (17개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `kicker` | str | '3. 기술 설명 — TECH 02 · Retrieve' |
| `headline` | str | '지식그래프 — 기준서의 구조를 그대로 옮긴 결정적 지도' |
| `narratives` | list | [('왜 필요한가', '문단 뭉치에는 순서도 관계도 없다. 기준서는 계층·상호참조·사례가 얽힌 구조 — 그 구조를 보존해... |
| `struct_head` | str | '구조 — 기준서의 위계 그대로' |
| `root` | str | '기준서 1115' |
| `concepts` | list | ['변동대가', '보증', '⋯ 80'] |
| `paras` | list | ['문단 50', '문단 56', 'B33'] |
| `layer_tag1` | str | '개념 80' |
| `layer_tag2` | str | '문단 250' |
| `term_chip` | str | '“볼륨디스카운트”' |
| `term_cap` | str | '용어 423 — 진입' |
| `case_chip` | str | '사례 188' |
| `xref_cap` | str | '사례는 인용 문단으로 연결 · 문단끼리는 상호참조(E3 244)' |
| `table_title` | tuple | ('무엇을 어떻게 만들었나', '    그래프 v14 — 노드 929 · 간선 2,697 · 고립 0') |
| `build_rows` | list | [('구성 요소', '어떻게 만들었나', '수'), ('개념', '기준서 공식 소제목을 그대로 노드로 — 이름도 경계도 ... |
| `bar` | str | "임베딩이 놓치는 '법적 이웃'을 관계가 잡는다 — 텍스트가 아니라 구조로 검색한다" |
| `source` | str | '출처: 2_DATA-TAXONOMY.md · 3_KNOWLEDGE-GRAPH.md (00_factsheet.md §C·... |

## `golden.tech_mechanism`  (s11_variants.variant_d)
content 키 (19개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `kicker` | str | '3. 기술 설명 — TECH 03 · Generate' |
| `headline` | str | '판단트리 41개 — 흩어진 조건-분기를 판단 순서로 미리 조립' |
| `narratives` | list | [('왜 필요한가', '하나의 회계 판단이 문단 32·35·36·37과 B9~B13에 흩어져 있다. 문단을 조각으로 검색... |
| `steps_head` | str | '트리는 이렇게 걸린다 — 주제기반 다중주입' |
| `steps` | list | [('주제 지목', '질문 분석 AI가 35개 주제 목록에서 최대 3개를 지목한다 — 개념 80개를 직접 고르지 않는다.... |
| `branch_label` | str | '주입 뒤 — 답변은 질문 성격에 따라 세 갈래로 갈린다' |
| `flow_head` | str | '적용 실물 — 질문에서 트리 분기까지' |
| `question` | str | '“볼륨디스카운트 조항이 있는 계약은…”' |
| `topic_label` | str | '주제 지목 — 변동대가' |
| `tree_chip` | str | '「변동대가 (추정)」 걸림' |
| `diamond` | str | '변동금액 포함?' |
| `diamond_anchor` | str | '문단 51' |
| `method1_head` | str | '기댓값' |
| `method1_sub` | str | '확률 가중 합' |
| `method2_head` | str | '가능성 최고 금액' |
| `method2_sub` | str | '단일 결과치' |
| `flow_foot` | str | '문단 53 — 더 잘 예측하는 방법 하나를 일관 적용(문단 54)' |
| `bar` | str | '질문이 여러 판단에 걸치면 걸린 절차를 모두 넣는다 — 트리 41개, 전부 원문 앵커' |
| `source` | str | '출처: data/ontology/judgment_trees.json(트리 41 · 사용자 검수 완료) · 3_KNOWL... |

## `golden.tech_capture`  (s12_variants.variant_b)
content 키 (14개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `kicker` | str | '3. 기술 설명 — TECH 04 · Generate' |
| `headline` | str | '구조화 출력 — 답변의 형식을 코드가 강제한다' |
| `narratives` | list | [('왜 필요한가', '같은 질문에 매번 다른 형식으로 답하면 정량 평가가 불가능하다. 자유형 텍스트를 폐기하고 답변의 ... |
| `mid_subhead` | str | '실물 답변 화면 — 발췌' |
| `caption` | str | '전체 답변 중 [확인 질문]·[조건부 결론 Case 1·2] 발췌 — 본인 vs 대리인 질의' |
| `schema_subhead` | str | '출력 스키마 — ClarifyOutput' |
| `glosses` | list | [('selected_branches', '고른 결론 분기'), ('answer', '답변 본문 (마크다운)'), ('c... |
| `validate_subhead` | str | '형식 검증 — 거부와 자동 재시도' |
| `flow_generate` | str | '답변 생성' |
| `diamond` | str | '스키마\n검사' |
| `flow_display` | str | '화면 표시' |
| `reject_note` | str | '거부 — 자동 재시도: 인용 0개 · 분기 0개 · 결론 없는 단정' |
| `bar` | str | '형식을 지키지 못한 답변은 화면에 도달하지 못한다 — 스키마 선언 + 검증기 자동 재시도' |
| `source` | str | '출처: app/agents.py (ClarifyOutput · output_validator) · 5_INTERFACE... |

## `golden.ab_simulation`  (s14_variants.variant_c)
content 키 (15개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `headline` | str | '같은 질문, 두 시스템 — 점수는 떨어뜨리고, 경로는 도달한다' |
| `kicker` | str | '4. 문제와 해결' |
| `left_title` | str | '리랭커 — 점수가 근거다, 그런데 점수엔 근거가 없다' |
| `question` | str | '“볼륨디스카운트 조항이 있는 계약은…”' |
| `rank_labels` | list | ['1위', '2위', '3위'] |
| `surface_card` | str | '표면 단어가\n비슷한 문단' |
| `cut_label` | str | '상위 컷 — 아래는 버려진다' |
| `reject_mark` | str | '✕' |
| `reject_card` | str | '전문가 큐레이션 문단 — 정답인데 표면 유사도가 낮다' |
| `left_note` | str | '실측 — 전문가 배정 큐레이션 문서가 표면 유사도가 낮아 105건 중 103건 탈락(98%). 이미 bypass로 우회... |
| `right_title` | str | '지식그래프 — 경로가 근거다' |
| `steps` | list | [('용어사전 매칭 — 볼륨디스카운트 → 변동대가', '사람이 전수 검수한 색인 (등재 423)'), ('개념 노드 — ... |
| `right_note` | str | '탈락시킬 점수가 없다 — 연결이 있으면 도달하고, 어떤 간선을 지났는지가 그대로 답변의 근거가 된다.' |
| `bar` | str | '확률 신호 전량 폐기 — 임베딩·가중치·리랭커 없이 온톨로지 그래프로 동작한다' |
| `source` | str | "출처: 7_JOURNEY.md §7.3·§7.4 (재구축 여정) · 00_factsheet.md §A'·§D" |

## `golden.validation`  (s15_variants.variant_c)
content 키 (20개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `kicker` | str | '4. 문제와 해결' |
| `headline` | str | '골든테스트 92건 — 정답 문서를 차단한 채 사람의 결론을 재현하는가' |
| `col1_head` | str | '① 어떻게 만든 시험인가' |
| `doc_caption` | str | '실제 질의회신 92건 — 질의자·회계기준원·해석위원회 작성' |
| `doc_q_head` | str | '질문' |
| `doc_q_desc` | str | '실무 사실관계' |
| `doc_r_head` | str | '회신' |
| `doc_r_desc` | str | '전문가의 결론' |
| `dests` | list | ['질문만 시스템에 입력', '회신은 채점 정답지로\n3축 대조 — 결론·문단·분기', '문서 자체는 검색 차단 ✕\n격... |
| `col1_note` | str | '자작 시나리오를 버리고 사람이 쓴 실제 문답으로 — AI가 만든 문제를 AI가 푸는 자가순환과, 자기 답을 자기가 인용... |
| `col2_head` | str | '② 최종 기록' |
| `final_big` | str | '78 / 92' |
| `final_label` | str | '사람 전문가 결론 재현' |
| `stats` | list | [('소프트 채점', '92 / 92'), ('실행 에러', '0건'), ('개선 추적', '재현 72 → 78'), (... |
| `col2_note` | str | '하드 인용 재현율 59.1%는 별도 보고 — 인용률만으로 품질을 재지 않는다.' |
| `col3_head` | str | '③ 미재현 14건 — 각각 왜' |
| `buckets` | list | [('6', '기타 소프트축', '타 기준서 4건(1115 코퍼스 밖 — 인용 불가) + 헤지 2건(근거 확보, 결론 미... |
| `col3_note` | str | '합 6+4+2+2 = 14, 전수 귀속 — 검색 재설계로 실제 움직일 수 있는 것은 1건뿐이다.' |
| `bar` | str | '정답이 적힌 문서를 못 보게 한 상태에서 78/92 — 실패 14건까지 전수 해부된 수치다' |
| `source` | str | '출처: 6_TEST-DECISIONS.md (00_factsheet.md §F·§H) — 사람 작성 골든테스트 92건 ... |

## `golden.mirror_matrix`  (s17_variants.variant_c)
content 키 (7개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `kicker` | str | '5. 차별점' |
| `headline` | str | '일반 임베딩 RAG과 무엇이 다른가 — 7개 축 전부 같은 방향' |
| `left_head` | str | '일반 임베딩 RAG' |
| `right_head` | str | '이 시스템' |
| `mirror` | list | [('검색 진입', '임베딩 유사도\n점수 경쟁', '용어사전 문자 매칭\n결정적 진입'), ('근거 설명', '유사도 ... |
| `bar` | str | '일반 RAG의 확률 신호 자리마다 기준서의 구조가 들어가 있다 — 7축 전부, 예외 없이' |
| `source` | str | '출처: 00_factsheet.md §G (6_TEST-DECISIONS.md·4_SEARCH-PIPELINE.md)' |

## `golden.boundary`  (s18_variants.variant_b)
content 키 (14개):

| 키 | 타입 | 골든 기본값(샘플) |
| -- | ---- | ----------------- |
| `headline` | str | '정직한 한계 — 이 시스템이 서 있는 경계' |
| `outside_label` | str | '경계 밖' |
| `inside_title` | str | 'K-IFRS 1115 — 경계 안' |
| `inside_policy_head` | str | '답하는 범위를 계약한다' |
| `inside_policy_body` | str | '경계 안 질문에만 결정적으로 답하고, 밖이면 거절, 모르면 유보 — 아는 척이 구조적으로 불가능하게 만든 설계다.' |
| `out_left_chip_head` | str | '타 기준서 질문' |
| `out_left_chip_sub` | str | 'IAS38 · 1002 · 1008 · 1037' |
| `out_left_desc` | str | "라우팅이 코드 레벨 강제 OUT — 추측 대신 거절. 실측 4건 (예: 1116호 '증분차입이자율' 감지)" |
| `out_right_chip_head` | str | '등재되지 않은 표현' |
| `out_right_chip_sub` | str | '임베딩식 유사 확장 없음' |
| `out_right_desc` | str | "색인에 없으면 진입 실패 → '못 찾음' 유보 응답. 홀드아웃 실측 진입 누락 2건" |
| `bottom_head` | str | '경계 그 밖의 한계 — 실측으로 아는 것' |
| `limits` | list | [('결론을 확정하지 못한 케이스', '헤지 2건 — 근거 문단을 다 찾고도 결론을 유보했다. 검색이 아니라 생성 계층의... |
| `bar` | str | '경계를 넓히는 대신 경계 안을 결정적으로 — 못 하는 것까지 실측으로 세어 두었다' |

