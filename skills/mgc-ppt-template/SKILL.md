---
name: "mgc-ppt-template"
description: "메가MGC커피向 PPT/프레젠테이션(가맹점 안내문, 사내 보고자료, 기능 중간 보고, 캠페인/이벤트 설명자료, 기능기획서/화면기획안/화면 흐름·요구사항 정의서/개발 연동 통합본·API 연동 정의서 등) 제작·수정 요청 시 적용 — 브랜드 컬러·폰트·슬라이드 구조를 표준화하고 문서 장르(안내문형 / 기능 보고형 / 간결 기획서형)로 분기. 기능 보고형은 번들 공통 코드(scripts/mgc_report.js)로 기준 개선본과 동일한 레이아웃을 재현. 메가커피·MGC 관련 발표자료·장표·슬라이드·보고 덱을 만들거나 고칠 때는 사용자가 '템플릿'을 언급하지 않아도 이 스킬을 먼저 읽는다"
---

# 메가MGC커피 PPT 제작 표준 (v5.5, 2026-10-26)

이 파일은 **읽기 지도**다. 규칙 본문은 `references/`에 주제별로 한 곳씩만 있다. 작업 유형에 맞는 파일만 읽는다.

- `.pptx` 생성 기술(PptxGenJS 사용법, hex 규칙, validate 절차)은 기본 `pptx` 스킬(`/mnt/skills/public/pptx/SKILL.md`)을 따른다. 이 스킬은 그 위에 메가MGC커피 고유의 브랜드 · 레이아웃 · 문구 규칙을 얹는다.
- 콘텐츠 확보가 끝난 뒤 적용한다.
- 실행 환경: Node + `pptxgenjs`, Python + `python-pptx` · `Pillow`.
- 스킬 폴더는 읽기 전용이다. `scripts/` 전체를 작업 폴더로 복사한 뒤 쓴다: `cp <스킬 폴더>/scripts/* /home/claude/`

## 작업별로 읽을 파일

| 지금 하는 작업 | 먼저 읽을 파일 | 함께 읽을 파일 |
|---|---|---|
| 새 덱: 안내문형 | `references/genre-notice.md` | `references/design-common.md` · `references/writing.md` · `references/text-fit.md` |
| 새 덱: 기능 보고형 | `references/genre-feature-report.md` | `references/design-common.md` · `references/writing.md` · `references/text-fit.md` |
| 새 덱: 간결 기획서형 | `references/genre-slim-spec.md` | `references/design-common.md` · `references/writing.md` · `references/text-fit.md` |
| 흐름도 장이 있는 덱 | `references/flowchart.md` | 간결 기획서형 파일 |
| 개발 연동 통합본 · API 연동 정의서 | `references/integration-doc.md` | 간결 기획서형 파일 · `references/writing.md` |
| 사용자가 다듬은 pptx 수정 · 전 장 확장 · 받은 덱 전체 정리 | `references/edit-user-pptx.md` | `references/writing.md` · `references/text-fit.md` |
| 사용자에게 넘길 최종본 · 산출물 세트 | `references/final-deliverable.md` | `references/writing.md` |
| 검수(모든 작업의 마지막) | `references/qa-checklist.md` | — |
| 이 스킬 자체를 고칠 때 | 아래 「규칙 위치 지도」 | `skill-maintainer` 스킬 |

## 장르 고르기

| 조건 | 장르 |
|---|---|
| 가맹점 · 파트너 안내, 정책 공지 | 안내문형 |
| 화면 캡처 위주로 3단계 이하 기능을 소개하는 짧은 보고(6장 안팎), 모든 장이 패턴 A~F로 그려짐 | 기능 보고형 |
| 화면이 많은 기획서 · 화면 흐름 정의서 · 요구사항 정의서 · 통합 기획서(10장 이상) | 간결 기획서형 |
| 개발 과제의 분기 · 처리 순서를 보여 주는 중간 보고(흐름도 중심) | 간결 기획서형 + 흐름도 |
| 개발사 · 연동사와 함께 보는 연동 문서(API · 처리 흐름 · 예외) | 간결 기획서형 + 연동 문서 |
| 판단이 애매함 | 간결 기획서형 |

- 기능 보고형 패턴 A~F에 없는 장(흐름도 등)이 하나라도 필요하면 기능 보고형이 아니다. 기능 보고형 덱에 손으로 그린 장을 섞지 않는다.

## 모든 장르에 항상 적용하는 규칙(요약 · 본문은 기준 위치)

| 규칙 | 기준 위치 |
|---|---|
| 맑은고딕 단일 글꼴, `LAYOUT_WIDE`, 테마 글꼴 교체 | `references/design-common.md` 「기본 설정」 |
| 장식 도형 금지, 슬라이드 밖 요소 금지 | `references/design-common.md` 「깔끔한 장표」 |
| 색은 토큰 표에 있는 값만 | `references/design-common.md` 「컬러 토큰」 |
| 줄 탈락 0건, 줄바꿈은 직접 넣은 위치에서만 | `references/text-fit.md` 「원칙」 |
| 표 · 상자는 한 정보 한 줄, 행 높이는 줄 수만큼, 마지막에 `finalize` | `references/text-fit.md` 「깔끔한 표 · 상자」 |
| 제목 「토픽: 서술」, 표 칸 · 설명 · 각주 개조식 | `references/writing.md` 「문체」 |
| 글머리 「- 」, 강조 빨강 FF0000 | `references/writing.md` 「기호」 · 「강조」 |
| 작업 이력 · 결정 과정은 덱에 쓰지 않음 | `references/final-deliverable.md` 「작업 이력 금지」 |
| 검사 스크립트 3종 통과 후 전 장 렌더 확인 | `references/qa-checklist.md` |

## 규칙이 겹칠 때 우선순위

1. 사용자의 이번 요청
2. 사용자가 직접 다듬어 보낸 pptx(기준본)의 배치 · 색 · 글자 크기
3. 문서 유형 규칙: 연동 문서(`references/integration-doc.md`), 최종본(`references/final-deliverable.md`)
4. 장르 규칙(`references/genre-*.md`, `references/flowchart.md`)
5. 공통 규칙(`references/design-common.md` · `references/writing.md` · `references/text-fit.md`)

- 같은 단계 안에서는 적용 범위가 좁은 규칙(특정 장 · 특정 표)이 넓은 규칙보다 먼저다.
- 규칙끼리 실제로 부딪히면 임의로 고르지 않는다. 사용자에게 묻고, 답을 기준 위치에 반영한다(아래 「규칙 위치 지도」).

## 규칙 위치 지도(한 주제 = 한 기준 위치)

새 지시를 반영하거나 규칙을 찾을 때 이 표로 기준 위치를 정한다. 다른 파일에는 규칙을 다시 쓰지 않고 기준 위치를 가리키기만 한다.

| 주제 | 기준 위치 |
|---|---|
| 슬라이드 크기 · 글꼴 · 장르별 여백 · 본문 하단 | `references/design-common.md` |
| 허용 도형 · 표지 모양 · 화면 밖 금지 | `references/design-common.md` |
| 색 토큰 · 색 하나의 의미 | `references/design-common.md` |
| 줄 탈락 · 문장 끝 줄바꿈 · 넘칠 때 고치는 순서 · 폭 보정 값 · 묶어 나누지 않는 말 | `references/text-fit.md` |
| 문체(서술형 · 개조식) · 주어 · 목적어 · 기호 · 글머리 | `references/writing.md` |
| 강조(빨강 · 신규 개발 배지 · 요약 띠) · 확인 태그 | `references/writing.md` |
| 명칭 · 줄인 말 · 예시값 · 기준점 · 단계 번호 | `references/writing.md` |
| 안내문형 장 순서 · 크기 | `references/genre-notice.md` |
| 기능 보고형 골격 · 패턴 A~F · 타이포 · 글자 수 한도 · 화면 이미지 준비 | `references/genre-feature-report.md` |
| 간결 기획서형 골격 · 부품 · 검증된 배치 · 장 순서 · 정렬 끝선 · 글자 크기 검증값 · 쪽 번호 | `references/genre-slim-spec.md` |
| 흐름도 논리 검수 · 흐름도 부품 · 배치 | `references/flowchart.md` |
| 연동 문서 장 구성 · 번호 원본 장 · 출처 줄 · 키(한글 이름) 사전 | `references/integration-doc.md` |
| 사용자 판 직접 편집 · 전 장 확장 · 장 이동 · 피드백 장 번호 대조 | `references/edit-user-pptx.md` |
| 버전 표기 · 작업 이력 금지 · 모순 대조 · 발표자 노트 · 보고 구조 · 산출물 세트 · docx 동기화 | `references/final-deliverable.md` |
| 검수 절차 · 체크리스트 | `references/qa-checklist.md` |
| 버전별 변경 내용 | `CHANGELOG.md` |

### 스킬을 고칠 때

- `skill-maintainer` 스킬의 「규칙 반영 절차」를 따른다: 위 지도로 기준 위치를 찾고 → 같은 주제의 기존 문장을 **고쳐 쓰고**(뒤에 덧붙이지 않음) → 검수표 · 예시 · 스크립트를 같이 맞추고 → `CHANGELOG.md`에 1~3줄 → 점검기 실행.
- 규칙 본문에 버전 꼬리표(괄호 속 버전 · 개정 표시)를 달지 않는다. 언제 바뀌었는지는 `CHANGELOG.md`에만 쓴다.
- 점검: `python3 <skill-maintainer 폴더>/scripts/skill_lint.py <이 스킬 폴더>` → 오류 0건. 기준값 · 금지어 · 주제별 기준 위치는 `lint.json`에 있다.

## 번들 파일 (`scripts/`)

| 파일 | 역할 | 언제 쓰나 |
|---|---|---|
| `mgc_report.js` | 기능 보고형 공통 부품(골격 · 패턴 함수 · 표지). 수정하지 않는다 | 기능 보고형 덱 생성 시 항상 `require` |
| `example_interim_report.js` | 기준 덱 6장(패턴 A~F) 재현 스크립트 | 기능 보고형 새 덱의 출발점. 복사해서 문구 · 이미지 경로만 교체 |
| `fit_text.js` | 줄 맞춤: 문장 끝 줄바꿈, 맑은고딕 폭 추정, 자간 자동 축소, 넘치면 생성 중단 | `mgc_report.js` · `slim.js`가 자동 사용. 직접 그리는 코드도 `fit()`을 거친다 |
| `slim.js` | 간결 기획서형 부품(골격 · phones · notes · key · stats · table · cap · divider · count · 표지) | 간결 기획서형 덱 생성 시 `require` |
| `kit.js` | `slim.js` 위 자동 줄나눔(`wrap`) · 표 행 높이 자동 계산(`tb`) · 여백이 남으면 줄 수가 늘지 않는 범위에서 글자 +2pt까지 키움 · 간지 목록 2열 12개 자동 채움 | 간결 기획서형 덱을 새로 만들 때(30장 넘는 통합본 권장) |
| `flow_parts.js` | 흐름도 부품(node · edge · legend · note) | 흐름도 장이 있는 간결 기획서형 덱 |
| `crop_screens.py` | 화면안에서 휴대폰 · B.O 영역 자르기, 잡티 덮기 | 화면 이미지 준비 |
| `check_text.py` | 글꼴(맑은고딕 통일), 줄 탈락, 줄 수, 문장 줄바꿈, 기호([기호] 「•」 · 「–」 글머리 · 「 — 」 구분 · 줄 머리 「·」, [참고] 줄 끝 「·」), 표 정리([정리] 조기 줄바꿈 · 빈 줄 공간 행), 단어 끊김 방지 설정 검사 | 모든 pptx 검수 |
| `check_clean.py` | 장식 도형 · 화면 밖 요소 검사 | 모든 pptx 검수 |
| `compare_layout.py` | 기준 덱과 좌표 · 색 · 글자 서식 비교 | 기준 덱이 있을 때 검수 |
| `pptx_edit.py` | python-pptx 직접 편집: 칸 · 글상자 문구 교체(줄나눔 · 빨강 서식 유지), 행 높이 재계산, 행 · 열 추가 · 삭제, 장 복제 · 번호 재계산, 출처 줄, 이미지 교체, 정보 도형(상자 · 화살표), 열 폭 재분배(`balance_cols`), 마무리(`finalize`: 문단 들여쓰기 · 단어 끊김 방지 · 이전 색 정리) | 사용자가 다듬은 pptx를 고칠 때, 모든 생성 · 수정의 마지막 |
| `edit_lib.py` | 글쓰기 도구: 개조식 끝맺음 변환(`gaejo`), 글머리 「- 」 · 나열 묶음 · 고유 명칭 붙임 줄바꿈(`wrap`), 이어지는 줄 문단 들여쓰기 · [[ ]] 빨강 유지 글 넣기(`set_items`) | 칸 · 상자 문구를 새로 쓸 때 |
| `format_pass.py` | 서식 일괄 적용: 빨강 FF0000 변환, 「- 」 열 판정(`include`) · 칸 앞 붙이기(`prefix_cell`), 빈 칸 가운데(`center_empty`), 요약 띠 라벨 굵게, 확인 태그 서술형(`tagfix`), 순서 번호(`number_chain`), 바뀐 표만 맞춤(`fit_tables`) | 사용자 수정 장을 전 장으로 넓힐 때 |
| `gaejo_pass.py` | 손으로 고치지 않은 표 칸 · 설명 · 도식 글을 개조식 · 글머리로 일괄 정리하고, 바뀐 칸이 있는 표만 맞춤(이미 기준에 맞는 칸은 그대로) | 장별 수정을 끝낸 뒤 마지막에 한 번 |
| `rename_pass.py` | 문단 단위 문구 치환(런 경계에 걸쳐도 서식 유지, pptx `A` · docx `W` 이름공간), 명칭 일괄 통일(`PG_RULES`), 표 행 높이를 문단 수로 다시 잡기(`tighten`) | 명칭 · 결정 문구를 모든 산출물에서 함께 바꿀 때 |
| `tidy_tables.py` | 이미 만든 덱의 표 정리: 한 줄에 들어가는 「- 」 정보 합치기, 열 폭 재분배, 행 높이 맞춤(하단이 내려가면 되돌림) | 받은 덱을 기준대로 정리할 때 |
| `move_remap.py` | 장 이동 + 쪽 번호 · 간지 목록 · 본문 장 참조 · 노트 · 목차 범위 재계산(다른 문서 장 번호 제외). 다른 영역으로 옮기면 간지 목록 항목도 옮김. 간지 자체는 옮기지 않음 | 보고 구조를 바꿔 장을 옮길 때 |
