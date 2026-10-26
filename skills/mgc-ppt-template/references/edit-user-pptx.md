# 사용자가 다듬은 pptx 고치기

사용자가 직접 다듬은 pptx를 받으면 **그 파일이 기준본**이다. 다시 생성하지 않고 `scripts/pptx_edit.py`로 그 파일을 직접 고친다. 사용자의 배치 · 색 · 글자 크기는 바꾸지 않는다.

- 사용자가 덱을 스크립트로 다시 만들어 달라고 할 때만 역반영한다: 기준본의 텍스트 · 좌표 · 글자 크기를 추출해 생성 스크립트 결과와 비교하고(바뀐 도형만 출력), 사용자가 고친 문구 · 배치 · 크기를 스크립트에 그대로 옮긴 뒤 다음 수정을 얹는다.

## 피드백 장 번호 대조

- 피드백의 "p.N" · "N장"은 **번호만 믿지 말고 그 장의 제목 · 내용과 대조**해 확인한 뒤 고친다. 표지 포함 · 제외 번호가 섞이면 한 장씩 어긋난다(표지 제외 번호 때문에 "12p"를 잘못 해석해 다른 장을 지운 사례, "71p"가 72장(결정 장)이었던 사례).
- 잘못 고친 것이 확인되면 되묻지 않고 바로 원상 복구한다.

## 칸 문구 교체

- `set_cell`은 칸 폭 − 0.24in에서 의미 단위로 줄을 나누고, 빨강이 아닌 첫 런 서식을 원본으로 쓴다(빨강 런을 원본으로 잡아 문장 전체가 빨갛게 되는 사고 방지).
- 한 줄에 들어갈 정보가 나뉘면 `balance_cols` → `rewrap_table`로 열 폭을 다시 나눈다(`references/text-fit.md` 「깔끔한 표 · 상자」).
- 줄 수가 바뀌면(늘거나 줄거나) `fit_rows`로 행 높이를 다시 계산하고, 표 아래 요소(핵심 한 줄 · 각주 · 도식)와 겹치는지 `bottom()`으로 확인한다.
- 겹치면 `references/text-fit.md` 「넘칠 때 고치는 순서」의 사용자 판 행을 따른다.

## 열 · 행 · 장 바꾸기

- **열 · 행**: 열을 빼거나 더할 때는 `tblGrid`의 `gridCol`과 모든 행의 `tc`를 같이 바꾼다(`add_col` · `del_col`). 열만 줄이고 `gridCol`을 남기면 표가 화면 밖으로 나간다(`check_clean` "화면 밖").
- **장 추가 · 분할**: `duplicate_slide`로 원본 장을 복제해 바로 뒤에 넣고 `renumber(p, 새 장 번호)`로 쪽 번호(n / N), 간지 목록 번호, 본문 "N장" · "N~M장" · "(N~M)" 참조를 다시 맞춘다. 목차 표 범위는 간지 위치로 다시 계산하고, 새 장의 간지 목록 항목은 2열 배치로 다시 놓는다.
- **장 이동**: 보고 구조를 바꿔 장을 옮기면 `move_remap.py 입력 출력 옮길장 새위치`로 번호를 다시 매기고, docx · 이관 문서 · 검토 엑셀의 장 번호까지 같이 바꾼다.
  - 다른 영역으로 옮기면 간지 목록 항목이 새 영역 간지로 옮겨지고 목차 범위가 간지 위치로 다시 계산된다. 목차 「담은 내용」 · 간지 설명 문구는 손으로 확인한다.
  - 간지 자체 이동, 두 런에 나뉜 「N장」, 100장 이상 덱의 장 번호는 자동으로 바뀌지 않는다.

## 색 정리

- 받은 파일에 토큰 밖 이전 값(회색 7A6A5E, 빨강 F53148 · D32F2F · E03C31 · C00000)이 있으면 `fix_colors(p)`로 토큰 값(7A6A5D · FF0000)으로 일괄 교체한다. 사용자가 정한 표지 · 간지 색(`references/design-common.md` 「깔끔한 장표」)은 바꾸지 않는다.

## 용어 · 명칭 일괄 치환

- 런 단위로 바꾸고, 한 문구가 여러 런에 걸친 곳은 목록으로 받아 따로 고친다. 치환 뒤 넘친 칸은 `check_text.py` 목록대로 `merged_text` → `set_cell`로 다시 나눈다.
- 명칭 · 결정 문구를 모든 산출물(pptx · docx)에서 함께 바꿀 때는 `rename_pass.py`(런 경계에 걸쳐도 서식 유지, `PG_RULES`, `tighten`)를 쓴다.

## 받은 덱 전체를 기준대로 정리

사용자가 새 판(예: v0.27)을 주며 「템플릿 · 문단 · 문장 형식 개선」을 요청하면 아래 순서로 한다. 사용자가 고친 글 · 빨강 · 굵게는 바꾸지 않는다.

1. `python3 check_text.py 입력.pptx`로 [정리] · [기호] · [단어 끊김 방지]를 세고, 표지 · 간지 색 · 끝선 · 하단 기준을 확인한다.
2. `python3 tidy_tables.py 입력.pptx 중간.pptx`: 한 줄에 들어가는 「- 」 정보를 합치고 열 폭 · 행 높이를 맞춘다(첫 열, 「- 」가 아닌 칸, 서식이 섞인 칸은 그대로. 하단이 내려가면 그 표는 되돌림).
3. 문장 끝맺음이 서술형으로 남은 칸이 있을 때만 `gaejo_pass.py`를 쓴다. 줄 끝 「·」만 있는 칸은 [참고]라 고치지 않는다.
4. 장 단위로 표지 · 간지 색, 오른쪽 끝선, 본문 하단을 고친다(`references/design-common.md` 「장르별 여백 · 본문 영역」, 하단을 넘는 본문 글상자는 줄 수 높이로 줄이고 위 요소와 겹치지 않는 범위에서 올림).
5. `python3 pptx_edit.py finalize 중간.pptx 출력.pptx` → 원본과 글 · 빨강 · 굵게가 같은지 대조 → `references/qa-checklist.md`.
6. 버전은 올리지 않는다(`references/final-deliverable.md` 「버전 표기」).

## 사용자 수정 장을 전 장으로 확장

사용자가 한 장을 고쳐 보내며 "전 화면 확장"을 요청하면:

1. 그 장에서 바뀐 서식 · 서술 패턴을 모두 뽑는다(원본과 run 단위 비교).
2. 먼저 `fix_colors(p)`로 이전 빨강을 FF0000으로 바꾼다(아래 도구는 FF0000만 강조로 알아봄). 그다음 같은 성격의 모든 위치에 적용한다: `format_pass.py`(빨강 · 글머리 열 · 빈 칸 · 요약 띠 · 확인 태그 · 순서 번호 · 바뀐 표 맞춤), 칸 문구를 새로 쓸 때는 `edit_lib.py`.
3. 장별 수정을 끝낸 뒤 손으로 고치지 않은 표 칸 · 설명 · 도식 글을 `gaejo_pass.py`로 한 번에 개조식 · 글머리로 정리한다.
4. 뽑은 패턴이 이 스킬에 없는 규칙이면 답변 끝에 "스킬 반영 후보"로 알린다(반영은 `skill-maintainer` 절차).

## 검수

- `references/qa-checklist.md`의 공통 항목 + 사용자 판 항목. 원본에 있던 `revisionInfo` 검증 경고는 그대로 둔다.

## `pptx_edit.py` 사용 예

```python
import sys; sys.path.insert(0, '/home/claude')   # scripts/를 작업 폴더로 복사한 뒤
from pptx import Presentation
from pptx_edit import *
p = Presentation('사용자판.pptx'); S = lambda n: p.slides[n - 1]

t = tbl(S(13)).table                                   # 칸 문구 교체 → 행 높이 → 겹침 확인
set_cell(t, 1, 1, '- 메가MGC커피 서버가 결제 거래번호 발급 ②에 payType(카드 구분)을 보냄\n- 핏콜라보가 authYN(인증 필요 여부)으로 돌려줌')
fit_rows(t, minh=0.4); print(bottom(tbl(S(13))))      # 핵심 한 줄 top(6.39)보다 작아야 함

new = duplicate_slide(p, 35, 36); renumber(p, 37)     # 36장 복제 → 37장으로, 번호 전부 재계산
add_source_line(p, '※ API 번호(…)는 34장 목록 기준', skip=(1, 2, 34))
replace_picture(S(42), '/home/claude/i42.png')        # 같은 비율로 자른 흐름도 이미지
box(S(33), 7.0, 1.95, 5.73, 0.6, ['인증 모듈(Webview 안)', '카드사 앱 주소를 열어 줌'], fill=FC_PINK, red_border=True)
arrow(S(33), 9.865, 2.55, 9.865, 2.78)
print(fix_colors(p))                                   # 이전 값 회색 · 빨강을 토큰 값으로
p.save('out.pptx')   # → python3 pptx_edit.py finalize out.pptx out.pptx → references/qa-checklist.md
```
