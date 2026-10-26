# 점검기 설정(lint.json) 작성법

스킬 폴더에 `lint.json`을 두면 `scripts/skill_lint.py`가 자동으로 읽는다. 없으면 기본값으로 크기 · 참조 · 중복 · 표류만 본다.

## 항목

| 항목 | 뜻 | 언제 고치나 |
|---|---|---|
| `budget` | `skill_lines` · `skill_bytes` · `ref_lines` 크기 예산 | 스킬 성격상 더 커야 할 때만 |
| `dup_threshold` · `dup_min_chars` | 중복 후보 판정 유사도(0~1) · 최소 글자 수 | 경고가 너무 많거나 적을 때 |
| `history_files` | 버전 꼬리표 · 금지어 · 없는 파일 검사에서 빼는 이력 파일 | 보통 `CHANGELOG.md` |
| `external_files` | 다른 스킬 · 환경에 있는 파일(없는 파일 오류에서 제외) | 다른 스킬 스크립트를 부를 때 |
| `pointer_files` | 주제를 가리키기만 하는 파일(읽기 지도 · 검수표). 흩어진 규칙 검사에서 제외 | 검수 파일을 만들었을 때 |
| `facts` | 기준값. `pattern`의 괄호 1개가 잡은 값이 `expect`와 다르면 오류. `scripts: true`면 스크립트도 검사 | 값 규칙을 바꿀 때마다 같이 고침 |
| `banned` | 금지어. `allow_files` · `allow_if_line_contains`(예: 「금지」 · 「(×)」가 있는 줄)는 제외 | 명칭 · 표기 · 옛 값을 폐기할 때 추가 |
| `topics` | 주제별 핵심어와 기준 위치(`home`). 기준 위치 밖에서 같은 주제를 다시 쓰면 경고. `allow_files`로 예외 파일 지정. `home`이 없으면 주제가 흩어진 절 수만 보고 | 규칙 위치 지도에 주제를 추가할 때 |

## 예시

```json
{
  "budget": {"skill_lines": 260, "skill_bytes": 24000, "ref_lines": 300},
  "history_files": ["CHANGELOG.md"],
  "external_files": ["validate.py"],
  "pointer_files": ["SKILL.md", "references/qa-checklist.md"],
  "facts": [
    {"label": "줄바꿈 폭 보정", "pattern": "줄바꿈[^\\n]{0,20}?폭 보정\\s*(\\d\\.\\d+)", "expect": "0.82"}
  ],
  "banned": [
    {"term": "CJ PG", "reason": "결제대행사 명칭은 OnePG 하나", "allow_if_line_contains": ["금지", "(×)"]},
    {"term": "F53148", "reason": "강조 빨강은 FF0000 하나", "allow_if_line_contains": ["이전"], "scripts": true}
  ],
  "topics": [
    {"name": "글머리", "keywords": ["글머리", "하이픈"], "home": "references/writing.md"}
  ]
}
```

## 결과 읽는 법

- **오류**: 반드시 고친다. 없는 파일 · 제목 참조, 기준값 불일치, 금지어 잔존, 앞머리 누락.
- **경고**: 하나씩 판단한다. 「흩어진 규칙」은 기준 위치를 가리키도록 고치거나, 그 파일에 있어야 할 이유가 있으면 `topics[].allow_files`에 넣는다. 「스크립트 표류」는 최신 스크립트인지 확인한다.
- **참고**: 값 비교 · 긴 절 · 버전 꼬리표. 재편 판단 자료로 쓴다.
- 처음 만든 설정으로 경고가 30건을 넘으면 대부분 설정 문제다(핵심어가 너무 넓음). 핵심어를 좁힌 뒤 다시 돌린다.
