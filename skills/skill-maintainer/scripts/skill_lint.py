"""스킬 구조 점검기. 사용: python3 skill_lint.py <스킬 폴더> [--config lint.json] [--out 보고서.md]

스킬 폴더의 SKILL.md · references/*.md · scripts/* 를 읽고 아래를 점검해 마크다운 보고서로 출력한다.
문제를 고치지는 않는다. 오류가 1건 이상이면 종료 코드 1.

  [오류] 반드시 고칠 것
    1. 앞머리(frontmatter) name · description 누락, description 1,024자 초과
    2. 본문이 가리키는 파일(scripts/x.py, references/x.md 등)이 폴더에 없음
    3. 본문이 가리키는 제목(「references/x.md」 「제목」, 또는 4-C-12 같은 절 번호)이 없음
    4. 설정의 기준값(facts)과 다른 값이 적힌 곳
    5. 설정의 금지어(banned)가 허용 위치 밖에 남은 곳
  [경고] 검토할 것
    6. 크기 예산 초과(SKILL.md 줄 수 · 바이트, 참조 파일 줄 수)
    7. 중복 후보: 다른 절에 거의 같은 문장이 있는 곳(글자 3-gram 유사도)
    8. 스크립트 표류: 변경 이력이 그 스크립트를 마지막으로 고친 버전보다 스크립트 머리 버전이 낮음
    9. 고아 파일: scripts/ · references/ 에 있으나 어디에서도 가리키지 않는 파일
   10. 읽기 지도 누락: references/*.md 가 SKILL.md 에 언급되지 않음
  [참고]
   11. 규칙 본문 속 버전 꼬리표 「(v3.8 …)」 개수 — 이력은 CHANGELOG.md 로 옮긴다
   12. 절 구성: 제목 깊이 · 절마다 줄 수 상위

설정(lint.json, 없으면 기본값):
{
  "budget": {"skill_lines": 300, "skill_bytes": 30000, "ref_lines": 450},
  "dup_threshold": 0.72, "dup_min_chars": 28,
  "history_files": ["CHANGELOG.md"],
  "facts": [{"label": "본문 하단", "pattern": "정규식(값 1개를 괄호로)", "expect": "6.85", "scripts": false}],
  "banned": [{"term": "CJ PG", "reason": "명칭은 OnePG", "allow_files": ["CHANGELOG.md"],
              "allow_if_line_contains": ["금지", "(×)"], "scripts": false}]
}
"""
import sys, re, json, os, argparse, itertools
from collections import defaultdict

DEFAULT = {
    "budget": {"skill_lines": 300, "skill_bytes": 30000, "ref_lines": 450},
    "dup_threshold": 0.72, "dup_min_chars": 28,
    "history_files": ["CHANGELOG.md"],
    "facts": [], "banned": [],
    "external_files": [],  # 다른 스킬 · 환경에 있는 파일(예: validate.py)
    "topics": [],          # [{"name", "keywords": [...], "home": "references/x.md", "allow_files": [...]}]
    "pointer_files": ["SKILL.md"],  # 주제를 가리키기만 하는 파일(읽기 지도 · 검수표)
}
NUM_RE = re.compile(r'(\d+(?:\.\d+)?(?:\s*~\s*\d+(?:\.\d+)?)?)\s*(pt|in|곳|자|줄|개|px)(?![가-힣])')
FILE_RE = re.compile(r'(?<![\w/.-])((?:scripts/|references/)?[\w.-]+\.(?:py|js|md|json|sh))(?![\w])')
HEAD_REF_RE = re.compile(r'(references/[\w.-]+\.md)\s*「([^」]+)」')
SECNUM_RE = re.compile(r'(?<![\w-])(\d+-[A-Z](?:-\d+)?)(?![\w-])')
VTAG_RE = re.compile(r'\((?:v[1-9]\d*\.\d+[^)]*)\)|\bv[1-9]\d*\.\d+ (?:개정|추가|최종)')
VER_RE = re.compile(r'v(\d+)\.(\d+)')
IGNORE_FILES = {'pptx', 'docx', 'xlsx'}  # 확장자 아님, 안전장치


def ver(s):
    m = VER_RE.search(s or '')
    return (int(m.group(1)), int(m.group(2))) if m else None


def read(p):
    with open(p, encoding='utf-8') as f:
        return f.read()


def strip_code(text):
    """``` 코드 블록을 빈 줄로 바꿔 줄 번호는 유지한다."""
    out, on = [], False
    for line in text.split('\n'):
        if line.strip().startswith('```'):
            on = not on; out.append(''); continue
        out.append('' if on else line)
    return '\n'.join(out)


def headings(text):
    hs = []
    for i, line in enumerate(text.split('\n'), 1):
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            hs.append((i, len(m.group(1)), m.group(2).strip()))
    return hs


def section_of(hs, lineno):
    cur = ''
    for i, _, t in hs:
        if i <= lineno: cur = t
        else: break
    return cur


def units(text):
    """문장 · 글머리 · 표 칸 단위로 쪼갠다. (줄 번호, 원문, 정규화)"""
    res = []
    for i, line in enumerate(text.split('\n'), 1):
        s = line.strip()
        if not s or s.startswith('#') or set(s) <= set('|-: '):
            continue
        parts = [c for c in s.strip('|').split('|')] if s.startswith('|') else re.split(r'(?<=[다음함됨]\.)\s+|(?<=\.)\s+(?=[가-힣A-Z「(])', s)
        for p in parts:
            raw = p.strip()
            body = re.sub(r'`[^`]*`|references/[\w.-]+|「[^」]{0,30}」', '', raw)
            norm = re.sub(r'[\s`*_「」()\[\]·,.:;"\'→~\-—|]', '', body)
            norm = re.sub(r'v\d+\.\d+', '', norm)
            if norm:
                res.append((i, raw, norm))
    return res


def grams(s, n=3):
    return {s[i:i + n] for i in range(max(len(s) - n + 1, 1))}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('skill_dir'); ap.add_argument('--config'); ap.add_argument('--out')
    a = ap.parse_args()
    root = os.path.abspath(a.skill_dir)
    cfg = json.loads(json.dumps(DEFAULT))
    cpath = a.config or os.path.join(root, 'lint.json')
    if os.path.exists(cpath):
        user = json.loads(read(cpath))
        for k, v in user.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict): cfg[k].update(v)
            else: cfg[k] = v

    docs = {}  # 상대 경로 → 글
    for dp, _, fs in os.walk(root):
        for f in fs:
            if f.endswith('.md'):
                rel = os.path.relpath(os.path.join(dp, f), root)
                docs[rel] = read(os.path.join(dp, f))
    scripts = {}
    sdir = os.path.join(root, 'scripts')
    if os.path.isdir(sdir):
        for f in sorted(os.listdir(sdir)):
            p = os.path.join(sdir, f)
            if os.path.isfile(p):
                try: scripts['scripts/' + f] = read(p)
                except UnicodeDecodeError: scripts['scripts/' + f] = ''
    all_files = set(docs) | set(scripts) | {os.path.relpath(os.path.join(dp, f), root) for dp, _, fs in os.walk(root) for f in fs}
    basenames = defaultdict(set)
    for f in all_files: basenames[os.path.basename(f)].add(f)
    hist = set(cfg['history_files'])
    rule_docs = {k: v for k, v in docs.items() if os.path.basename(k) not in hist}

    E, W, I = [], [], []

    # 1. 앞머리
    skill = docs.get('SKILL.md')
    if skill is None:
        E.append(('SKILL.md', '-', 'SKILL.md 가 없음')); skill = ''
    fm = re.match(r'^---\n(.*?)\n---', skill, re.S)
    if not fm:
        E.append(('SKILL.md', 1, '앞머리(--- name/description ---)가 없음'))
    else:
        body = fm.group(1)
        for key in ('name', 'description'):
            if not re.search(rf'^{key}\s*:', body, re.M):
                E.append(('SKILL.md', 1, f'앞머리 {key} 누락'))
        m = re.search(r'^description\s*:\s*(.*)', body, re.M)
        if m and len(m.group(1)) > 1024:
            E.append(('SKILL.md', 1, f'description {len(m.group(1))}자 — 1,024자 이하로'))
    skill_ver = ver(re.search(r'^#\s.*', skill, re.M).group(0)) if re.search(r'^#\s.*', skill, re.M) else None

    # 2 · 3. 파일 · 제목 참조
    mentioned = defaultdict(set)
    heads = {k: headings(v) for k, v in docs.items()}
    all_secnums = {}
    for k, hs in heads.items():
        for _, _, t in hs:
            for n in SECNUM_RE.findall(t): all_secnums.setdefault(n, k)
    for k, txt in docs.items():
        if os.path.basename(k) in hist:
            continue
        plain = strip_code(txt)  # 코드 블록 · 템플릿 속 경로는 예시로 본다
        for i, line in enumerate(plain.split('\n'), 1):
            for m in FILE_RE.finditer(line):
                ref = m.group(1)
                base = os.path.basename(ref)
                if base.split('.')[-1] in IGNORE_FILES or base in ('SKILL.md',) or base in cfg['external_files']:
                    continue
                if re.search(r'[가-힣<>]', base):  # 「스크립트.js」 같은 자리 표시
                    continue
                if base == os.path.basename(k):
                    continue
                if base in basenames:
                    for f in basenames[base]: mentioned[f].add(k)
                elif '/' in ref or k == 'SKILL.md' or base.startswith(('check_', 'slim', 'kit', 'mgc_', 'fit_', 'pptx_', 'edit_', 'format_', 'gaejo_', 'rename_', 'move_', 'compare_', 'crop_', 'example_', 'flow_')):
                    E.append((k, i, f'가리키는 파일 없음: {ref}'))
            if os.path.basename(k) in hist:
                continue
            for m in HEAD_REF_RE.finditer(line):
                f, h = m.group(1), m.group(2)
                if f not in docs:
                    E.append((k, i, f'가리키는 파일 없음: {f}'))
                elif not any(h in t for _, _, t in heads[f]):
                    E.append((k, i, f'가리키는 제목 없음: {f} 「{h}」'))
            for n in SECNUM_RE.findall(line):
                if n not in all_secnums and not any(s.startswith(n + '-') for s in all_secnums):
                    E.append((k, i, f'가리키는 절 번호 없음: {n}'))

    # 4 · 5. 기준값 · 금지어
    def targets(use_scripts):
        t = dict(rule_docs)
        if use_scripts: t.update(scripts)
        return t
    for fct in cfg['facts']:
        rx = re.compile(fct['pattern'])
        for k, txt in targets(fct.get('scripts')).items():
            if os.path.basename(k) in hist: continue
            for i, line in enumerate(txt.split('\n'), 1):
                for m in rx.finditer(line):
                    if m.group(1) != str(fct['expect']):
                        E.append((k, i, f"기준값 불일치 [{fct['label']}] {m.group(1)} → 기준 {fct['expect']}: {line.strip()[:90]}"))
    for b in cfg['banned']:
        allow_f = set(b.get('allow_files', [])) | hist
        allow_l = b.get('allow_if_line_contains', [])
        for k, txt in targets(b.get('scripts')).items():
            if os.path.basename(k) in allow_f: continue
            for i, line in enumerate(txt.split('\n'), 1):
                if b['term'] in line and not any(x in line for x in allow_l):
                    E.append((k, i, f"금지어 「{b['term']}」({b.get('reason', '')}): {line.strip()[:90]}"))

    grouped, order = defaultdict(list), []
    for k, i, msg in E:
        key = (k, msg)
        if key not in grouped: order.append(key)
        grouped[key].append(str(i))
    E = [(k, ', '.join(grouped[(k, m)][:8]) + (' …' if len(grouped[(k, m)]) > 8 else ''), m) for k, m in order]

    # 6. 크기
    bud = cfg['budget']
    n_lines, n_bytes = skill.count('\n') + 1, len(skill.encode('utf-8'))
    if n_lines > bud['skill_lines'] or n_bytes > bud['skill_bytes']:
        W.append(('SKILL.md', '-', f'크기 예산 초과: {n_lines}줄 / {n_bytes:,}바이트 (예산 {bud["skill_lines"]}줄 / {bud["skill_bytes"]:,}바이트)'))
    for k, txt in docs.items():
        if k.startswith('references/') and txt.count('\n') + 1 > bud['ref_lines']:
            W.append((k, '-', f'참조 파일 {txt.count(chr(10)) + 1}줄 — 예산 {bud["ref_lines"]}줄, 주제별로 나눌 것'))

    # 7. 중복 후보
    pool = []
    for k, txt in rule_docs.items():
        hs = heads[k]
        for i, raw, norm in units(strip_code(txt)):
            if len(norm) >= cfg['dup_min_chars']:
                pool.append((k, i, section_of(hs, i), raw, norm, grams(norm)))
    index = defaultdict(set)
    for idx, u in enumerate(pool):
        for g in u[5]: index[g].add(idx)
    seen, dups = set(), []
    for idx, u in enumerate(pool):
        cand = defaultdict(int)
        for g in u[5]:
            if len(index[g]) > 60: continue
            for j in index[g]:
                if j > idx: cand[j] += 1
        for j, c in cand.items():
            v = pool[j]
            if (u[0], u[2]) == (v[0], v[2]): continue
            sim = c / max(len(u[5] | v[5]), 1)
            if sim >= cfg['dup_threshold'] and (idx, j) not in seen:
                seen.add((idx, j)); dups.append((sim, u, v))
    dups.sort(key=lambda x: -x[0])
    for sim, u, v in dups[:60]:
        W.append((f'{u[0]}:{u[1]}', f'{v[0]}:{v[1]}', f'중복 후보 {sim:.0%} 「{u[2][:24]}」 ↔ 「{v[2][:24]}」: {u[3][:70]}'))

    # 7-1. 주제 흩어짐 · 값 충돌 후보
    for tp in cfg['topics']:
        kws, home = tp['keywords'], tp.get('home')
        hits = []
        for k, txt in rule_docs.items():
            hs = heads[k]
            for i, line in enumerate(strip_code(txt).split('\n'), 1):
                if line.lstrip().startswith('#'): continue
                if any(w in line for w in kws):
                    hits.append((k, i, section_of(hs, i), line.strip()))
        secs = sorted({(k, sec) for k, _, sec, _ in hits})
        if home:
            stray = [(k, i, sec, l) for k, i, sec, l in hits
                     if k != home and k not in cfg['pointer_files'] and k not in tp.get('allow_files', [])
                     and os.path.basename(home) not in l]
            for k, i, sec, l in stray[:12]:
                W.append((f'{k}:{i}', home, f"흩어진 규칙 [{tp['name']}] 기준 위치 밖에서 규칙을 다시 씀 — 기준 위치를 가리키도록: {l[:70]}"))
        elif len(secs) >= 3:
            W.append(('-', '-', f"주제 흩어짐 [{tp['name']}] {len(secs)}개 절: " + ', '.join(f'「{sec[:18]}」' for _, sec in secs[:10])))
        byunit = defaultdict(lambda: defaultdict(set))
        for k, i, sec, l in hits:
            for v, u in NUM_RE.findall(l):
                byunit[u][re.sub(r'\s', '', v)].add(f'{k}:{i}')
        for u, vals in byunit.items():
            locs = set().union(*vals.values())
            if 2 <= len(vals) <= 6 and len(locs) >= 2:
                I.append(('-', '-', f"값 비교 [{tp['name']}] 단위 {u}: " + ' / '.join(f'{v}({", ".join(sorted(ls)[:3])})' for v, ls in vals.items()) + ' — 같은 대상이면 하나로'))

    # 8. 스크립트 표류
    changelog = ''.join(v for k, v in docs.items() if os.path.basename(k) in hist) or skill
    entries = [e.split('\n#')[0] for e in re.split(r'\n(?=- v\d+\.\d+)', changelog)]
    for s, txt in scripts.items():
        base = os.path.basename(s)
        last = None
        for e in entries:
            if base in e and not re.search(re.escape(base) + r'[^\n]{0,25}(변경 없음|그대로)', e):
                v = ver(e[:20])
                if v and (last is None or v > last): last = v
        head = '\n'.join(txt.split('\n')[:3])
        hv = ver(head)
        if last and hv and hv < last:
            W.append((s, 1, f'스크립트 표류: 머리 버전 v{hv[0]}.{hv[1]}, 변경 이력의 마지막 수정 v{last[0]}.{last[1]} — 최신본인지 확인'))
        elif last and not hv:
            I.append((s, 1, f'스크립트 머리에 버전 표기 없음(변경 이력 마지막 수정 v{last[0]}.{last[1]})'))

    # 9 · 10. 고아 · 읽기 지도
    for f in sorted(all_files):
        if f.startswith(('scripts/', 'references/')) and not mentioned.get(f):
            W.append((f, '-', '고아 파일: 어디에서도 가리키지 않음'))
    for f in docs:
        if f.startswith('references/') and os.path.basename(f) not in skill:
            W.append(('SKILL.md', '-', f'읽기 지도 누락: {f} 를 언제 읽는지 SKILL.md 에 없음'))

    # 11 · 12. 참고
    vt = []
    for k, txt in rule_docs.items():
        for i, line in enumerate(strip_code(txt).split('\n'), 1):
            if line.startswith('# '): continue
            for m in VTAG_RE.finditer(line): vt.append((k, i, m.group(0)))
    if vt:
        I.append(('-', '-', f'규칙 본문 속 버전 꼬리표 {len(vt)}개 — 예: ' + ', '.join(f'{k}:{i} {t[:20]}' for k, i, t in vt[:5])))
    sizes = []
    for k, hs in heads.items():
        lines = docs[k].count('\n') + 1
        for idx, (i, lv, t) in enumerate(hs):
            end = next((j for j, l2, _ in hs[idx + 1:] if l2 <= lv), lines + 1)
            sizes.append((end - i, k, t))
    sizes.sort(reverse=True)
    I.append(('-', '-', '긴 절 상위: ' + ', '.join(f'{k} 「{t[:20]}」 {n}줄' for n, k, t in sizes[:6])))
    deep = [(k, t) for k, hs in heads.items() for _, lv, t in hs if lv >= 5]
    if deep:
        I.append(('-', '-', f'제목 깊이 5 이상 {len(deep)}개 — 위계를 3단(##/###/####) 이내로'))

    # 보고서
    total = sum(len(v) for v in docs.values())
    out = [f'# 스킬 구조 점검 결과: {os.path.basename(root)}', '',
           f'- 파일: 문서 {len(docs)}개 · 스크립트 {len(scripts)}개 · 문서 합계 {total:,}자',
           f'- SKILL.md: {n_lines}줄 · {n_bytes:,}바이트' + (f' · v{skill_ver[0]}.{skill_ver[1]}' if skill_ver else ''),
           f'- 오류 {len(E)}건 · 경고 {len(W)}건 · 참고 {len(I)}건', '']
    for title, rows in (('오류(반드시 고칠 것)', E), ('경고(검토할 것)', W), ('참고', I)):
        out += [f'## {title}', '']
        if not rows: out += ['- 없음', '']; continue
        out += ['| 위치 | 위치 2 | 내용 |', '|---|---|---|']
        out += [f'| {a1} | {a2} | {str(c).replace("|", "/")} |' for a1, a2, c in rows]
        out.append('')
    rep = '\n'.join(out)
    if a.out:
        with open(a.out, 'w', encoding='utf-8') as f: f.write(rep)
    print(rep)
    sys.exit(1 if E else 0)


if __name__ == '__main__':
    main()
