"""재편 전후 보존 대조. 사용: python3 coverage_check.py <원본 SKILL.md> <새 스킬 폴더> [--until 줄번호] [--min 0.85]

원본의 문장 · 표 칸을 하나씩 새 스킬의 모든 .md(CHANGELOG 제외)에서 찾는다.
글자 4-gram 포함률이 기준(기본 0.85)보다 낮은 원문을 출력한다.
출력된 줄은 「옮김(표현만 바뀜) · 의도적 대체 · 삭제(사유)」 중 하나로 분류해 보고한다.
--until: 원본에서 변경 이력 · 부록 코드가 시작하기 전 줄까지만 대조할 때."""
import sys, re, glob, os, argparse

ap = argparse.ArgumentParser()
ap.add_argument('orig'); ap.add_argument('new_dir')
ap.add_argument('--until', type=int, default=0); ap.add_argument('--min', type=float, default=0.85)
a = ap.parse_args()

lines = open(a.orig, encoding='utf-8').read().split('\n')
if a.until: lines = lines[:a.until]
new = ''.join(open(f, encoding='utf-8').read()
              for f in glob.glob(os.path.join(a.new_dir, '**', '*.md'), recursive=True)
              if os.path.basename(f) != 'CHANGELOG.md')
norm = lambda t: re.sub(r'[\s`*_「」()\[\]·,.:;"\'→~\-—|]', '', t)
G = lambda s: {s[i:i + 4] for i in range(len(s) - 3)}
NG = G(norm(new))
low, total, code = [], 0, False
for i, l in enumerate(lines, 1):
    if l.strip().startswith('```'): code = not code; continue
    if code or l.startswith('#') or not l.strip() or set(l.strip()) <= set('|-: '): continue
    parts = l.strip().strip('|').split('|') if l.strip().startswith('|') else [l]
    for p in parts:
        n = norm(p)
        if len(n) < 12: continue
        total += 1
        g = G(n); cov = len(g & NG) / max(len(g), 1)
        if cov < a.min: low.append((cov, i, p.strip()))
print(f'# 보존 대조: 원문 단위 {total}개 중 포함률 {a.min:.0%} 미만 {len(low)}개\n')
print('| 포함률 | 원본 줄 | 원문 | 분류(옮김 / 대체 / 삭제 사유) |\n|---|---|---|---|')
for cov, i, p in sorted(low):
    print(f'| {cov:.0%} | {i} | {p[:120].replace("|", "/")} | |')
