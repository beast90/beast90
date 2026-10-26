"""gaejo_pass.py — 개조식 · 글머리 「•」 · 나열 줄바꿈 자동 정리와 표 맞춤 (mgc-ppt-template v4.1, 4-C-12)
사용:
    from gaejo_pass import run
    run(p, done=[손으로 고친 칸의 _tc · 도형 _element ...])   # 손으로 고친 칸은 다시 쓰지 않고 표 맞춤만 받는다
    또는 python3 gaejo_pass.py in.pptx out.pptx
- 제목 · 핵심 문장 · 영역 이름 · 간지 목록 · 캡션 · 범례 · 분기 이름 · 출처는 건드리지 않는다(서술형 유지)
- 표 칸 · 「설명」 · 「본문」 상자: 정보마다 「• 」 줄, 끝맺음 개조식, 이어지는 줄은 문단 들여쓰기
- 그 밖의 도식 상자: 문단 구조를 지키고 문단마다 끝맺음 · 글머리만 바꾼다
- 표가 아래 요소와 겹치면 행 높이 → 글자 크기(0.5pt씩, 10.5pt까지) 순서로 맞춘다"""
import sys, re
sys.path.insert(0, __import__('os').path.dirname(__file__))
from pptx.util import Inches
from pptx.oxml.ns import qn
import pptx_edit as PE
from edit_lib import BULLET, gaejo, red_text, set_items, cell_w, row_h

def run(p, done=()):
    DONE = list(done)
    SKIP_NAMES = {'제목', '핵심 문장', '핵심 문장 요약', '영역 이름', '간지 목록', '화면 캡션', '범례', '분기 이름', '출처', '구역 이름', '도식 제목'}
    def done(el): return any(e is el for e in DONE)
    def norm(s): return re.sub(r'\s+', '', s.replace('[[', '').replace(']]', ''))
    STAT = {'table': 0, 'box': 0, 'para': 0}
    PROC = []
    def logical(tf):
        lines = red_text(tf); out = []
        for raw in lines:
            s = raw.strip()
            if not s: continue
            head = s.startswith('- ') or s.startswith('• ') or s.startswith('※')
            if s.startswith('- '): s = BULLET + s[2:]
            if out and (raw.startswith('  ') or not head) and not out[-1].endswith('.'):
                joiner = '' if out[-1].endswith(('(', '/', '·')) and False else ' '
                out[-1] = out[-1] + joiner + s
            else: out.append(s)
        res = []
        for it in out:   # 한 칸에 문장이 여럿이면 정보마다 나눈다
            lead = BULLET if it.startswith(BULLET) else ('※ ' if it.startswith('※ ') else '')
            body = it[len(lead):]
            parts = [x for x in re.split(r'(?<=[다음함됨])\.\s+', body) if x]
            for k, part in enumerate(parts): res.append((lead if (k == 0 or lead == '※ ') else lead) + gaejo(part))
        if len(res) >= 2 and not any(r.startswith((BULLET, '※')) for r in res) and len(out) < len(res):
            res = [BULLET + r for r in res]
        return [r.replace('  ', ' ') for r in res]

    for s in p.slides:
        if any(sh.name == '간지 목록' for sh in s.shapes): continue
        for sh in s.shapes:
            if sh.has_table:
                t = sh.table
                for ri, row in enumerate(t.rows):
                    changed = False
                    for ci, c in enumerate(row.cells):
                        if done(c._tc): changed = True; continue
                        old = [pa.text for pa in c.text_frame.paragraphs]
                        if not ''.join(old).strip(): continue
                        new = logical(c.text_frame)
                        dangling = any(o.rstrip().endswith('·') or o.lstrip().startswith('· ') for o in old)
                        hyph = any(o.lstrip().startswith('- ') for o in old)
                        if norm(' '.join(new)) == norm(' '.join(old)) and not dangling and not hyph: continue
                        set_items(c.text_frame, new, cell_w(t, ci)); changed = True; STAT['table'] += 1; PROC.append(c._tc)
                    if changed:
                        h = Inches(row_h(t, ri))
                        if h > row.height: row.height = h
            elif sh.has_text_frame and sh.name not in SKIP_NAMES and not done(sh._element):
                tf = sh.text_frame; old = [pa.text for pa in tf.paragraphs]
                if not ''.join(old).strip() or re.match(r'^\d+ / \d+$', tf.text.strip()): continue
                if sh.name in ('설명', '본문'):
                    new = logical(tf)
                    if norm(' '.join(new)) == norm(' '.join(old)) and not any(o.lstrip().startswith('- ') for o in old): continue
                    set_items(tf, new, sh.width / 914400); STAT['box'] += 1
                else:   # 도식 상자 등: 문단 구조 유지, 문단마다 끝맺음 · 글머리만
                    for pa in tf.paragraphs:
                        t0 = pa.text
                        t1 = gaejo(t0)
                        if t0.startswith('- '): t1 = BULLET + gaejo(t0[2:])
                        if t1 != t0 and pa.runs:
                            pa.runs[0].text = t1
                            for r in pa.runs[1:]: r.text = ''
                            STAT['para'] += 1
    print('auto', STAT)

    # ── 표 맞춤: 표가 아래 요소(핵심 한 줄 · 도식 · 쪽 번호)를 침범하면 행 높이 → 글자 크기(10.5pt까지) 순서로 줄인다
    def set_tbl_size(t, size):
        for e in list(t._tbl.iter(qn('a:rPr'))) + list(t._tbl.iter(qn('a:endParaRPr'))): e.set('sz', str(int(size * 100)))
    def tbl_size(t):
        s, _ = PE.run_props(t.cell(1 if len(t.rows) > 1 else 0, len(t.columns) - 1).text_frame); return s
    FIT = []
    for si, s in enumerate(p.slides, 1):
        for sh in [x for x in s.shapes if x.has_table]:
            t = sh.table; top = sh.top / 914400; L, R = sh.left, sh.left + sh.width
            below = [o.top / 914400 for o in s.shapes if o is not sh and o.top / 914400 > top + 0.3
                     and not (o.left + o.width <= L or R <= o.left) and not (o.has_text_frame and re.match(r'^\d+ / \d+$', o.text_frame.text.strip()))]
            limit = min(below + [6.9]) - 0.08
            def bot(): return top + sum(r.height for r in t.rows) / 914400
            if bot() <= limit + 0.005: continue
            for ri in range(len(t.rows)): t.rows[ri].height = Inches(row_h(t, ri, 0.32))
            size = tbl_size(t)
            while bot() > limit + 0.005 and size > 10.5:
                size -= 0.5; set_tbl_size(t, size)
                for ri, row in enumerate(t.rows):
                    for ci, c in enumerate(row.cells):
                        if not c.text_frame.text.strip(): continue
                        if done(c._tc) or any(e is c._tc for e in PROC):
                            set_items(c.text_frame, logical(c.text_frame), cell_w(t, ci))
                    t.rows[ri].height = Inches(row_h(t, ri, 0.32))
            FIT.append((si, round(size, 1), round(bot(), 2), round(limit, 2)))
    print('fit', FIT)

    return STAT, FIT

if __name__ == '__main__':
    from pptx import Presentation
    p = Presentation(sys.argv[1]); run(p); p.save(sys.argv[2])
