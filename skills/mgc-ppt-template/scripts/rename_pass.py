"""rename_pass.py — 문단 단위 문구 치환(런 서식 유지, pptx · docx 공통) · 명칭 통일 · 표 행 높이 재계산 (mgc-ppt-template v4.2)"""
import re
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PG_RULES = [('CJ ONS(OnePG)', 'OnePG'), ('CJ ONS를 추가', 'OnePG를 추가'), ('CJ ONS 추가', 'OnePG 추가'),
            ('pgType = CJ ONS', 'pgType = OnePG'), ('CJ PG', 'OnePG')]

def para_replace(p_el, old, new, ns):
    """한 문단에서 old → new. 여러 런에 걸쳐도 시작 런 서식으로 넣고 나머지 런에서 지운다"""
    ts = [t for t in p_el.iter(ns + 't')]
    full = ''.join(t.text or '' for t in ts); n = 0; start = 0
    if old == new: return 0
    while full.find(old, start) >= 0:
        i = full.index(old, start); j = i + len(old); pos = 0; placed = False; start = i + len(new)
        for t in ts:
            s = t.text or ''; a, b = pos, pos + len(s); pos = b
            if b <= i or a >= j: continue
            lo, hi = max(i, a) - a, min(j, b) - a
            t.text = s[:lo] + ('' if placed else new) + s[hi:]; placed = True
        full = ''.join(t.text or '' for t in ts); n += 1
    return n

def pg_unify(root, ns):
    n = 0
    for p in root.iter(ns + 'p'):
        for o, w in PG_RULES: n += para_replace(p, o, w, ns)
    return n

def tighten(sh, size=10.5, top_in=None):
    """표 행 높이를 문단 수 × 글자 크기 × 1.2 + 0.12in로 다시 잡는다(줄 수 추정 과대로 하단 각주와 겹칠 때)"""
    from pptx.util import Inches
    if top_in: sh.top = Inches(top_in)
    for row in sh.table.rows:
        n = max(len(c.text_frame.paragraphs) for c in row.cells)
        row.height = Inches(max(0.30, n * size * 1.2 / 72 + 0.12))
    return round((sh.top + sum(r.height for r in sh.table.rows)) / 914400, 2)
