"""pptx_edit.py — 사용자가 다듬은 pptx를 python-pptx로 직접 고칠 때 쓰는 도구 (4-C-10, 부록 F)

from pptx_edit import *
- 줄나눔: wrap(text, w_in, size, bold) — 칸 폭 기준 띄어쓰기 단위, '- ' 항목 이어지는 줄은 두 칸 들여쓰기
- 글: set_tf(tf, text, w_in, nowrap) — 빨강이 아닌 첫 런 서식을 원본으로 문단 재구성, [[문구]]는 빨강 굵게
- 표: set_cell · fit_rows · del_row · add_col · del_col · tbl · bottom
- 도형: textbox · box · arrow (맑은 고딕, 정보 도형만)
- 장: duplicate_slide · renumber(p, inserted_at) · add_source_line · replace_picture
"""
import copy, re
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ── 폭 추정 (fit_text.js와 같은 표, 여유 0.88: LibreOffice·PowerPoint 실측 보정)
def em(ch, bold):
    c = ord(ch); b = 1.06 if bold else 1
    if (0xAC00 <= c <= 0xD7A3) or (0x1100 <= c <= 0x11FF) or (0x3130 <= c <= 0x318F) or (0x4E00 <= c <= 0x9FFF) or (0x3000 <= c <= 0x303F) or (0xFF00 <= c <= 0xFFEF):
        return 0.98 if bold else 0.96
    if ch == ' ': return 0.30
    if ch.isdigit(): return 0.56 * b
    if 'A' <= ch <= 'Z': return 0.64 * b
    if 'a' <= ch <= 'z': return 0.52 * b
    if ch in ".,:;'\"`|!‘’“”·": return 0.30
    if ch in '()[]{}': return 0.34
    if ch in '→←↑↓↔—@': return 1.0
    if ch in '~+=<>#': return 0.62
    if ch == '%': return 0.85
    if ch in '/\\': return 0.40
    if ch == '-': return 0.38
    return 0.92
def wpt(s, size, bold): return sum(em(c, bold) * size for c in s)
def wrap(text, w_in, size, bold=False, slack=0.88):
    avail = w_in * 72 * slack - 1.5; out = []
    for par in str(text).split('\n'):
        hy = par.startswith('- '); cur = ''
        for tok in par.split(' '):
            cand = tok if cur == '' else cur + ' ' + tok
            if cur != '' and wpt(re.sub(r'\[\[|\]\]', '', cand), size, bold) > avail:
                out.append(cur); cur = ('  ' if hy else '') + tok
            else: cur = cand
        out.append(cur)
    return out

RED = {'F53148', 'D32F2F', 'E03C31', 'FF0000'}
def _isred(r):
    rp = r.find(qn('a:rPr'))
    if rp is None: return False
    c = rp.find('.//' + qn('a:srgbClr'))
    return c is not None and c.get('val', '').upper() in RED
def run_props(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            rp = r._r.find(qn('a:rPr'))
            return (int(rp.get('sz')) / 100 if rp is not None and rp.get('sz') else 12), (rp is not None and rp.get('b') == '1')
    return 12, False
def set_tf(tf, text, w_in=None, size=None, bold=None, nowrap=False):
    sz, b = run_props(tf); size = size or sz; bold = b if bold is None else bold
    lines = str(text).split('\n') if (nowrap or w_in is None) else wrap(text, w_in, size, bold)
    body = tf._txBody; ps = body.findall(qn('a:p')); p0 = ps[0]
    allr = [r for p in ps for r in p.findall(qn('a:r'))]
    nonred = [r for r in allr if not _isred(r)]
    tmpl = copy.deepcopy(nonred[0] if nonred else allr[0])
    for p in ps: body.remove(p)
    for ln in lines:
        np_ = copy.deepcopy(p0)
        for r in np_.findall(qn('a:r')) + np_.findall(qn('a:br')) + np_.findall(qn('a:fld')): np_.remove(r)
        end = np_.find(qn('a:endParaRPr'))
        for part in re.split(r'(\[\[.*?\]\])', ln):
            if not part: continue
            red = part.startswith('[[')
            nr = copy.deepcopy(tmpl); nr.find(qn('a:t')).text = part[2:-2] if red else part
            if red:
                rp = nr.find(qn('a:rPr')); rp.set('b', '1')
                for f in rp.findall(qn('a:solidFill')): rp.remove(f)
                sf = rp.makeelement(qn('a:solidFill'), {}); sf.append(sf.makeelement(qn('a:srgbClr'), {'val': 'F53148'})); rp.insert(0, sf)
            (end.addprevious(nr) if end is not None else np_.append(nr))
        body.append(np_)
    return len(lines)

# ── 표
def tbl(slide, idx=0): return [sh for sh in slide.shapes if sh.has_table][idx]
def find(slide, name): return [sh for sh in slide.shapes if sh.name == name]
def bottom(sh):
    return (sh.top + sum(r.height for r in sh.table.rows)) / 914400 if sh.has_table else (sh.top + sh.height) / 914400
def set_cell(table, ri, ci, text):
    return set_tf(table.cell(ri, ci).text_frame, text, table.columns[ci].width / 914400 - 0.24)
def merged_text(cell):
    """자동으로 나뉜 줄을 다시 한 정보로 합친다('- ' 항목과 온점 끝 줄은 유지)"""
    out = []
    for ln in [pa.text for pa in cell.text_frame.paragraphs]:
        s = ln.strip()
        if out and not s.startswith('- ') and not out[-1].rstrip().endswith('.'): out[-1] = out[-1].rstrip() + ' ' + s
        else: out.append(s)
    return '\n'.join(out)
def rewrap_table(table):
    for ri in range(len(table.rows)):
        for ci in range(len(table.columns)): set_cell(table, ri, ci, merged_text(table.cell(ri, ci)))
def fit_rows(table, rows=None, minh=0.36):
    for ri, row in enumerate(table.rows):
        if rows is not None and ri not in rows: continue
        mx = 0
        for c in row.cells:
            s, _ = run_props(c.text_frame); mx = max(mx, len(c.text_frame.paragraphs) * s * 1.3 / 72)
        row.height = Inches(max(minh, mx + 0.14))
def set_size(table, size):
    for e in list(table._tbl.iter(qn('a:rPr'))) + list(table._tbl.iter(qn('a:endParaRPr'))): e.set('sz', str(int(size * 100)))
def del_row(table, ri): table._tbl.remove(table._tbl.tr_lst[ri])
def add_row(table): table._tbl.append(copy.deepcopy(table._tbl.tr_lst[-1]))
def add_col(table, src=1, at=1):
    grid = table._tbl.find(qn('a:tblGrid')); grid.insert(at, copy.deepcopy(grid.findall(qn('a:gridCol'))[src]))
    for tr in table._tbl.tr_lst:
        tcs = tr.findall(qn('a:tc')); tr.insert(list(tr).index(tcs[at]), copy.deepcopy(tcs[src]))
def del_col(table, ci):
    grid = table._tbl.find(qn('a:tblGrid')); grid.remove(grid.findall(qn('a:gridCol'))[ci])
    for tr in table._tbl.tr_lst: tr.remove(tr.findall(qn('a:tc'))[ci])

# ── 도형 (정보 도형만, 1-1)
INK = RGBColor(0x2E, 0x24, 0x1D); MUTED = RGBColor(0x7A, 0x6A, 0x5E); BROWN = RGBColor(0x3B, 0x2A, 0x20)
BEIGE = RGBColor(0xF3, 0xEE, 0xE8); LINE = RGBColor(0xC9, 0xBB, 0xAA); FC_RED = RGBColor(0xD3, 0x2F, 0x2F)
FC_PINK = RGBColor(0xFD, 0xEC, 0xEC); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
def _font(r, size, bold, color):
    r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    rp = r._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        e = rp.find(qn(tag))
        if e is None: e = rp.makeelement(qn(tag), {}); rp.append(e)
        e.set('typeface', 'Malgun Gothic')
def textbox(s, x, y, w, h, lines, size=11, color=INK, bold=False, align=PP_ALIGN.LEFT, name='본문'):
    out = []
    for ln in lines: out += wrap(ln, w, size, bold)
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    for m in ('margin_left', 'margin_right', 'margin_top', 'margin_bottom'): setattr(tf, m, 0)
    for i, ln in enumerate(out):
        pa = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); pa.alignment = align
        _font(pa.add_run(), size, bold, color); pa.runs[0].text = ln
    tb.height = Inches(len(out) * size * 1.3 / 72 + 0.04)
    return tb
def box(s, x, y, w, h, lines, fill=BEIGE, colors=None, sizes=None, bolds=None, red_border=False, name='도식 상자'):
    b = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)); b.name = name
    b.fill.solid(); b.fill.fore_color.rgb = fill
    b.line.color.rgb = FC_RED if red_border else LINE; b.line.width = Pt(1.25 if red_border else 0.75)
    tf = b.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.06); tf.margin_top = tf.margin_bottom = Inches(0.02)
    for i, ln in enumerate(lines):
        pa = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); pa.alignment = PP_ALIGN.CENTER
        r = pa.add_run(); r.text = ln
        _font(r, (sizes or [11.5] + [10] * 9)[i], (bolds or [True] + [False] * 9)[i], (colors or [INK] + [MUTED] * 9)[i])
    return b
def arrow(s, x1, y1, x2, y2, color=MUTED, name='도식 화살표'):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)); c.name = name
    c.line.color.rgb = color; c.line.width = Pt(1.25)
    ln = c.line._get_or_add_ln(); ln.append(ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
    return c

# ── 장
def duplicate_slide(p, src_index, insert_at):
    """src_index 장을 복제해 insert_at(0부터) 위치에 넣는다. 노트는 노트 틀이 있을 때만 복사"""
    src = p.slides[src_index]; new = p.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes): sh._element.getparent().remove(sh._element)
    for el in src.shapes._spTree.iterchildren():
        if el.tag.split('}')[1] in ('sp', 'graphicFrame', 'cxnSp', 'grpSp', 'pic'): new.shapes._spTree.append(copy.deepcopy(el))
    ntf = new.notes_slide.notes_text_frame
    if ntf is not None and src.has_notes_slide: ntf.text = src.notes_slide.notes_text_frame.text
    ids = p.slides._sldIdLst; last = list(ids)[-1]; ids.remove(last); ids.insert(insert_at, last)
    return p.slides[insert_at]
def _runs(s):
    for sh in s.shapes:
        if sh.has_text_frame:
            for pa in sh.text_frame.paragraphs:
                for r in pa.runs: yield sh, r
        if sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    for pa in c.text_frame.paragraphs:
                        for r in pa.runs: yield sh, r
def renumber(p, new_no):
    """new_no(1부터) 장을 새로 넣은 뒤: 쪽 번호 n / N, 간지 목록 번호, 본문 'N장' · 'N~M장' · '(N~M)' 참조를 다시 맞춘다"""
    N = len(p.slides); PG = re.compile(r'^\d+ / \d+$')
    bump = lambda n: n + 1 if n >= new_no else n
    RNG = re.compile(r'(\d{1,2})~(\d{1,2})(?=\)|장)'); ONE = re.compile(r'(?<![\d.~])(\d{1,2})(?=장)')
    for i, s in enumerate(p.slides, 1):
        for sh, r in _runs(s):
            t = r.text
            if PG.match(t.strip()): r.text = f'{i} / {N}'; continue
            if sh.name == '간지 목록':
                m = re.match(r'^(\d+)(\s+)(.*)$', t)
                if m and i != new_no: r.text = f'{bump(int(m.group(1)))}{m.group(2)}{m.group(3)}'
                continue
            t2 = RNG.sub(lambda m: f'{bump(int(m.group(1)))}~{int(m.group(2)) + 1 if int(m.group(2)) >= new_no - 1 else int(m.group(2))}', t)
            t2 = ONE.sub(lambda m: str(bump(int(m.group(1)))), t2)
            if t2 != t: r.text = t2
def add_source_line(p, text, skip=(), pattern=r'[①②③④⑤⑥⑦⑧⑨]'):
    """번호를 쓰는 장 하단(y 7.02)에 출처 1줄. skip: 표지 · 간지 · 번호 원본 장 번호"""
    pat = re.compile(pattern)
    for i, s in enumerate(p.slides, 1):
        if i in skip or '간지 목록' in [x.name for x in s.shapes]: continue
        if pat.search(' '.join(r.text for _, r in _runs(s))) and not find(s, '출처'):
            textbox(s, 0.6, 7.02, 10.9, 0.26, [text], size=9, color=MUTED, name='출처')
def replace_picture(slide, png_path):
    """그림 위치 · 크기는 그대로 두고 이미지만 교체(같은 비율로 자른 이미지일 것)"""
    pic = [x for x in slide.shapes if x.shape_type == 13][0]
    pic.part.related_part(pic._element.blipFill.blip.rEmbed)._blob = open(png_path, 'rb').read()
    return pic
