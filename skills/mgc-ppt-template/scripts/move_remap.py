"""move_remap.py — 장 하나를 다른 위치로 옮기고 쪽 번호 · 간지 목록 · 본문 「N장」 · 「a~b」 범위 · 노트 · 목차 범위를 새 번호로 바꾼다 (mgc-ppt-template v5.3, references/edit-user-pptx.md 「열 · 행 · 장 바꾸기」)
사용: python3 move_remap.py in.pptx out.pptx 옮길장 새위치   (예: 78 5)
다른 영역으로 옮기면 간지 목록 항목도 새 영역 간지로 옮기고 두 간지 목록을 2열(6줄씩)로 다시 놓는다. 간지 자체는 옮기지 않는다
규격서 · 정의서 같은 다른 문서의 장 번호(「인증 모듈 2.1.2 6장」 「정의서 v0.2 12장」)는 바꾸지 않는다(DOC 제외 규칙)"""
import sys, re, copy
sys.path.insert(0, __import__('os').path.dirname(__file__))
from pptx import Presentation
from pptx.util import Inches
import pptx_edit as PE

SRC, OUT = sys.argv[1], sys.argv[2]
OLD_POS, NEW_POS = int(sys.argv[3]), int(sys.argv[4])
def f(n):
    if n == OLD_POS: return NEW_POS
    if NEW_POS <= n < OLD_POS: return n + 1
    if OLD_POS < n <= NEW_POS: return n - 1
    return n

p = Presentation(SRC)
def is_div(s): return any(sh.name == '간지 목록' for sh in s.shapes)
_old = list(p.slides)
if is_div(_old[OLD_POS - 1]): sys.exit('간지는 옮기지 않습니다: 영역 구성을 바꾸려면 간지 목록 · 목차를 손으로 고치세요')
def sec_of(slides, pos): return next((k for k in range(pos - 1, -1, -1) if is_div(slides[k]) and k != pos - 1), None)
OLD_DIV = _old[sec_of(_old, OLD_POS)] if sec_of(_old, OLD_POS) is not None else None
ids = p.slides._sldIdLst; lst = list(ids)
el = lst[OLD_POS - 1]; ids.remove(el); ids.insert(NEW_POS - 1, el)
N = len(p.slides)

PG = re.compile(r'^\d+ / \d+$')
LIST = re.compile(r'^(\d+)(\s+)(.*)$')
REF = re.compile(r'(?<![\d.~])(\d{1,2})((?:\s*·\s*\d{1,2})*)\s*장')
RNG = re.compile(r'(?<![\d.])(\d{1,2})~(\d{1,2})(?=\s*장|\))')
DOC = re.compile(r'(규격|정의서|모듈|간편결제|로그인|가이드|요약서|답변|v\d|\d\.\d|안내|화면설계|Flipbook|부록|별표)[^장]{0,6}$')
def is_doc(t, start): return bool(DOC.search(t[max(0, start - 14):start]))
def remap_text(t):
    t = RNG.sub(lambda m: m.group(0) if is_doc(t, m.start()) else f'{f(int(m.group(1)))}~{f(int(m.group(2)))}', t)
    def one(m):
        if is_doc(t, m.start()): return m.group(0)
        nums = [m.group(1)] + re.findall(r'\d{1,2}', m.group(2))
        seps = re.findall(r'\s*·\s*', m.group(2))
        out = str(f(int(nums[0])))
        for s, n in zip(seps, nums[1:]): out += s + str(f(int(n)))
        return out + m.group(0)[len(m.group(1) + m.group(2)):]
    return REF.sub(one, t)

def runs_all(s):
    for sh in s.shapes:
        if sh.has_text_frame:
            for pa in sh.text_frame.paragraphs:
                for r in pa.runs: yield sh, r
        if sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    for pa in c.text_frame.paragraphs:
                        for r in pa.runs: yield sh, r

for i, s in enumerate(p.slides, 1):
    for sh, r in runs_all(s):
        t = r.text
        if PG.match(t.strip()): r.text = f'{i} / {N}'; continue
        if sh.name == '간지 목록':
            m = LIST.match(t)
            if m: r.text = f'{f(int(m.group(1)))}{m.group(2)}{m.group(3)}'
            continue
        t2 = remap_text(t)
        if t2 != t: r.text = t2
    if s.has_notes_slide and s.notes_slide.notes_text_frame is not None:
        tf = s.notes_slide.notes_text_frame
        for pa in tf.paragraphs:
            for r in pa.runs:
                t2 = remap_text(r.text)
                if t2 != r.text: r.text = t2

# 다른 영역으로 옮겼으면 간지 목록 항목을 새 영역 간지로 옮긴다
_new = list(p.slides)
k = sec_of(_new, NEW_POS); NEW_DIV = _new[k] if k is not None else None
_lefts = [sh.left for s in p.slides if is_div(s) for sh in s.shapes if sh.name == '간지 목록']
_x0 = min(_lefts) if _lefts else Inches(0.9)
_col2 = [x for x in _lefts if x > _x0 + Inches(1)]
X2_DECK = max(set(_col2), key=_col2.count) if _col2 else _x0 + Inches(5.9)   # 덱에서 가장 많이 쓴 둘째 열 x
LAYOUT = {}   # 옮기기 전 간지별 배치(둘째 열 x, 첫 열 줄 수)를 기억해 사용자 배치를 유지
def remember(div):
    its = [sh for sh in div.shapes if sh.name == '간지 목록']
    c2 = [sh.left for sh in its if sh.left > _x0 + Inches(1)]
    LAYOUT[id(div)] = (c2[0] if c2 else X2_DECK, max(6, sum(1 for sh in its if sh.left <= _x0 + Inches(1))))
def relayout(div):
    items = sorted([sh for sh in div.shapes if sh.name == '간지 목록'], key=lambda sh: int(LIST.match(sh.text_frame.text).group(1)))
    if not items: return
    x2, rows = LAYOUT.get(id(div), (X2_DECK, 6))
    for i, sh in enumerate(items):
        sh.left = _x0 if i < rows else x2; sh.top = Inches(4.05 + 0.38 * (i if i < rows else i - rows))
    if len(items) > 12: print(f'[확인] 간지 목록 {len(items)}개: 12개를 넘음')
if OLD_DIV is not None and NEW_DIV is not None and OLD_DIV is not NEW_DIV:
    remember(OLD_DIV); remember(NEW_DIV)
    moved = [sh for sh in OLD_DIV.shapes if sh.name == '간지 목록' and LIST.match(sh.text_frame.text) and int(LIST.match(sh.text_frame.text).group(1)) == NEW_POS]
    for sh in moved: NEW_DIV.shapes._spTree.append(sh._element)
    relayout(OLD_DIV); relayout(NEW_DIV)
    print('[확인] 다른 영역으로 옮김: 목차 「담은 내용」 · 간지 설명 문구는 손으로 확인')

# 2장 목차의 「장」 열(장 글자 없음, 목차가 2장에 있을 때)
_toc = [sh for sh in p.slides[1].shapes if sh.has_table] if len(p.slides) > 1 else []
if _toc:
    t2 = _toc[0].table
    for ri in range(1, len(t2.rows)):
        c = t2.cell(ri, 1); m = re.match(r'^(\d+)~(\d+)$', c.text.strip())
        if m: PE.set_cell(t2, ri, 1, f'{f(int(m.group(1)))}~{f(int(m.group(2)))}')
    # 영역 범위를 실제 간지 위치로 다시 계산(옮긴 장이 영역 경계를 넘었을 때)
    divs = [i for i, s in enumerate(p.slides, 1) if is_div(s)]
    for ri in range(1, len(t2.rows)):
        if ri - 1 < len(divs):
            a = divs[ri - 1]; b = (divs[ri] - 1) if ri < len(divs) else len(p.slides)
            if re.match(r'^\d+~\d+$', t2.cell(ri, 1).text.strip()): PE.set_cell(t2, ri, 1, f'{a}~{b}')
p.save(OUT)
print('moved', N)
