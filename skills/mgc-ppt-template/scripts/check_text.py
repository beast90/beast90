"""글꼴 · 줄 탈락 · 문장 줄바꿈 · 기호 검사 (mgc-ppt-template v4.2). 사용: python3 check_text.py out.pptx
모든 pptx에 쓸 수 있다(기능 보고형, 안내문형, QA 수정 요청, 사용자가 손으로 고친 파일).
1) 글꼴: 테마와 모든 글자가 맑은고딕인가
2) 줄 탈락: 직접 넣은 줄(문단)마다 맑은고딕 추정 폭이 상자·표 칸 폭을 넘지 않는가
   (넘으면 PowerPoint가 자동으로 줄을 바꿔 글자가 다음 줄로 밀려 내려간다. 자간 spc는 반영)
3) 줄 수: 줄 수 × 글자 크기 × 1.3이 상자·행 높이를 넘지 않는가
4) 문장 줄바꿈: 한 줄 안에서 온점(.)으로 문장이 끝난 뒤 다음 문장이 이어지지 않는가
   (숫자 1.5, 약어 B.O, 날짜 2026.10.15는 해당 없음)
5) 기호(v4.2): 글머리 「• 」(→ 「- 」), 제목 · 영역 이름 구분 「 — 」(→ 괄호), 줄 머리 「·」. 줄 끝 「·」는 [참고]로만 알린다(문장 속 긴 나열 예외)
글자 폭은 fit_text.js와 같은 맑은고딕 추정표, 여유 3%. 문제 0건이면 '글꼴·줄바꿈 통과'"""
import sys, re, zipfile
from pptx import Presentation
from pptx.util import Emu

OK_FONTS = {'Malgun Gothic', '맑은 고딕'}
NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
SAFE, LINE, DEF_SIZE = 0.97, 1.3, 18.0
NOTE = []
SENT = re.compile(r"[가-힣A-Za-z)\]'\"’”%]\.[ \t]+\S")

def em(ch, bold):
    c = ord(ch); b = 1.06 if bold else 1.0
    if (0xAC00 <= c <= 0xD7A3 or 0x1100 <= c <= 0x11FF or 0x3130 <= c <= 0x318F or
            0x4E00 <= c <= 0x9FFF or 0x3000 <= c <= 0x303F or 0xFF00 <= c <= 0xFFEF): return 0.98 if bold else 0.96
    if ch == ' ': return 0.30
    if ch.isdigit(): return 0.56 * b
    if 'A' <= ch <= 'Z': return 0.64 * b
    if 'a' <= ch <= 'z': return 0.52 * b
    if ch in '.,:;\'"`|!‘’“”': return 0.30
    if ch in '()[]{}': return 0.34
    if ch == '·': return 0.30
    if ch == '•': return 0.62
    if ch == '\u00a0': return 0.30
    if ch in '→←↑↓↔': return 1.0
    if ch in '~+=<>#': return 0.62
    if ch == '%': return 0.85
    if ch in '/\\': return 0.40
    if ch == '-': return 0.38
    if ch == '–': return 0.6
    if ch in '—@': return 1.0
    if ch == '&': return 0.70
    if ch == '*': return 0.50
    return 0.92

def para_lines(p_el, def_size):
    """문단 하나를 <a:br> 기준 줄로 나눠 [(글자, 폭pt, 최대 글자 크기)]"""
    lines, cur, w, mx = [], '', 0.0, 0.0
    for el in p_el:
        tag = el.tag.replace(NS, '')
        if tag == 'br':
            lines.append((cur, w, mx)); cur, w, mx = '', 0.0, 0.0
        elif tag in ('r', 'fld'):
            rpr = el.find(NS + 'rPr'); t = ''.join(x.text or '' for x in el.iter(NS + 't'))
            sz = int(rpr.get('sz')) / 100 if rpr is not None and rpr.get('sz') else def_size
            bold = rpr is not None and rpr.get('b') in ('1', 'true')
            spc = int(rpr.get('spc')) / 100 if rpr is not None and rpr.get('spc') else 0.0
            cur += t; w += sum(em(ch, bold) * sz for ch in t) + spc * len(t); mx = max(mx, sz)
    if mx == 0:
        epr = p_el.find(NS + 'endParaRPr')
        mx = int(epr.get('sz')) / 100 if epr is not None and epr.get('sz') else def_size
    lines.append((cur, w, mx))
    return lines

def check_box(out, where, txBody, w_in, h_in, wrap=True):
    avail = w_in * 72 * SAFE
    n_lines, height = 0, 0.0
    for p_el in txBody.iter(NS + 'p'):
        for t, wd, sz in para_lines(p_el, DEF_SIZE):
            n_lines += 1; height += sz * LINE / 72
            if not t.strip(): continue
            if wd > avail:
                over = (wd - avail) / (sz * 0.96)
                out.append(f'{where} [줄 탈락] 폭 {wd/72:.2f}in > {w_in*SAFE:.2f}in (약 {max(1, round(over + 0.49))}자 초과): {t!r}')
            if SENT.search(t):
                out.append(f'{where} [문장 줄바꿈] 한 줄에 문장이 이어짐: {t!r}')
            st = t.strip()
            if st.startswith('•'):
                out.append(f'{where} [기호] 글머리는 「- 」: {t!r}')
            if ' — ' in st:
                out.append(f'{where} [기호] 「 — 」 구분 대신 괄호 · 「: 」: {t!r}')
            if st.startswith('·'):
                out.append(f'{where} [기호] 줄 머리 「·」: {t!r}')
            if st.endswith('·'):
                NOTE.append(f'{where} [참고] 줄 끝 「·」(문장 속 긴 나열인지 확인): {t!r}')
    if h_in is not None and height > h_in * 1.05 and n_lines > 1:
        out.append(f'{where} [줄 수 초과] {n_lines}줄 ({height:.2f}in) > 높이 {h_in:.2f}in')

def fonts_of(el):
    res = []
    for r in el.iter(NS + 'r'):
        t = ''.join(x.text or '' for x in r.iter(NS + 't'))
        if not t.strip(): continue
        rpr = r.find(NS + 'rPr')
        lat = rpr.find(NS + 'latin') if rpr is not None else None
        ea = rpr.find(NS + 'ea') if rpr is not None else None
        res.append((t, lat.get('typeface') if lat is not None else None, ea.get('typeface') if ea is not None else None))
    return res

def emu_in(v): return Emu(v).inches if v is not None else 0.0

path = sys.argv[1]
font_bad, wrap_bad = [], []
z = zipfile.ZipFile(path)
x = z.read('ppt/theme/theme1.xml').decode('utf8') if 'ppt/theme/theme1.xml' in z.namelist() else ''
for kind in ('majorFont', 'minorFont'):
    m = re.search(rf'<a:{kind}>\s*<a:latin typeface="([^"]*)"', x)
    if m and m.group(1) not in OK_FONTS: font_bad.append(f'테마 {kind} 글꼴 {m.group(1)!r}')
theme_ok = not font_bad

p = Presentation(path)
for i, s in enumerate(p.slides, 1):
    for sh in s.shapes:
        for t, lat, ea in fonts_of(sh._element):
            for f in (lat, ea):
                if (f is None and not theme_ok) or (f is not None and not f.startswith('+') and f not in OK_FONTS):
                    font_bad.append(f'{i}장 [{sh.name}] 글꼴 {f or "미지정(테마 글꼴)"}: {t[:20]!r}'); break
        if sh.has_text_frame and sh.text_frame.text.strip():
            tf = sh.text_frame
            body = tf._txBody
            bpr = body.find(NS + 'bodyPr')
            w = emu_in(sh.width) - emu_in(tf.margin_left) - emu_in(tf.margin_right)
            h = emu_in(sh.height) - emu_in(tf.margin_top) - emu_in(tf.margin_bottom)
            check_box(wrap_bad, f'{i}장 [{sh.name}]', body, w, h, (bpr is None or bpr.get('wrap') != 'none'))
        if getattr(sh, 'has_table', False):
            tb = sh.table
            for ri, row in enumerate(tb.rows):
                for ci, c in enumerate(row.cells):
                    if not c.text_frame.text.strip(): continue
                    w = emu_in(tb.columns[ci].width) - emu_in(c.margin_left) - emu_in(c.margin_right)
                    h = emu_in(row.height) - emu_in(c.margin_top) - emu_in(c.margin_bottom)
                    check_box(wrap_bad, f'{i}장 [{sh.name} {ri + 1}행 {ci + 1}열]', c._tc.txBody, w, h)

for m in font_bad + wrap_bad: print(m)
for m in NOTE: print(m)
print('글꼴·줄바꿈 통과' if not (font_bad or wrap_bad) else f'글꼴 {len(font_bad)}건, 줄바꿈 {len(wrap_bad)}건')
