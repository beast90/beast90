"""기준 덱과 새 덱의 도형 좌표·색·글자 서식을 비교한다.
사용: python3 compare_layout.py 기준.pptx 새덱.pptx
출력: 장별 불일치 도형 목록 + 총 불일치 수 (0이면 레이아웃 동일)"""
import sys
from pptx import Presentation
E = 914400
def sig(path):
    out = []
    for s in Presentation(path).slides:
        items = []
        for sh in s.shapes:
            g = tuple(round(v / E, 2) for v in (sh.left, sh.top, sh.width, sh.height))
            fill = ''
            try:
                if sh.fill.type == 1: fill = str(sh.fill.fore_color.rgb)
            except Exception: pass
            fmt = ''
            if sh.has_text_frame and sh.text_frame.text:
                r = [r for pa in sh.text_frame.paragraphs for r in pa.runs][0]
                try: c = str(r.font.color.rgb)
                except Exception: c = ''
                fmt = f"{r.font.size.pt if r.font.size else ''}/{r.font.bold}/{c}"
            items.append((str(sh.shape_type).split(' ')[0], g, fill, fmt))
        out.append(items)
    return out
a, b = sig(sys.argv[1]), sig(sys.argv[2])
bad = 0
for i, (x, y) in enumerate(zip(a, b), 1):
    d1, d2 = set(x) - set(y), set(y) - set(x)
    bad += len(d1)
    if d1 or d2:
        print(f'--- {i}장: 기준에만 {len(d1)}개, 새 덱에만 {len(d2)}개')
        for t in sorted(d1, key=lambda t: t[1]): print('  기준', t)
        for t in sorted(d2, key=lambda t: t[1]): print('  새덱', t)
print('장 수', len(a), len(b), '| 불일치 도형', bad)
