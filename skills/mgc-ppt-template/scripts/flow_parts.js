// flow_parts.js — 흐름도 부품 (mgc-ppt-template v5.0, references/flowchart.md)
// 사용: const { N, edge, legend, note, g, CX } = require('./flow_parts')(D);   // D = require('./slim').deck('제목')
// 좌표 · 색 · 글자 크기는 고치지 않고 N() · edge() 인자로만 내용을 넣는다.
module.exports = function flowParts(D) {
  const { pres, T, C } = D;
  const NW = 2.1, DH = 0.85, PH = 0.66, FS = 11;
  const CX = [0, 1, 2, 3, 4].map(i => 0.6 + NW / 2 + i * ((12.13 - NW) / 4));   // 5열 중심 x, 좌우 끝 0.6 · 12.73
  const g = (a, b) => (CX[a] + CX[b]) / 2;
  function node(s, n) {
    const w = n.w || NW, h = n.h || (n.k === 'dec' ? DH : PH);
    const x = n.cx - w / 2, y = n.cy - h / 2;
    Object.assign(n, { w, h, x, y });
    const S = pres.shapes;
    if (n.k === 'end') s.addShape(S.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: C.dark }, line: { color: C.dark, width: 0.75 } });
    if (n.k === 'proc') s.addShape(S.RECTANGLE, { x, y, w, h, fill: { color: C.soft }, line: { color: C.soft, width: 0.75 } });
    if (n.k === 'dec') s.addShape(S.DIAMOND, { x, y, w, h, fill: { color: C.white }, line: { color: C.dark, width: 1 } });
    if (n.k === 'scr') s.addShape(S.RECTANGLE, { x, y, w, h, fill: { color: C.white }, line: { color: C.muted3, width: 1 } });
    const inset = n.k === 'dec' ? { w: w * 0.66, h: h * 0.72 } : { w: w - 0.16, h: h - 0.06 };
    T(s, n.t, { x: n.cx - inset.w / 2, y: n.cy - inset.h / 2, w: inset.w, h: inset.h, fontSize: n.fs || FS,
      bold: n.k === 'dec' || n.k === 'end', color: n.k === 'end' ? C.white : C.text, align: 'center' }, '흐름 ' + n.k);
    return n;
  }
  const anchor = (n, side) => !n.w ? n : ({ r: { x: n.x + n.w, y: n.cy }, l: { x: n.x, y: n.cy }, t: { x: n.cx, y: n.y }, b: { x: n.cx, y: n.y + n.h } })[side];
  function seg(s, a, b, arrow) {
    const x = Math.min(a.x, b.x), y = Math.min(a.y, b.y), w = Math.abs(b.x - a.x), h = Math.abs(b.y - a.y);
    if (w < 0.001 && h < 0.001) return;
    s.addShape(pres.shapes.LINE, { x, y, w, h, flipH: b.x < a.x, flipV: b.y < a.y, line: { color: C.muted, width: 1, endArrowType: arrow ? 'triangle' : undefined } });
  }
  function edge(s, a, sa, b, sb, { lab, via, lx = 0, ly = 0, lw = 0.7 } = {}) {
    const p1 = anchor(a, sa), p2 = anchor(b, sb), H = v => v === 'r' || v === 'l';
    let pts;
    if (via) pts = [p1, ...via, p2];
    else if (H(sa) && (H(sb) || !b.w)) pts = Math.abs(p1.y - p2.y) < 0.01 ? [p1, p2] : [p1, { x: (p1.x + p2.x) / 2, y: p1.y }, { x: (p1.x + p2.x) / 2, y: p2.y }, p2];
    else if (!H(sa) && !H(sb)) pts = Math.abs(p1.x - p2.x) < 0.01 ? [p1, p2] : [p1, { x: p1.x, y: (p1.y + p2.y) / 2 }, { x: p2.x, y: (p1.y + p2.y) / 2 }, p2];
    else if (H(sa)) pts = [p1, { x: p2.x, y: p1.y }, p2];
    else pts = [p1, { x: p1.x, y: p2.y }, p2];
    for (let i = 0; i < pts.length - 1; i++) seg(s, pts[i], pts[i + 1], i === pts.length - 2);
    if (lab) {
      const o = { r: [0.05, -0.27], l: [-lw - 0.05, -0.27], b: [0.07, 0.02], t: [0.07, -0.27] }[sa];
      T(s, lab, { x: p1.x + o[0] + lx, y: p1.y + o[1] + ly, w: lw, h: 0.25, fontSize: 10, bold: true, color: C.muted, align: sa === 'l' ? 'right' : 'left' }, '분기 이름');
    }
  }
  // 범례: 오른쪽 끝을 본문 끝선(12.73)에 맞춤
  function legend(s) {
    const items = [['end', '시작 · 끝', 0.75], ['proc', '앱 · 서버 · B.O 처리', 1.55], ['dec', '판단', 0.4], ['scr', '손님 · 운영자 안내', 1.35]];
    const total = items.reduce((a, it) => a + 0.38 + it[2], 0) + 0.2 * (items.length - 1);
    let x = 12.73 - total;
    const S = pres.shapes;
    items.forEach(([k, t, tw]) => {
      const o = { x, y: 0.45, w: 0.3, h: 0.18 };
      if (k === 'end') s.addShape(S.ROUNDED_RECTANGLE, Object.assign(o, { rectRadius: 0.05, fill: { color: C.dark }, line: { color: C.dark, width: 0.75 } }));
      if (k === 'proc') s.addShape(S.RECTANGLE, Object.assign(o, { fill: { color: C.soft }, line: { color: C.soft, width: 0.75 } }));
      if (k === 'dec') s.addShape(S.DIAMOND, Object.assign(o, { fill: { color: C.white }, line: { color: C.dark, width: 1 } }));
      if (k === 'scr') s.addShape(S.RECTANGLE, Object.assign(o, { fill: { color: C.white }, line: { color: C.muted3, width: 1 } }));
      T(s, t, { x: x + 0.38, y: 0.42, w: tw, h: 0.24, fontSize: 10, color: C.muted }, '범례');
      x += 0.38 + tw + 0.2;
    });
  }
  const N = (s, k, c, cy, t, o = {}) => node(s, Object.assign({ k, cx: CX[c], cy, t }, o));
  const note = (s, t, o) => T(s, t, Object.assign({ h: 0.3, fontSize: 11, color: C.muted, valign: 'top' }, o), '보조 설명');   // 항목이 여럿이면 줄마다 '- '
  return { node, edge, legend, note, N, g, CX, anchor };
};
// 상수: 상자 폭 NW 2.1, 판단 높이 0.85, 처리 높이 0.66, 글자 11pt — 5열 중심 x는 좌우 끝 0.6 · 12.73 기준
