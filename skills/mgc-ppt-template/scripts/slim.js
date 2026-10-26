// slim.js — 간결 기획서형 부품 (mgc-ppt-template v5.1). 상단 흐름 헤더 · 장 구분 칩 · 순번 원 없는 골격
// 색 · 글꼴 · 줄맞춤은 references/design-common.md 토큰과 fit_text.js를 그대로 사용한다. 좌표 · 색 · 글자 크기는 고치지 않고 함수 인자로만 내용을 넣는다.
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const { fit } = require('./fit_text');
const F = 'Malgun Gothic';
const C = { yellow: 'FFDE00', text: '2E241D', dark: '3B2A20', muted: '7A6A5D', muted3: 'A6968A',
  coverBg: '000000', onDarkSub: 'F2F2F2', onDarkMeta: 'D9D9D9', darkRule: '595959',   // 표지 · 간지 전용(references/design-common.md 「컬러 토큰」)
  soft: 'F3EEE8', tan: 'E2D9CF', beige: 'E6DDD4', white: 'FFFFFF',
  red: 'FF0000',      // 강조 빨강(기본 빨강 단일 값): 본문 속 핵심 정책 강조(빨간 굵은 글씨) 전용
  newCell: 'FDF9C9'   // MGC badge-bg 토큰: 표의 「신규 개발」 칸 배경 전용
};
// [[...]] 로 감싼 부분 = 빨간 굵은 글씨. fit()은 표시를 뺀 글자로 줄맞춤하고, 결과 줄바꿈 위치에 맞춰 강조 구간을 다시 입힌다
const strip = t => String(t).replace(/\[\[|\]\]/g, '');
function rich(marked, fitted, extra) {
  marked = String(marked);
  if (marked.indexOf('[[') < 0) return fitted;
  let plain = '', flags = [], on = false;
  for (let i = 0; i < marked.length; i++) {
    if (marked.startsWith('[[', i)) { on = true; i++; continue; }
    if (marked.startsWith(']]', i)) { on = false; i++; continue; }
    plain += marked[i]; flags.push(on);
  }
  const lines = [[]]; let p = 0;
  const push = (ch, red) => { const L = lines[lines.length - 1], last = L[L.length - 1]; if (last && last.red === red) last.t += ch; else L.push({ t: ch, red }); };
  for (const ch of fitted) {
    if (ch === '\n') { if (plain[p] === ' ' || plain[p] === '\n') p++; lines.push([]); continue; }
    const red = !!flags[p]; if (plain[p] === ch) p++; push(ch, red);
  }
  const runs = [];
  lines.forEach((L, li) => {
    if (!L.length) L.push({ t: ' ', red: false });
    L.forEach((g, gi) => {
      const o = Object.assign({}, extra || {}, g.red ? { bold: true, color: C.red } : {});
      if (gi === L.length - 1 && li < lines.length - 1) o.breakLine = true;
      runs.push({ text: g.t, options: o });
    });
  });
  return runs;
}
const NEW_CELL = ['신규 개발'];   // badge-bg 대상: references/writing.md 「강조」
const X0 = 0.6, W0 = 12.13;
const ratio = p => { const b = fs.readFileSync(p); return b.readUInt32BE(16) / b.readUInt32BE(20); };

function deck(title) {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; pres.title = title;
  pres.theme = { headFontFace: F, bodyFontFace: F };
  const pages = [];
  let hasCover = false;
  const T = (s, t, o, name) => {
    const r = fit(strip(t), { w: o.w, h: o.h, size: o.fontSize, bold: !!o.bold, name: name || '' });
    const opt = Object.assign({ fontFace: F, color: C.text, margin: 0, isTextBox: true, lang: 'ko-KR', valign: 'middle' }, name ? { objectName: name } : {}, o);
    if (r.charSpacing) opt.charSpacing = r.charSpacing;
    return s.addText(rich(t, r.text, r.charSpacing ? { charSpacing: r.charSpacing } : null), opt);
  };
  function cover({ kind, title, sub, meta }) {
    const s = pres.addSlide(); s.background = { color: C.coverBg }; hasCover = true;
    T(s, kind, { x: 0.9, y: 2.3, w: 11.5, h: 0.4, fontSize: 14, bold: true, color: C.yellow });
    T(s, title, { x: 0.9, y: 2.8, w: 11.5, h: 1.0, fontSize: 36, bold: true, color: C.white, valign: 'top' });
    T(s, sub, { x: 0.9, y: 3.9, w: 11.5, h: 0.5, fontSize: 16, color: C.onDarkSub, valign: 'top' });
    T(s, meta, { x: 0.9, y: 6.4, w: 11.5, h: 0.35, fontSize: 12, color: C.onDarkMeta });
  }
  // 골격: 영역 이름(작은 글씨) + 제목 + 핵심 문장. 장식 도형 없음.
  function slide({ sec, title, msg, notes }) {
    const s = pres.addSlide(); s.background = { color: C.white };
    T(s, sec, { x: X0, y: 0.38, w: 6, h: 0.3, fontSize: 11, bold: true, color: C.muted }, '영역 이름');
    T(s, title, { x: X0, y: 0.7, w: W0, h: 0.6, fontSize: 24, bold: true }, '제목');
    if (msg) T(s, msg, { x: X0, y: 1.32, w: W0, h: 0.34, fontSize: 13, color: C.muted }, '핵심 문장');
    if (notes) s.addNotes(notes);
    pages.push(s); return s;
  }
  function phones(s, items, { x = X0, y = 1.95, h = 4.2, gap = 0.4, arrows = false, r }) {
    const w = h * (r || ratio(items[0].path));
    items.forEach((it, i) => {
      s.addImage({ path: it.path, x, y, w, h, altText: it.alt || it.cap });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { type: 'none' }, line: { color: C.beige, width: 0.75 } });
      const cw = w + Math.min(gap, 0.6) - 0.05;
      T(s, it.cap, { x: x - (cw - w) / 2, y: y + h + 0.1, w: cw, h: 0.3, fontSize: 12, bold: true, align: 'center' }, '화면 캡션');
      if (it.sub) T(s, it.sub, { x: x - (cw - w) / 2, y: y + h + 0.42, w: cw, h: 0.28, fontSize: 10.5, color: C.muted, align: 'center' }, '화면 설명');
      if (arrows && i < items.length - 1) T(s, '→', { x: x + w, y: y + h / 2 - 0.2, w: gap, h: 0.4, fontSize: 16, color: C.muted, align: 'center' });
      x += w + gap;
    });
    return x - gap;
  }
  // 설명 목록: 굵은 제목 + 회색 설명 (번호 원 없음)
  function notes(s, items, { x, y, w, gap }) {
    items.forEach((it, i) => {
      const yy = y + i * gap;
      T(s, it[0], { x, y: yy, w, h: 0.36, fontSize: 14, bold: true }, '설명 제목');
      T(s, it[1], { x, y: yy + 0.4, w, h: gap - 0.5, fontSize: 12, color: C.muted, valign: 'top' }, '설명');
    });
  }
  // 핵심 한 줄: 배경색 없이 왼쪽 얇은 세로선 + 굵은 글씨(시선을 뺏지 않는 요약). [[ ]] 부분만 빨간 굵은 글씨
  function key(s, t, { x, y, w, h = 0.5, size = 13 }) {
    s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.08, w: 0.05, h: h - 0.16, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
    T(s, t, { x: x + 0.2, y, w: w - 0.2, h, fontSize: size, bold: true, color: C.text }, '핵심 문장 요약');
  }
  // 간지: 영역 시작 장. 어두운 배경 + 영역 이름 + 담은 장 목록
  function divider({ name, desc, items }) {
    const s = pres.addSlide(); s.background = { color: C.coverBg }; s._dark = true;
    T(s, name, { x: 0.9, y: 2.2, w: 11.5, h: 0.9, fontSize: 34, bold: true, color: C.white });
    T(s, desc, { x: 0.9, y: 3.15, w: 11.5, h: 0.45, fontSize: 15, color: C.onDarkSub });
    s.addShape(pres.shapes.LINE, { x: 0.9, y: 3.85, w: 11.5, h: 0, line: { color: C.darkRule, width: 0.75 } });
    (items || []).forEach((it, i) => T(s, it, { x: 0.9, y: 4.05 + i * 0.38, w: 11.5, h: 0.34, fontSize: 13, color: C.onDarkMeta }, '간지 목록'));
    pages.push(s); return s;
  }
  function cap(s, { path, x = X0, y, w, alt, caption }) {
    const h = w / ratio(path);
    s.addImage({ path, x, y, w, h, altText: alt });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { type: 'none' }, line: { color: C.beige, width: 0.75 } });
    if (caption) T(s, caption, { x, y: y + h + 0.1, w, h: 0.3, fontSize: 11, color: C.muted }, '캡처 캡션');
    return y + h;
  }
  const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: F, fontSize: 12, color: C.text, valign: 'middle', margin: [0.05, 0.12, 0.05, 0.12] }, o) });
  // 표: head 있으면 머리행(tan), 첫 열 굵게. 구분 칸이 '신규 개발' 등이면 badge-bg 배경
  function table(s, head, rows, { x = X0, y, colW, rowH = 0.48, headH = 0.42, firstBold = true, size = 12, align }) {
    const data = (head ? [head.map(t => cell(t, { bold: true, fill: { color: C.tan }, fontSize: size }))] : [])
      .concat(rows.map(r => r.map((t, j) => cell(t, { bold: firstBold && j === 0, fontSize: size, fill: firstBold && j === 0 && !head ? { color: C.soft } : undefined }))));
    const rh = data.map((_, i) => (head && i === 0) ? headH : (Array.isArray(rowH) ? rowH[i - (head ? 1 : 0)] : rowH));
    data.forEach((r, i) => r.forEach((c, j) => {
      const raw = c.text, isHead = head && i === 0;
      const f = fit(strip(raw), { w: colW[j] - 0.24, h: rh[i] - 0.1, size: c.options.fontSize, bold: !!c.options.bold, name: `표 ${i + 1}행 ${j + 1}열` });
      if (f.charSpacing) c.options.charSpacing = f.charSpacing;
      c.text = rich(raw, f.text, f.charSpacing ? { charSpacing: f.charSpacing } : null);
      if (!isHead && NEW_CELL.indexOf(strip(raw)) >= 0) { c.options.fill = { color: C.newCell }; c.options.bold = true; }
      if (align && align[j]) c.options.align = align[j];   // 숫자 열은 'right'
    }));
    s.addTable(data, { x, y, w: colW.reduce((a, v) => a + v, 0), colW, rowH: rh, border: { type: 'solid', pt: 0.75, color: C.beige }, objectName: '표' });
    return y + rh.reduce((a, v) => a + v, 0);
  }
  // 핵심 수치 칸: 큰 숫자 + 설명. hi: 숫자를 빨강으로 강조할 칸 번호
  function stats(s, items, { y, h = 1.25, hi }) {
    const n = items.length, gap = 0.2, w = (W0 - gap * (n - 1)) / n;
    items.forEach((it, i) => {
      const x = X0 + i * (w + gap), on = i === hi;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.soft }, line: { color: C.soft, width: 1 } });
      T(s, it[0], { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: 0.6, fontSize: 28, bold: true, color: on ? C.red : C.text }, '핵심 수치');
      T(s, it[1], { x: x + 0.25, y: y + 0.78, w: w - 0.5, h: 0.32, fontSize: 12, color: C.muted }, '수치 설명');
    });
  }
  // 쪽 번호 = PowerPoint 슬라이드 번호(표지 포함). 피드백의 'n p'와 화면 번호가 어긋나지 않게 한다
  async function save(f) {
    const off = hasCover ? 1 : 0, total = pages.length + off;
    pages.forEach((s, i) => T(s, `${i + 1 + off} / ${total}`, { x: 11.73, y: 7.0, w: 1.0, h: 0.28, fontSize: 9, color: s._dark ? C.onDarkMeta : C.muted3, align: 'right' }));
    return pres.writeFile({ fileName: f });
  }
  return { pres, T, cover, slide, divider, count: () => pages.length + (hasCover ? 1 : 0), phones, notes, key, cap, table, stats, save, C };
}
module.exports = { deck, X0, W0 };
