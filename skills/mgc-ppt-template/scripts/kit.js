// kit.js (mgc-ppt-template v5.1) — slim.js 위에 자동 줄나눔(의미 단위: 띄어쓰기 기준) · 표 행 높이 자동 계산
const { deck, X0, W0 } = require('./slim');
const { widthPt, breakSentences } = require('./fit_text');
const strip = t => String(t).replace(/\[\[|\]\]/g, '');
const LINE = 1.3, BOTTOM = 6.85;
const bad = m => { if (process.env.MGC_FIT === 'warn') console.warn(m); else throw new Error(m); };

// 칸 폭(in)에 맞춰 띄어쓰기 단위로 줄을 나눈다. '- ' 항목의 이어지는 줄은 두 칸 들여쓴다
function wrap(text, wIn, size, bold) {
  const avail = wIn * 72 * 0.97 - 1.5;
  const out = [];
  breakSentences(String(text)).split('\n').forEach(par => {
    const hy = par.startsWith('- ');
    const toks = par.split(' ');
    let cur = '';
    toks.forEach(tok => {
      const cand = cur === '' ? tok : cur + ' ' + tok;
      if (cur !== '' && widthPt(strip(cand), size, bold) > avail) {
        out.push(cur);
        cur = (hy ? '  ' : '') + tok;
      } else cur = cand;
    });
    out.push(cur);
  });
  return out.join('\n');
}
const nlines = t => String(t).split('\n').length;

function kit(title) {
  const D = deck(title);
  const { T, C } = D; const dkey = D.key;
  const SECS = []; let CUR = null; const PG = {};
  function sec(name, desc, look) {
    const d = D.divider({ name, desc });
    CUR = { name, desc, look, div: d, start: D.count(), items: [] };
    SECS.push(CUR);
  }
  function sl(o) {
    const s = D.slide(Object.assign({}, o, { notes: o.notes || ('근거: ' + (o.src || '앱카드 연동 통합본 v0.2') + '.') }));
    if (CUR) CUR.items.push([D.count(), o.t || o.title]);
    if (o.id) PG[o.id] = D.count();
    return s;
  }
  // 글상자: 자동 줄나눔 후 필요한 높이로 배치. 반환 = 아래 끝 y
  function tx(s, text, { x = X0, y, w = W0, size = 12, bold = false, color, align, name }) {
    const t = wrap(text, w, size, bold);
    const h = nlines(t) * size * LINE / 72 + 0.04;
    if (y + h > BOTTOM + 0.001) bad(`[본문 넘침] ${name || ''} y ${y.toFixed(2)} + h ${h.toFixed(2)} > ${BOTTOM}: ${strip(t).slice(0, 40)}`);
    T(s, t, { x, y, w, h, fontSize: size, bold, color: color || C.text, valign: 'top', align: align || 'left' }, name || '본문');
    return y + h;
  }
  // 굵은 제목 + 회색 설명 목록(번호 없음). 반환 = 아래 끝 y
  function nl(s, items, { x = X0, y = 1.95, w = W0, gap = 0.22, tsize = 14, dsize = 12 }) {
    let yy = y;
    items.forEach(([t, d]) => {
      yy = tx(s, t, { x, y: yy, w, size: tsize, bold: true, name: '설명 제목' }) + 0.04;
      if (d) yy = tx(s, d, { x, y: yy, w, size: dsize, color: C.muted, name: '설명' });
      yy += gap;
    });
    return yy - gap;
  }
  // 표: 칸마다 자동 줄나눔, 행 높이 = 가장 긴 칸 줄 수 기준
  // 여백이 남으면 글자를 키운다(최대 +2pt, 14pt까지). 표 아래 핵심 한 줄 · 각주 자리 0.8in는 남긴다
  function tb(s, head, rows, o) {
    const base = o.size || 11, lim = (o.maxB || BOTTOM) - (o.reserve != null ? o.reserve : 0.8);
    let best = base;
    if (o.grow !== false) {
      for (let z = base + 0.5; z <= Math.min(base + 2, 14) + 0.001; z += 0.5) {
        if (measure(head, rows, Object.assign({}, o, { size: z })) <= lim) best = z; else break;
      }
    }
    return tb0(s, head, rows, Object.assign({}, o, { size: best }));
  }
  function measure(head, rows, { y = 1.95, colW, size = 11, firstBold = true, minH = 0.36 }) {
    const wr = (t, j, b) => wrap(t, colW[j] - 0.24, size, b);
    const hh = r => Math.max(...r.map(nlines)) * size * LINE / 72 + 0.14;
    const H = head ? head.map((t, j) => wr(t, j, true)) : null;
    const R = rows.map(r => r.map((t, j) => wr(t, j, firstBold && j === 0)));
    return y + (H ? Math.max(0.38, hh(H)) : 0) + R.map(r => Math.max(minH, hh(r))).reduce((a, v) => a + v, 0);
  }
  function tb0(s, head, rows, { x = X0, y = 1.95, colW, size = 11, align, firstBold = true, maxB = BOTTOM, minH = 0.36 }) {
    const sum = colW.reduce((a, v) => a + v, 0);
    if (Math.abs(sum - W0) > 0.011 && x === X0) bad(`[표 폭] 열 폭 합 ${sum.toFixed(2)} ≠ ${W0}`);
    const wr = (t, j, b) => wrap(t, colW[j] - 0.24, size, b);
    const H = head ? head.map((t, j) => wr(t, j, true)) : null;
    const R = rows.map(r => r.map((t, j) => wr(t, j, firstBold && j === 0)));
    const hh = r => Math.max(...r.map(nlines)) * size * LINE / 72 + 0.14;
    const headH = H ? Math.max(0.38, hh(H)) : 0;
    const rowH = R.map(r => Math.max(minH, hh(r)));
    const bottom = y + headH + rowH.reduce((a, v) => a + v, 0);
    if (bottom > maxB + 0.001) bad(`[표 넘침] 아래 끝 ${bottom.toFixed(2)} > ${maxB} (${(rows[0] || [''])[0]} 외 ${rows.length}행)`);
    D.table(s, H, R, { x, y, colW, rowH, headH: headH || 0.42, size, align, firstBold });
    return bottom;
  }
  function key(s, text, { y, x = X0, w = W0, size = 15 }) {
    if (y + 0.46 > 6.9 + 0.001) bad(`[핵심 한 줄 넘침] y ${y.toFixed(2)} ${text.slice(0,20)}`);
    dkey(s, text, { x, y, w, h: 0.46, size });
  }
  function img(s, path, { x = X0, y = 1.95, w = W0, maxB = BOTTOM, alt }) {
    const fs = require('fs'); const b = fs.readFileSync(path);
    const r = b.readUInt32BE(16) / b.readUInt32BE(20);
    let h = w / r;
    if (y + h > maxB) { h = maxB - y; w = h * r; x = x + (W0 - w) / 2 * (x === X0 ? 1 : 0); }
    s.addImage({ path, x, y, w, h, altText: alt || '' });
    s.addShape(D.pres.shapes.RECTANGLE, { x, y, w, h, fill: { type: 'none' }, line: { color: C.beige, width: 0.75 } });
    return { x, y, w, h, b: y + h };
  }
  function finish() {
    SECS.forEach(sc => {
      const its = sc.items;
      sc.end = its.length ? its[its.length - 1][0] : sc.start;
      its.forEach((it, i) => {
        const col = i < 6 ? 0 : 1, row = i % 6;
        T(sc.div, `${it[0]}    ${it[1]}`, { x: 0.9 + col * 5.9, y: 4.05 + row * 0.38, w: 5.6, h: 0.34, fontSize: 13, color: C.onDarkMeta }, '간지 목록');
      });
      if (its.length > 12) throw new Error('간지 목록 12개 초과: ' + sc.name);
    });
  }
  return Object.assign(D, { sec, sl, tx, nl, tb, key, img, finish, SECS, PG, wrap });
}
module.exports = { kit, X0, W0 };
