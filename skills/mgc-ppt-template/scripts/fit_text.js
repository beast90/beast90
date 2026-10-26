// fit_text.js — 줄 탈락 방지 · 문장 단위 줄바꿈 (mgc-ppt-template v3.4)
// 1) 문장이 온점으로 끝나면 줄을 바꾼다(숫자 1.5, 약어 B.O, 날짜 2026.10.15는 제외)
// 2) 줄마다 맑은고딕 폭을 추정해 칸 폭을 넘으면 자간을 최대 -0.5pt까지 좁히고, 그래도 넘으면 오류로 멈춘다
//    (자동 줄바꿈에 맡기지 않는다. 모든 줄바꿈은 직접 넣은 \n 위치에서만 생긴다)
// 3) 줄 수 × 줄 높이가 상자 높이를 넘으면 오류
// 글자 폭은 맑은고딕 추정값(em). check_text.py와 같은 표를 쓴다. 3% 여유를 둔다.
const MAX_TIGHT = 0.5;           // 자간 최대 축소량(pt)
const SAFE = 0.97;               // 추정 오차 여유
const LINE = 1.3;                // 줄 높이 = 글자 크기 × 1.3

function em(ch, bold) {
  const c = ch.codePointAt(0);
  const b = bold ? 1.06 : 1;
  if ((c >= 0xAC00 && c <= 0xD7A3) || (c >= 0x1100 && c <= 0x11FF) || (c >= 0x3130 && c <= 0x318F) ||
      (c >= 0x4E00 && c <= 0x9FFF) || (c >= 0x3000 && c <= 0x303F) || (c >= 0xFF00 && c <= 0xFFEF)) return bold ? 0.98 : 0.96;
  if (ch === ' ') return 0.30;
  if (/[0-9]/.test(ch)) return 0.56 * b;
  if (/[A-Z]/.test(ch)) return 0.64 * b;
  if (/[a-z]/.test(ch)) return 0.52 * b;
  if (/[.,:;'"`|!]/.test(ch) || ch === '‘' || ch === '’' || ch === '“' || ch === '”') return 0.30;
  if (/[()[\]{}]/.test(ch)) return 0.34;
  if (ch === '·') return 0.30;
  if (/[→←↑↓↔]/.test(ch)) return 1.0;
  if (/[~+=<>#]/.test(ch)) return 0.62;
  if (ch === '%') return 0.85;
  if (/[/\\]/.test(ch)) return 0.40;
  if (ch === '-') return 0.38;
  if (ch === '–') return 0.6;
  if (ch === '—' || ch === '@') return 1.0;
  if (ch === '&') return 0.70;
  if (ch === '*') return 0.50;
  return 0.92;
}
// 한 줄 폭(pt). spc = 자간(pt, 음수면 좁힘)
function widthPt(line, size, bold, spc = 0) {
  const chars = [...line];
  return chars.reduce((a, ch) => a + em(ch, bold) * size, 0) + spc * Math.max(chars.length - 1, 0);
}
function breakSentences(t) {
  return String(t).replace(/([가-힣A-Za-z)\]'"’”%])\.[ \t]+(?=\S)/g, '$1.\n');
}
// 상자 하나 맞추기. 반환: { text, charSpacing } / 넘치면 오류(MGC_FIT=warn이면 경고만)
function fit(text, { w, h, size, bold = false, name = '', margin = 0 }) {
  const t = breakSentences(text);
  const lines = t.split('\n');
  const avail = (w - 2 * margin) * 72 * SAFE;
  let spc = 0, errs = [];
  for (const l of lines) {
    const n = [...l].length;
    const wd = widthPt(l, size, bold);
    if (wd <= avail) continue;
    const need = n > 1 ? (wd - avail) / (n - 1) : Infinity;
    if (need <= MAX_TIGHT) spc = Math.min(spc, -Math.ceil(need * 10) / 10);
    else {
      const over = Math.ceil((widthPt(l, size, bold, -MAX_TIGHT) - avail) / (size * 0.96));
      errs.push(`[줄 탈락] ${name} "${l}" — 칸 폭을 넘습니다. 약 ${Math.max(over, 1)}자 줄이거나, 의미 단위로 줄을 나누거나, 칸을 넓히세요`);
    }
  }
  if (h != null) {
    const need = lines.length * size * LINE / 72, have = h;
    if (need > have * 1.05) errs.push(`[줄 수 초과] ${name} ${lines.length}줄 — 상자 높이 ${h}in에 ${Math.max(Math.floor(have * 1.05 * 72 / (size * LINE)), 1)}줄까지 들어갑니다: "${t.replace(/\n/g, ' / ')}"`);
  }
  if (errs.length) {
    if (process.env.MGC_FIT === 'warn') errs.forEach(e => console.warn(e));
    else throw new Error(errs.join('\n'));
  }
  return { text: t, charSpacing: spc };
}
module.exports = { fit, breakSentences, widthPt, em, MAX_TIGHT, SAFE, LINE };
