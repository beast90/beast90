// mgc_report.js — 메가MGC커피 '기능 보고형' PPT 공통 부품 (mgc-ppt-template v3.4)
// 기준 템플릿: 신규 오픈 매장 프로모션 고도화 3개 기능 중간 보고 개선본(2026-10-02)
// 모든 수치 단위는 inch, LAYOUT_WIDE(13.333 x 7.5) 기준. 값을 바꾸지 말고 함수 인자로만 내용을 넣는다.
// 글자가 들어가는 도형·표에는 역할 이름(objectName)을 붙인다. check_text.py가 이 이름으로 글자 수 한도를 검사한다.
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const { fit } = require('./fit_text');   // 줄 탈락 방지 · 문장 단위 줄바꿈

const F = 'Malgun Gothic';
const C = {
  yellow: 'FFDE00', black: '000000',
  text: '2E241D',      // doc-bg-900: 제목·본문
  dark: '3B2A20',      // doc-bg-800: 활성 단계, 칩 테두리, 설명 순번 원, 노랑 위 글씨
  muted: '7A6A5D',     // doc-muted-500
  muted3: 'A6968A',    // doc-muted-300
  soft: 'F3EEE8',      // doc-surface-soft: 박스·비활성 단계
  tan: 'E2D9CF',       // 표 머리행
  beige: 'E6DDD4',     // 테두리
  gold: 'D98C1F',      // 화면안 번호 참조 전용
  white: 'FFFFFF',
};
const M = { left: 0.5, width: 12.33 };   // 좌우 여백 0.5, 본문 폭 12.33

function pngRatio(path) {               // PNG만 사용 (w/h)
  const b = fs.readFileSync(path);
  if (b.readUInt32BE(0) !== 0x89504e47) throw new Error('PNG만 지원: ' + path);
  return b.readUInt32BE(16) / b.readUInt32BE(20);
}

function createDeck({ title, date, docType = '중간 보고', stages }) {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.title = title;
  pres.theme = { headFontFace: F, bodyFontFace: F };   // 테마 글꼴도 맑은고딕: PowerPoint에서 새로 추가하는 글상자가 Calibri로 들어가지 않게
  if (!Array.isArray(stages) || stages.length < 1 || stages.length > 4)
    throw new Error('단계는 1~4개만 지원한다(흐름 헤더 폭 한계). 5개 이상이면 시점이 같은 기능을 한 단계로 묶거나 덱을 나눈다.');
  const slides = [];   // 쪽 번호 대상(본문 장)
  // 글상자: 문장 끝 줄바꿈 + 폭 검사(넘치면 자간 최대 -0.5pt, 그래도 넘치면 오류)
  const T = (s, t, o, name) => {
    const r = fit(t, { w: o.w, h: o.h, size: o.fontSize, bold: !!o.bold, name: name || '' });
    const opt = Object.assign({ fontFace: F, color: C.text, margin: 0, isTextBox: true, lang: 'ko-KR' }, name ? { objectName: name } : {}, o);
    if (r.charSpacing) opt.charSpacing = r.charSpacing;
    return s.addText(r.text, opt);
  };

  // ── 슬라이드 골격: 흐름 헤더 + 장 구분 칩 + 제목 + 목적 문구 (+ 진행 현황) ──
  function addSlide({ stage, kind, title, purpose, status, notes }) {
    const s = pres.addSlide();
    s.background = { color: C.white };
    const h = 0.66, y = 0.2, n = stages.length, ov = 0.23;
    const w = (M.width + ov * (n - 1)) / n;            // 2단계 6.28 / 3단계 4.263 / 4단계 3.255
    stages.forEach((st, i) => {
      const x = M.left + i * (w - ov), on = i === stage;
      s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x, y, w, h,
        fill: { color: on ? C.dark : C.soft }, line: { color: C.white, width: 1.5 } });
      const tx = x + (i === 0 ? 0.25 : 0.5), tw = w - (i === 0 ? 0.6 : 0.85);
      const a = fit(st[0], { w: tw, size: 10.5, bold: true, name: '흐름 헤더 1줄' }), b = fit(st[1], { w: tw, size: 12.5, bold: true, name: '흐름 헤더 2줄' });
      if (a.text.includes('\n') || b.text.includes('\n')) throw new Error('[흐름 헤더] 단계 문구는 줄마다 한 문장(온점 없이)으로 씁니다: ' + st.join(' / '));
      s.addText([
        { text: a.text, options: Object.assign({ fontSize: 10.5, bold: true, color: on ? C.yellow : C.muted, breakLine: true }, a.charSpacing ? { charSpacing: a.charSpacing } : {}) },
        { text: b.text, options: Object.assign({ fontSize: 12.5, bold: on, color: on ? C.white : C.muted }, b.charSpacing ? { charSpacing: b.charSpacing } : {}) },
      ], { x: tx, y, w: tw, h, fontFace: F, valign: 'middle', margin: 0, isTextBox: true, lang: 'ko-KR', objectName: '흐름 헤더 글씨' });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.98, w: 1.25, h: 0.44,
      fill: { color: C.white }, line: { color: C.dark, width: 1.25 }, rectRadius: 0.08 });
    T(s, kind, { x: 0.5, y: 0.98, w: 1.25, h: 0.44, fontSize: 12, bold: true, color: C.dark, align: 'center', valign: 'middle' }, '장 구분 칩');
    T(s, title, { x: 1.95, y: 0.92, w: 10.9, h: 0.56, fontSize: 22, bold: true, valign: 'middle' }, '제목');
    T(s, purpose, { x: 0.5, y: 1.5, w: 12.33, h: 0.34, fontSize: 12.5, color: C.muted, valign: 'middle' }, '목적 문구');
    if (status) T(s, status, { x: 0.5, y: 6.95, w: 9.0, h: 0.3, fontSize: 10.5, color: C.muted, valign: 'middle' }, '진행 현황');
    if (notes) s.addNotes(notes);
    slides.push(s);
    return s;
  }


  // ── 표지 (선택): 단색 배경 + 글자만. 장식 도형(원·띠·사선·이미지) 금지, 슬라이드 밖으로 걸치는 요소 금지 ──
  function addCover({ docType, title, subtitle, meta, notes }) {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    T(s, docType, { x: 0.9, y: 2.3, w: 11.5, h: 0.4, fontSize: 14, bold: true, color: C.yellow, valign: 'middle' });
    T(s, title, { x: 0.9, y: 2.8, w: 11.5, h: 1.4, fontSize: 36, bold: true, color: C.white, valign: 'top' });
    if (subtitle) T(s, subtitle, { x: 0.9, y: 4.3, w: 11.5, h: 0.5, fontSize: 16, color: C.tan, valign: 'top' });
    if (meta) T(s, meta, { x: 0.9, y: 6.4, w: 11.5, h: 0.35, fontSize: 12, color: C.muted3, valign: 'middle' });
    if (notes) s.addNotes(notes);
    return s;   // 표지는 쪽 번호 대상에서 제외
  }

  // ── 영역 소제목 ── size 16(설명 영역) / 15(표·캡처 위)
  function label(s, text, { x, y, w = 6, size = 15 }) {
    T(s, text, { x, y, w, h: size >= 16 ? 0.4 : 0.35, fontSize: size, bold: true, valign: 'middle' }, '영역 소제목');
  }

  // ── 휴대폰 화면 나열: 같은 줄은 첫 이미지 비율로 폭 통일 ──
  // items: [{ path, caption, sub?, alt }], 반환: 오른쪽 끝 x
  function phones(s, items, { x = 0.5, y, h, gap, arrows = false, ratio }) {
    const w = h * (ratio || pngRatio(items[0].path));
    items.forEach((it, i) => {
      s.addImage({ path: it.path, x, y, w, h, altText: it.alt || it.caption });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { type: 'none' }, line: { color: C.beige, width: 0.75 } });
      T(s, it.caption, { x: x - 0.15, y: y + h + 0.1, w: w + 0.3, h: 0.3, fontSize: 12, bold: true, align: 'center', valign: 'middle' }, '화면 캡션');
      if (it.sub) T(s, it.sub, { x: x - 0.2, y: y + h + 0.4, w: w + 0.4, h: 0.5, fontSize: 10.5, color: C.muted, align: 'center', valign: 'top' }, '화면 설명');
      if (arrows && i < items.length - 1)
        T(s, '→', { x: x + w + 0.1, y: y + h / 2 - 0.2, w: 0.4, h: 0.4, fontSize: 16, color: C.muted, align: 'center', valign: 'middle' });
      x += w + gap;
    });
    return x - gap;
  }

  // ── 설명 순번 목록 (다크 브라운 원) ── items: [[제목, '설명1\n설명2'], ...]
  function points(s, items, { x = 7.6, y, w = 5.23, gap }) {
    items.forEach((it, i) => {
      const yy = y + i * gap;
      s.addShape(pres.shapes.OVAL, { x, y: yy, w: 0.4, h: 0.4, fill: { color: C.dark }, line: { color: C.dark, width: 1 } });
      T(s, String(i + 1), { x, y: yy, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'middle' });
      T(s, it[0], { x: x + 0.58, y: yy, w: w - 0.58, h: 0.36, fontSize: 14, bold: true, valign: 'middle' }, '설명 순번 제목');
      T(s, it[1], { x: x + 0.58, y: yy + 0.4, w: w - 0.58, h: gap - 0.5, fontSize: 12, color: C.muted, valign: 'top' }, '설명 순번 설명');
    });
  }

  // ── 가로 흐름 박스 (결과 단계 1칸만 노랑) ──
  function flowBoxes(s, labels, { y = 1.95, h = 0.85, key }) {
    const n = labels.length, ar = 0.3, w = (M.width - ar * (n - 1)) / n;
    labels.forEach((t, i) => {
      const x = M.left + i * (w + ar), k = i === key;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: k ? C.yellow : C.soft }, line: { color: k ? C.yellow : C.soft, width: 1 } });
      T(s, t, { x: x + 0.1, y, w: w - 0.2, h, fontSize: 12.5, bold: true, color: k ? C.dark : C.text, align: 'center', valign: 'middle' }, '흐름 박스');
      if (i < n - 1) T(s, '→', { x: x + w, y, w: ar, h, fontSize: 14, color: C.muted, align: 'center', valign: 'middle' });
    });
  }

  // ── 단계 카드 (운영자 작업 순서, 마지막/핵심 1장 노랑) ── items: [[제목, 설명], ...]
  function stepCards(s, items, { y = 2.35, h = 1.55, key }) {
    const n = items.length, ar = 0.3, w = (M.width - ar * (n - 1)) / n;
    items.forEach((it, i) => {
      const x = M.left + i * (w + ar), k = i === key;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: k ? C.yellow : C.soft }, line: { color: k ? C.yellow : C.soft, width: 1 } });
      s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.15, w: 0.4, h: 0.4, fill: { color: C.dark }, line: { color: C.dark, width: 1 } });
      T(s, String(i + 1), { x: x + 0.2, y: y + 0.15, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: k ? C.yellow : C.white, align: 'center', valign: 'middle' });
      T(s, it[0], { x: x + 0.75, y: y + 0.15, w: w - 0.9, h: 0.4, fontSize: 14, bold: true, color: k ? C.dark : C.text, valign: 'middle' }, '단계 카드 제목');
      T(s, it[1], { x: x + 0.2, y: y + 0.65, w: w - 0.4, h: 0.8, fontSize: 11.5, color: k ? C.dark : C.muted, valign: 'top' }, '단계 카드 설명');
      if (i < n - 1) T(s, '→', { x: x + w, y, w: ar, h, fontSize: 14, color: C.muted, align: 'center', valign: 'middle' });
    });
  }

  // ── 화면 캡처 1장 (B.O·표 화면) ── 반환: 캡처 하단 y
  function capture(s, { path, x = 0.5, y, w = 12.33, alt, caption, captionBold = false, ratio }) {
    const h = w / (ratio || pngRatio(path));
    s.addImage({ path, x, y, w, h, altText: alt });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { type: 'none' }, line: { color: C.beige, width: 0.75 } });
    if (caption) T(s, caption, captionBold
      ? { x, y: y + h + 0.1, w, h: 0.3, fontSize: 12, bold: true, valign: 'middle' }
      : { x, y: y + h + 0.1, w, h: 0.3, fontSize: 11.5, color: C.muted, valign: 'middle' }, '캡처 캡션');
    return y + h;
  }

  // ── 화면안 번호 설명 카드 (골드 원 = 화면 캡처 속 주황 번호와 동일 번호) ── items: [[번호, 제목, 설명], ...]
  function numberCards(s, items, { y, h = 1.75, footnote }) {
    const n = items.length, gap = 0.25, w = (M.width - gap * (n - 1)) / n;
    items.forEach((it, i) => {
      const x = M.left + i * (w + gap);
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.soft }, line: { color: C.soft, width: 1 } });
      s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.2, w: 0.4, h: 0.4, fill: { color: C.gold }, line: { color: C.gold, width: 1 } });
      T(s, String(it[0]), { x: x + 0.2, y: y + 0.2, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'middle' });
      T(s, it[1], { x: x + 0.2, y: y + 0.72, w: w - 0.4, h: 0.35, fontSize: 13.5, bold: true, valign: 'middle' }, '번호 카드 제목');
      T(s, it[2], { x: x + 0.2, y: y + 1.1, w: w - 0.4, h: 0.6, fontSize: 11.5, color: C.muted, valign: 'top' }, '번호 카드 설명');
    });
    // 각주는 하단 정보줄(6.95)과 겹치지 않게 폭 9.0 이내
    if (footnote) T(s, footnote, { x: 0.5, y: y + h + 0.2, w: 9.0, h: 0.3, fontSize: 11.5, color: C.muted, valign: 'middle' }, '각주');
  }

  const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: F, fontSize: 12, color: C.text, valign: 'middle', margin: [0.05, 0.12, 0.05, 0.12] }, o) });
  // 표 셀 맞추기: 열 폭(좌우 여백 0.12씩 제외)·행 높이(상하 0.05씩 제외)
  function fitCells(rows, colW, rowH, name) {
    rows.forEach((r, i) => r.forEach((c, j) => {
      const o = c.options, hh = Array.isArray(rowH) ? rowH[i] : rowH;
      const f = fit(c.text, { w: colW[j] - 0.24, h: hh - 0.1, size: o.fontSize, bold: !!o.bold, name: `${name} ${i + 1}행 ${j + 1}열` });
      c.text = f.text; if (f.charSpacing) o.charSpacing = f.charSpacing;
    }));
    return rows;
  }
  // ── 머리행 표 (상황별 안내 등): 첫 열 Bold ──
  function headTable(s, head, rows, { x, y, w, colW, rowH }) {
    const data = [head.map(t => cell(t, { bold: true, fill: { color: C.tan } }))]
      .concat(rows.map(r => r.map((t, j) => cell(t, { bold: j === 0 }))));
    fitCells(data, colW, rowH, '머리행 표');
    s.addTable(data, { x, y, w, h: rowH.reduce((a, v) => a + v, 0), colW, rowH, border: { type: 'solid', pt: 0.75, color: C.beige }, objectName: '머리행 표' });
  }
  // ── 라벨 표 (머리행 없음, 첫 열 soft 배경 Bold) ──
  function labelTable(s, rows, { x, y, w, colW, rowH = 0.46 }) {
    const data = fitCells(rows.map(r => [cell(r[0], { bold: true, fill: { color: C.soft } }), cell(r[1])]), colW, rowH, '라벨 표');
    s.addTable(data, { x, y, w, colW, rowH, border: { type: 'solid', pt: 0.75, color: C.beige }, objectName: '라벨 표' });
  }
  // ── 키-값 줄 목록 (설정 요약) ──
  function kvList(s, rows, { x, y, w, labelW = 1.4, step = 0.372 }) {
    rows.forEach((r, i) => {
      const yy = y + i * step;
      T(s, r[0], { x, y: yy, w: labelW, h: 0.36, fontSize: 12, bold: true, valign: 'middle' }, '키-값 라벨');
      T(s, r[1], { x: x + labelW, y: yy, w: w - labelW, h: 0.36, fontSize: 12, valign: 'middle' }, '키-값 값');
    });
  }

  // ── 저장: 모든 장 오른쪽 아래에 "날짜 문서종류 · n / N" ──
  async function save(fileName) {
    slides.forEach((s, i) => T(s, `${date} ${docType} · ${i + 1} / ${slides.length}`,
      { x: 9.83, y: 6.95, w: 3.0, h: 0.3, fontSize: 10, color: C.muted, align: 'right', valign: 'middle' }));
    return pres.writeFile({ fileName });
  }

  return { pres, C, addCover, addSlide, label, phones, points, flowBoxes, stepCards, capture, numberCards, headTable, labelTable, kvList, save };
}

module.exports = { createDeck, C, F, pngRatio };
