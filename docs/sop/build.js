// 產出 docs/BOM_Matcher_v<VER>_操作SOP.pptx
//   node build.js
// 版面座標與 preview.js 共用 layout.js，改版面只要改那一支。

const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const { VER, DATE, THEME, SLIDES } = require('./content.js');
const L = require('./layout.js');

const OUT = path.resolve(__dirname, '..', `BOM_Matcher_v${VER}_操作SOP.pptx`);

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';            // 10" x 5.625"
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = 'ADATA Technology';
pres.title = `BOM Matcher v${VER} 操作 SOP`;

const C = pres.SchemeColor;
const K = THEME.colors;

// ── 版面（layout） ──────────────────────────────────────────────────────
pres.defineSlideMaster({
  title: 'COVER', background: { color: K.dk1 },
  objects: [],
});
pres.defineSlideMaster({
  title: 'DIVIDER', background: { color: K.accent1 },
  objects: [],
});
pres.defineSlideMaster({
  title: 'CONTENT', background: { color: 'FFFFFF' },
  objects: [
    { text: { text: `BOM Matcher v${VER} 操作 SOP`,
      options: { x: L.M, y: 5.17, w: 5, h: 0.3, fontSize: 9,
                 color: K.accent5, isTextBox: true, margin: 0 } } },
    { text: { text: DATE,
      options: { x: 7.4, y: 5.17, w: 1.5, h: 0.3, fontSize: 9, align: 'right',
                 color: K.accent5, isTextBox: true, margin: 0 } } },
  ],
  slideNumber: { x: 9.2, y: 5.17, w: 0.5, h: 0.3, fontSize: 9,
                 color: K.accent5, align: 'right' },
});

// ── 目錄（Word 大綱式，可點選跳頁）────────────────────────────────────
// 目錄插在封面之後 = 第 2 頁，所以原本 index>=1 的投影片頁碼要 +2。
const DECK = [SLIDES[0], { kind: 'toc', section: '封面', title: '目錄' },
              ...SLIDES.slice(1)];
const TOC_SLIDE = 2;
const TOC = [];
DECK.forEach((s, i) => {
  const page = i + 1;
  if (s.kind === 'divider') {
    TOC.push({ lvl: 1, n: s.n, title: s.title, page });
  } else if (s.kind !== 'cover' && s.kind !== 'toc') {
    TOC.push({ lvl: 2, title: s.title, page });
  }
});

let sectionSeen = new Set();
function slideFor(s, masterName) {
  if (!sectionSeen.has(s.section)) {
    pres.addSection({ title: s.section });
    sectionSeen.add(s.section);
  }
  return pres.addSlide({ masterName, sectionTitle: s.section });
}

function title(sl, t, backLink = true) {
  sl.addText(t, { x: L.M, y: 0.3, w: 9, h: 0.55, fontSize: 28, bold: true,
                  color: K.dk1, isTextBox: true, margin: 0, valign: 'middle' });
  // 回目錄連結必須放在「投影片」上，不能放在版面上：
  // pptxgenjs 在 slideLayout 裡寫不出 hlinkClick 的關聯，會變成 rIdundefined，
  // PowerPoint 會判定檔案毀損。
  if (backLink) {
    sl.addText([{ text: '☰ 目錄', options: { hyperlink: { slide: TOC_SLIDE } } }],
      { x: 6.2, y: 5.17, w: 0.95, h: 0.3, fontSize: 9, align: 'right',
        color: K.accent1, isTextBox: true, margin: 0 });
  }
}

// 小圓點（沿用程式本身的色票當視覺母題）
function dot(sl, x, y, color, d = 0.13) {
  sl.addShape(pres.ShapeType.ellipse,
    { x, y, w: d, h: d, fill: { color }, line: { color, width: 0 } });
}

function card(sl, x, y, w, h, opts = {}) {
  sl.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: opts.fill || K.lt2 },
    line: { color: opts.line || K.accent6, width: 1 },
    objectName: opts.name,
  });
}

// ════════════════════════════════════════════════════════════════════════
for (const s of DECK) {
  if (s.kind === 'toc') {
    const sl = slideFor(s, 'CONTENT');
    title(sl, '目錄', false);
    // 依「高度」分兩欄（不是依筆數），並在章節標題處斷欄；
    // 欄高超過可用空間時等比壓縮行距 —— 頁數一多，固定行距會超出版面。
    const step = e => (e.lvl === 1 ? 0.36 : 0.27);
    const total = TOC.reduce((a, e) => a + step(e), 0);
    let acc = 0, cut = TOC.length;
    for (let i = 0; i < TOC.length; i++) {
      if (acc >= total / 2 && TOC[i].lvl === 1) { cut = i; break; }
      acc += step(TOC[i]);
    }
    const AVAIL = 4.7 - 1.02;
    [TOC.slice(0, cut), TOC.slice(cut)].forEach((col, ci) => {
      const x = L.M + ci * 4.6;
      const colH = col.reduce((a, e) => a + step(e), 0);
      const f = Math.min(1, AVAIL / colH);
      let y = 1.02;
      col.forEach(e => {
        const lv1 = e.lvl === 1;
        const h = lv1 ? 0.3 : 0.26;              // 框高不縮，只縮行距（縮框會讓文字爆框）
        sl.addText([{ text: lv1 ? `${e.n}　${e.title}` : `　　${e.title}`,
          options: { bold: lv1, color: lv1 ? K.accent1 : K.dk2,
                     hyperlink: { slide: e.page } } }],
          { x, y, w: 3.75, h,
            fontSize: lv1 ? 12.5 : 10.5, isTextBox: true, margin: 0 });
        sl.addText([{ text: String(e.page),
          options: { hyperlink: { slide: e.page } } }],
          { x: x + 3.78, y, w: 0.34, h,
            fontSize: lv1 ? 11 : 10, align: 'right',
            color: lv1 ? K.accent1 : K.accent5, isTextBox: true, margin: 0 });
        y += step(e) * f;
      });
    });
    sl.addText('點選任一列可直接跳到該頁；每頁右下角的「☰ 目錄」可以回到這裡。',
      { x: L.M, y: 4.78, w: 9.0, h: 0.3, fontSize: 10, color: K.accent5,
        isTextBox: true, margin: 0 });
    sl.addNotes('目錄，含超連結。');
    continue;
  }

  if (s.kind === 'cover') {
    const sl = slideFor(s, 'COVER');
    sl.addText(s.title, { x: L.M, y: 1.68, w: 8.6, h: 1.15, fontSize: 54,
      bold: true, color: 'FFFFFF', isTextBox: true, margin: 0 });
    sl.addText(s.sub, { x: L.M, y: 2.72, w: 8.6, h: 0.45, fontSize: 22,
      color: K.accent1 === K.accent1 ? 'A5A0F5' : 'FFFFFF',
      isTextBox: true, margin: 0 });
    sl.addText(s.meta.join('\n'), { x: L.M, y: 4.25, w: 8.6, h: 0.75,
      fontSize: 12, color: '9AA0AE', lineSpacing: 20, isTextBox: true, margin: 0 });
    [K.accent4, K.accent2, K.accent3].forEach((c, i) =>
      dot(sl, L.M + i * 0.26, 1.33, c, 0.16));
    sl.addNotes('BOM Matcher 操作 SOP 封面。');
    continue;
  }

  if (s.kind === 'divider') {
    const sl = slideFor(s, 'DIVIDER');
    sl.addText(s.n, { x: L.M, y: 1.52, w: 2, h: 1.2, fontSize: 60, bold: true,
      color: '8D86F0', isTextBox: true, margin: 0 });
    sl.addText(s.title, { x: L.M, y: 2.75, w: 8.6, h: 0.75, fontSize: 36,
      bold: true, color: 'FFFFFF', isTextBox: true, margin: 0 });
    sl.addText('☰ 回目錄', { x: 8.1, y: 5.1, w: 1.4, h: 0.3, fontSize: 10,
      align: 'right', color: 'C9C5F8', isTextBox: true, margin: 0,
      hyperlink: { slide: TOC_SLIDE } });
    sl.addNotes(s.title);
    continue;
  }

  const sl = slideFor(s, 'CONTENT');
  title(sl, s.title);

  if (s.kind === 'flow') {
    // 卡片寬度依輸入個數平均分配（v0.10 起五個輸入）
    const y = 1.1, h = 1.25, gap = 0.19;
    const w = (9.0 - gap * (s.inputs.length - 1)) / s.inputs.length;
    s.inputs.forEach((it, i) => {
      const x = L.M + i * (w + gap);
      card(sl, x, y, w, h, { name: `in${i}` });
      dot(sl, x + 0.14, y + 0.17, K[it.c] || K.accent1);
      sl.addText(it.k, { x: x + 0.33, y: y + 0.1, w: w - 0.45, h: 0.28,
        fontSize: s.inputs.length > 4 ? 11 : 12, bold: true, color: K.dk1,
        isTextBox: true, margin: 0 });
      sl.addText(it.v, { x: x + 0.14, y: y + 0.45, w: w - 0.28, h: 0.65,
        fontSize: 10, color: K.accent5, lineSpacing: 14, isTextBox: true, margin: 0 });
    });
    sl.addShape(pres.ShapeType.triangle, { x: 4.78, y: 2.55, w: 0.44, h: 0.3,
      fill: { color: K.accent1 }, line: { width: 0 }, rotate: 180 });
    card(sl, L.M, 2.98, 9.0, 0.92, { fill: 'EEF0FF', line: 'C9CDF5', name: 'out' });
    sl.addText(s.out, { x: L.M + 0.22, y: 3.1, w: 8.6, h: 0.35, fontSize: 17,
      bold: true, color: K.accent1, isTextBox: true, margin: 0 });
    sl.addText(s.outNote, { x: L.M + 0.22, y: 3.47, w: 8.6, h: 0.3, fontSize: 11,
      color: K.dk2, isTextBox: true, margin: 0 });
    sl.addText(s.note, { x: L.M, y: 4.12, w: 9.0, h: 0.6, fontSize: 11,
      color: K.accent5, lineSpacing: 16, isTextBox: true, margin: 0 });
    sl.addNotes(s.note.replace(/\n/g, ' '));
    continue;
  }

  if (s.kind === 'cards') {
    const w = 2.93, gap = 0.2, y = 1.08;
    s.cards.forEach((c, i) => {
      const x = L.M + i * (w + gap);
      card(sl, x, y, w, 3.5, { name: `card${i}` });
      sl.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.78, rectRadius: 0.06,
        fill: { color: i === 0 ? K.accent5 : i === 1 ? K.accent1 : '2F8F5B' },
        line: { width: 0 } });
      sl.addText(c.t, { x: x + 0.16, y: y + 0.1, w: w - 0.32, h: 0.3,
        fontSize: 12, bold: true, color: 'FFFFFF', isTextBox: true, margin: 0 });
      sl.addText(c.s, { x: x + 0.16, y: y + 0.42, w: w - 0.32, h: 0.26,
        fontSize: 10, color: 'E6E7F5', isTextBox: true, margin: 0 });
      sl.addText(c.b.map((t, j) => ({ text: t,
        options: { bullet: true, breakLine: j !== c.b.length - 1 } })), {
        x: x + 0.18, y: y + 0.95, w: w - 0.36, h: 2.42, fontSize: 11,
        color: K.dk2, lineSpacing: 16, paraSpaceAfter: 7,
        isTextBox: true, margin: 0 });
    });
    sl.addNotes(s.cards.map(c => `${c.t}：${c.s}`).join('；'));
    continue;
  }

  if (s.kind === 'compare') {
    const w = 4.4, y = 1.08, h = 2.85;
    [[s.left, L.M, K.accent1], [s.right, L.M + w + 0.2, K.accent5]]
      .forEach(([c, x, col], i) => {
        card(sl, x, y, w, h, { name: `cmp${i}` });
        sl.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.8, rectRadius: 0.06,
          fill: { color: col }, line: { width: 0 } });
        sl.addText(c.t, { x: x + 0.2, y: y + 0.11, w: w - 0.4, h: 0.32,
          fontSize: 15, bold: true, color: 'FFFFFF', isTextBox: true, margin: 0 });
        sl.addText(c.s, { x: x + 0.2, y: y + 0.45, w: w - 0.4, h: 0.26,
          fontSize: 11, color: 'E6E7F5', isTextBox: true, margin: 0 });
        sl.addText(c.b.map((t, j) => ({ text: t,
          options: { bullet: true, breakLine: j !== c.b.length - 1 } })), {
          x: x + 0.22, y: y + 0.98, w: w - 0.44, h: 1.75, fontSize: 11.5,
          color: K.dk2, lineSpacing: 17, paraSpaceAfter: 8,
          isTextBox: true, margin: 0 });
      });
    sl.addText(s.note, { x: L.M, y: 4.18, w: 9.0, h: 0.4, fontSize: 11.5,
      color: K.accent5, isTextBox: true, margin: 0 });
    sl.addNotes(s.note);
    continue;
  }

  if (s.kind === 'steps') {
    let y = 1.08;
    s.steps.forEach((st, i) => {
      sl.addShape(pres.ShapeType.ellipse, { x: L.M, y: y + 0.04, w: 0.42, h: 0.42,
        fill: { color: K.accent1 }, line: { width: 0 } });
      sl.addText(st.n, { x: L.M, y: y + 0.04, w: 0.42, h: 0.42, fontSize: 15,
        bold: true, color: 'FFFFFF', align: 'center', valign: 'middle',
        isTextBox: true, margin: 0 });
      sl.addText(st.t, { x: L.M + 0.62, y: y, w: 8.3, h: 0.3, fontSize: 15,
        bold: true, color: K.dk1, isTextBox: true, margin: 0 });
      sl.addText(st.d, { x: L.M + 0.62, y: y + 0.32, w: 8.3, h: 0.56,
        fontSize: 11.5, color: K.dk2, lineSpacing: 16, isTextBox: true, margin: 0 });
      y += 1.02;
    });
    sl.addNotes(s.steps.map(x => `${x.n}. ${x.t}`).join('　'));
    continue;
  }

  if (s.kind === 'shot') {
    const iw = 5.95, ih = iw / s.ar;
    sl.addImage({ path: path.resolve(__dirname, s.img), x: L.M, y: 1.05,
      w: iw, h: ih, rounding: false, objectName: 'shot' });
    sl.addShape(pres.ShapeType.rect, { x: L.M, y: 1.05, w: iw, h: ih,
      fill: { color: 'FFFFFF', transparency: 100 },
      line: { color: K.accent6, width: 1 } });

    // v0.10：方案 A —— 截圖上疊紅色編號圓點，右側說明用同一組號碼
    // marks 的 x/y 是「截圖原始像素座標」，這裡換算成吋。
    const SHOTW = s.shotw || 1280;
    const sc = iw / SHOTW, D = 0.185;
    (s.marks || []).forEach(([n, mx, my]) => {
      const cx = L.M + mx * sc - D / 2, cy = 1.05 + my * sc - D / 2;
      sl.addShape(pres.ShapeType.ellipse, { x: cx, y: cy, w: D, h: D,
        fill: { color: K.mark }, line: { color: 'FFFFFF', width: 1.5 } });
      // 數字的文字框要比圓點大一圈，否則行高撐破框（QA 會報「文字爆框」）
      const TD = D * 2;
      sl.addText(String(n), { x: cx - (TD - D) / 2, y: cy - (TD - D) / 2,
        w: TD, h: TD, fontSize: 8.5,
        bold: true, color: 'FFFFFF', align: 'center', valign: 'middle',
        isTextBox: true, margin: 0 });
    });

    const CIRC = ['', '①', '②', '③', '④', '⑤',
                  '⑥', '⑦', '⑧', '⑨', '⑩'];
    const bx = L.M + iw + 0.3, bw = 9.5 - bx;
    sl.addText(s.bullets.map((t, j) => {
      const parts = String(t).split('**');
      const runs = parts.map((p, k) => ({ text: p, options: { bold: k % 2 === 1 } }));
      const num = s.marks ? [{ text: CIRC[j + 1] + ' ',
        options: { bold: true, color: K.mark, fontSize: 12 } }] : [];
      return num.concat(runs.map((r, k) => ({ text: r.text, options: {
        ...r.options,
        bullet: !s.marks && k === 0,
        breakLine: k === runs.length - 1 && j !== s.bullets.length - 1 } })));
    }).flat(), { x: bx, y: 1.05, w: bw, h: 2.6, fontSize: 11,
      color: K.dk2, lineSpacing: 16, paraSpaceAfter: 9,
      isTextBox: true, margin: 0 });
    if (s.warn) {
      const wy = Math.max(1.05 + 2.65, 3.6);
      card(sl, bx, wy, bw, 1.42, { fill: 'FFF6EC', line: 'F0D3AE', name: 'warn' });
      dot(sl, bx + 0.14, wy + 0.16, K.accent3);
      sl.addText('注意', { x: bx + 0.33, y: wy + 0.09, w: bw - 0.5, h: 0.25,
        fontSize: 10.5, bold: true, color: '6B4413', isTextBox: true, margin: 0 });
      sl.addText(s.warn, { x: bx + 0.14, y: wy + 0.38, w: bw - 0.28, h: 0.98,
        fontSize: 10, color: '6B4413', lineSpacing: 14, isTextBox: true, margin: 0 });
    }
    sl.addNotes(s.bullets.map(b => b.replace(/\*\*/g, '')).join('　'));
    continue;
  }

  if (s.kind === 'legend') {
    let y = 1.08;
    s.legend.forEach((g, i) => {
      card(sl, L.M, y, 4.4, 0.6, { fill: g.bg, line: g.c, name: `lg${i}` });
      sl.addText(g.t, { x: L.M + 0.18, y: y + 0.07, w: 0.9, h: 0.46,
        fontSize: 13, bold: true, color: K.dk1, valign: 'middle',
        isTextBox: true, margin: 0 });
      sl.addText(g.d, { x: L.M + 1.05, y: y + 0.07, w: 3.2, h: 0.46,
        fontSize: 11, color: K.dk2, valign: 'middle', isTextBox: true, margin: 0 });
      y += 0.72;
    });
    const x = L.M + 4.7, w = 4.8;
    card(sl, x, 1.08, w, 2.88, { name: 'split' });
    sl.addText(s.split.title, { x: x + 0.2, y: 1.2, w: w - 0.4, h: 0.3,
      fontSize: 14, bold: true, color: K.accent1, isTextBox: true, margin: 0 });
    sl.addText(s.split.before, { x: x + 0.2, y: 1.58, w: w - 0.4, h: 0.28,
      fontSize: 10.5, color: K.accent5, fontFace: 'Consolas',
      isTextBox: true, margin: 0 });
    sl.addText(s.split.after.join('\n'), { x: x + 0.2, y: 1.95, w: w - 0.4, h: 0.78,
      fontSize: 10.5, color: K.dk1, fontFace: 'Consolas', lineSpacing: 15,
      isTextBox: true, margin: 0 });
    sl.addText(s.split.note, { x: x + 0.2, y: 2.84, w: w - 0.4, h: 0.95,
      fontSize: 11, color: K.dk2, lineSpacing: 16, isTextBox: true, margin: 0 });
    sl.addNotes('輸出 Excel 的色底規則與替代料拆列。');
    continue;
  }

  if (s.kind === 'table') {
    const rows = [s.head.map(h => ({ text: h, options: {
      bold: true, color: 'FFFFFF', fill: { color: K.accent1 } } }))];
    s.rows.forEach((r, i) => rows.push(r.map((cvl, j) => ({ text: cvl, options: {
      color: j === 0 ? K.accent1 : K.dk2,
      bold: j === 0,
      fontFace: j === 0 ? 'Consolas' : undefined,
      fill: { color: i % 2 ? 'FFFFFF' : K.lt2 } } }))));
    const rowH = s.rowH || 0.42;
    sl.addTable(rows, { x: L.M, y: 1.08, w: 9.0, colW: s.colW || [1.35, 3.0, 4.65],
      fontSize: 10.5, border: { type: 'solid', color: K.accent6, pt: 1 },
      rowH, valign: 'middle', margin: 6 });
    // 說明文字放在表格下方（列數不固定，不能寫死 y）
    const noteY = Math.max(3.55, 1.08 + rowH * (s.rows.length + 1) + 0.15);
    sl.addText(s.note, { x: L.M, y: noteY, w: 9.0, h: 0.7, fontSize: 11,
      color: K.accent5, lineSpacing: 16, isTextBox: true, margin: 0 });
    sl.addNotes(s.note.replace(/\n/g, ' '));
    continue;
  }

  if (s.kind === 'qa') {
    const w = 4.4, gap = 0.2;
    s.qa.forEach((it, i) => {
      const x = L.M + (i % 2) * (w + gap);
      const y = 1.08 + Math.floor(i / 2) * 1.25;
      card(sl, x, y, w, 1.1, { name: `qa${i}` });
      sl.addText([{ text: 'Q　', options: { bold: true, color: K.accent1 } },
                  { text: it.q, options: { bold: true, color: K.dk1 } }],
        { x: x + 0.16, y: y + 0.1, w: w - 0.32, h: 0.3, fontSize: 12,
          isTextBox: true, margin: 0 });
      sl.addText(it.a, { x: x + 0.16, y: y + 0.44, w: w - 0.32, h: 0.58,
        fontSize: 10.5, color: K.dk2, lineSpacing: 15, isTextBox: true, margin: 0 });
    });
    sl.addNotes(s.qa.map(x => x.q).join('　'));
    continue;
  }

  if (s.kind === 'final') {
    let y = 1.08;
    s.points.forEach((p, i) => {
      dot(sl, L.M + 0.02, y + 0.11, [K.accent1, K.accent2, K.accent3,
        K.accent4, K.accent5][i % 5], 0.15);
      sl.addText(p[0], { x: L.M + 0.32, y: y, w: 3.1, h: 0.32, fontSize: 14,
        bold: true, color: K.dk1, isTextBox: true, margin: 0 });
      sl.addText(p[1], { x: L.M + 3.5, y: y + 0.02, w: 5.5, h: 0.32, fontSize: 11.5,
        color: K.dk2, isTextBox: true, margin: 0 });
      y += 0.62;
    });
    sl.addText(s.tail, { x: L.M, y: 4.35, w: 9.0, h: 0.3, fontSize: 11,
      color: K.accent5, isTextBox: true, margin: 0 });
    sl.addNotes(s.points.map(p => p[0]).join('　'));
    continue;
  }

  throw new Error('未知的 slide kind: ' + s.kind);
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  const themePath = path.resolve(__dirname, 'theme.tmp.json');
  fs.writeFileSync(themePath, JSON.stringify(THEME, null, 2), 'utf8');
  // 把目錄結構輸出給 preview.py 用（HTML 版 SOP 的目錄要跟 pptx 一致）
  fs.writeFileSync(path.resolve(__dirname, 'toc.json'),
                   JSON.stringify(TOC, null, 2), 'utf8');
  console.log('✓ 產出', OUT);
  console.log('  投影片', DECK.length, '頁（含目錄）');
})();
