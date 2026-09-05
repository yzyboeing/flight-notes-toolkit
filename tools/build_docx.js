/*  build_docx.js —— markdown 源 → 带目录的 A4 横版 Word
 *  用法：node build_docx.js <input.md> [output.docx]
 *  支持：# / ## / ### 标题、段落、- 项目符号、1. 编号、表格、```代码块```、> 引用、**加粗**、`代码`
 *  目录用 Word TOC 域，打开文档后按 F9 或「更新域」刷新页码。
 */
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, TableOfContents,
  WidthType, ShadingType, BorderStyle, AlignmentType, VerticalAlign, HeadingLevel,
  PageBreak, Footer, PageNumber, PageOrientation, Bookmark, PageReference
} = require('docx');
const fs = require('fs');

const SRC = process.argv[2] || 'flight_theory_notes_prompt_v5.md';
const OUT = process.argv[3] || 'prompt.docx';

// 字体：默认用 macOS 自带族，保证本机 Word 与 LibreOffice 排版一致。
// Windows 上跑可覆盖：DOC_FONT_CN="Microsoft YaHei" DOC_FONT_EN="Segoe UI" DOC_FONT_MONO=Consolas
const CN   = process.env.DOC_FONT_CN   || 'PingFang SC',
      EN   = process.env.DOC_FONT_EN   || 'Helvetica Neue',
      MONO = process.env.DOC_FONT_MONO || 'Menlo';
const GRAY = '595959', LINE = 'BFBFBF', HDR = 'D9D9D9', ALT = 'F7F7F7', CODE = 'F2F2F2';
const RED = 'C00000', PREMISE = 'EDEDED';
const TOTAL = 14400; // A4 横版减页边距，DXA

/* ---------- 行内解析：**bold** `code` <em>红</em> <strong>粗</strong> ---------- */
const unesc = (t) => t.replace(/&lt;/g,'<').replace(/&gt;/g,'>')
  .replace(/&nbsp;/g,' ').replace(/&quot;/g,'"').replace(/&amp;/g,'&');
function runs(text, o = {}) {
  text = unesc(text);
  const out = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`|<em>[\s\S]*?<\/em>|<strong>[\s\S]*?<\/strong>)/g;
  let last = 0, m;
  const push = (t, kind) => {
    if (!t) return;
    out.push(new TextRun({
      text: t,
      font: kind === 'code' ? { ascii: MONO, eastAsia: CN } : { ascii: EN, eastAsia: CN },
      size: o.size || 20,
      bold: kind === 'bold' || kind === 'red' || o.bold,
      color: kind === 'red' ? RED : (kind === 'code' ? '9C2A00' : (o.color || '000000'))
    }));
  };
  /* 嵌套标记：<em> 与 <strong> 可互相嵌套，红色优先（红色本身已是粗体） */
  const walk = (s, kind) => {
    const r = /(\*\*[^*]+\*\*|`[^`]+`|<em>[\s\S]*?<\/em>|<strong>[\s\S]*?<\/strong>)/g;
    let l = 0, mm;
    while ((mm = r.exec(s)) !== null) {
      push(s.slice(l, mm.index), kind);
      const tk = mm[0];
      if (tk.startsWith('**')) walk(tk.slice(2, -2), kind === 'red' ? 'red' : 'bold');
      else if (tk.startsWith('`')) push(tk.slice(1, -1), 'code');
      else if (tk.startsWith('<em>')) walk(tk.slice(4, -5), 'red');
      else walk(tk.slice(8, -9), kind === 'red' ? 'red' : 'bold');
      l = mm.index + tk.length;
    }
    push(s.slice(l), kind);
  };
  walk(text, undefined);
  last = text.length; m = null; re.lastIndex = 0;
  return out.length ? out : [new TextRun({ text: '', font: { ascii: EN, eastAsia: CN }, size: 20 })];
}

const P = (text, o = {}) => new Paragraph({
  children: runs(text, o),
  spacing: { before: o.before !== undefined ? o.before : 60, after: o.after !== undefined ? o.after : 60, line: 300 },
  indent: o.ind ? { left: o.ind } : undefined,
  alignment: o.align
});

/* 「0.1　飞机尺寸…」→ 书签名 SEC_0_1，供导航表 PAGEREF 引用 */
function bmk(text) {
  const m = String(text).match(/^(\d+(?:\.\d+)*)[\s\u3000]/);
  return m ? 'SEC_' + m[1].replace(/\./g, '_') : null;
}
function H(text, level, brk) {
  const sizes = { 1: 30, 2: 24, 3: 21, 4: 20 };
  const id = level <= 2 ? bmk(text) : null;
  const tr = new TextRun({ text, font: { ascii: EN, eastAsia: CN }, size: sizes[level], bold: true, color: '000000' });
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : level === 3 ? HeadingLevel.HEADING_3 : HeadingLevel.HEADING_4,
    children: id ? [new Bookmark({ id, children: [tr] })] : [tr],
    spacing: { before: level === 1 ? 320 : 220, after: level === 1 ? 140 : 100 },
    border: level === 1 ? { bottom: { style: BorderStyle.SINGLE, size: 8, color: GRAY, space: 6 } } : undefined,
    pageBreakBefore: !!brk
  });
}

function bullet(text, depth) {
  return new Paragraph({
    children: runs(text),
    bullet: { level: depth },
    spacing: { before: 40, after: 40, line: 290 }
  });
}
function numbered(text, depth) {
  return new Paragraph({
    children: runs(text),
    numbering: { reference: 'num', level: depth },
    spacing: { before: 40, after: 40, line: 290 }
  });
}

function codeBlock(lines) {
  return new Table({
    columnWidths: [TOTAL],
    width: { size: TOTAL, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: TOTAL, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: CODE },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 2, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
          left: { style: BorderStyle.SINGLE, size: 12, color: GRAY }, right: { style: BorderStyle.SINGLE, size: 2, color: LINE }
        },
        margins: { top: 80, bottom: 80, left: 140, right: 100 },
        children: lines.map(l => new Paragraph({
          children: [new TextRun({ text: l || ' ', font: { ascii: MONO, eastAsia: MONO }, size: 17 })],
          spacing: { before: 0, after: 0, line: 250 }
        }))
      })]
    })]
  });
}


/* ---------- 模块速查导航表（页码用 PAGEREF 域，F9 刷新） ---------- */
function navTable(rows) {
  const W = [1500, 10300, 2600];
  const cell = (children, o = {}) => new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: o.fill || 'FFFFFF' },
    margins: { top: 20, bottom: 20, left: 110, right: 110 },
    verticalAlign: VerticalAlign.CENTER,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 2, color: LINE }, right: { style: BorderStyle.SINGLE, size: 2, color: LINE }
    },
    children
  });
  const hdr = new TableRow({
    tableHeader: true,
    children: ['编　号', '知　识　点', '页　码'].map((t, i) => cell(
      [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 4, after: 4, line: 205 },
        children: [new TextRun({ text: t, font: { ascii: EN, eastAsia: CN }, size: 18, bold: true })] })],
      { w: W[i], fill: HDR }))
  });
  const body = rows.map((r, ri) => {
    const fill = ri % 2 ? ALT : 'FFFFFF';
    const id = 'SEC_' + r.id.replace(/\./g, '_');
    return new TableRow({
      cantSplit: true,
      children: [
        cell([new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 4, after: 4, line: 205 },
          children: [new TextRun({ text: r.id, font: { ascii: EN, eastAsia: CN }, size: 18, bold: true })] })], { w: W[0], fill }),
        cell([new Paragraph({ spacing: { before: 4, after: 4, line: 205 }, children: runs(r.title, { size: 18 }) })], { w: W[1], fill }),
        cell([new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 4, after: 4, line: 205 },
          children: [new PageReference(id)] })], { w: W[2], fill })
      ]
    });
  });
  return new Table({ columnWidths: W, width: { size: W[0] + W[1] + W[2], type: WidthType.DXA }, rows: [hdr, ...body] });
}

/* ---------- 内嵌 HTML 表格 ---------- */
function parseHtmlTable(html) {
  const rows = [];
  const trRe = /<tr([^>]*)>([\s\S]*?)<\/tr>/g;
  let m;
  while ((m = trRe.exec(html)) !== null) {
    const cls = (m[1].match(/class="([^"]*)"/) || [, ''])[1];
    const cells = [];
    const cRe = /<(td|th)([^>]*)>([\s\S]*?)<\/\1>/g;
    let c;
    while ((c = cRe.exec(m[2])) !== null) {
      const at = c[2];
      cells.push({
        head: c[1] === 'th',
        colspan: parseInt((at.match(/colspan="(\d+)"/) || [, 1])[1], 10),
        rowspan: parseInt((at.match(/rowspan="(\d+)"/) || [, 1])[1], 10),
        text: c[3]
      });
    }
    rows.push({ cls, cells });
  }
  return rows;
}

/* 表格后的间隔段落：若紧接着是标题或分页符，则不加（避免多出一张空白页） */
function tableGap(src, i) {
  let j = i;
  while (j < src.length && !src[j].trim()) j++;
  const nxt = j < src.length ? src[j].trim() : '';
  if (!nxt || nxt.startsWith('#') || nxt === '%%PAGEBREAK%%' || nxt === '---') return null;
  return P('', { before: 0, after: 60 });
}

function htmlTable(html) {
  const parsed = parseHtmlTable(html);
  if (!parsed.length) return null;
  const nCols = parsed[0].cells.reduce((a, c) => a + c.colspan, 0) || 2;

  /* 列宽按内容长度加权：取每列最长单元格的视觉宽度（中日韩字符算 2） */
  const vis = (t) => {
    const s = unesc(String(t)).replace(/<[^>]+>/g, '').replace(/\*\*/g, '');
    let n = 0;
    for (const ch of s) n += /[\u2E80-\u9FFF\uFF00-\uFFEF]/.test(ch) ? 2 : 1;
    return n;
  };
  /* 占位网格：rowspan 会让后续行少一个单元格，必须据此推算每个单元格真正的起始列 */
  const occ = [];
  const startCol = [];                         // startCol[ri][k] = 第 ri 行第 k 个单元格的起始列
  parsed.forEach((r, ri) => {
    if (!occ[ri]) occ[ri] = [];
    startCol[ri] = [];
    let ci = 0;
    r.cells.forEach((c, k) => {
      while (occ[ri][ci]) ci++;
      startCol[ri][k] = ci;
      for (let rr = ri; rr < ri + c.rowspan; rr++) {
        if (!occ[rr]) occ[rr] = [];
        for (let cc = ci; cc < ci + c.colspan; cc++) occ[rr][cc] = true;
      }
      ci += c.colspan;
    });
  });

  const dataNeed = new Array(nCols).fill(0);   // 数据行
  const hdrNeed = new Array(nCols).fill(0);   // 表头行
  parsed.forEach((r, ri) => {
    const isH = r.cls.includes('hdr');
    if (r.cls.includes('note') || r.cls.includes('premise')) return;
    r.cells.forEach((c, k) => {
      const ci = startCol[ri][k];
      if (c.colspan === 1 && ci < nCols) {
        const longest = Math.max(...String(c.text).split(/<br\s*\/?>/).map(vis));
        const t = isH || c.head ? hdrNeed : dataNeed;
        t[ci] = Math.max(t[ci], longest);
      }
    });
  });
  /* 内容多的表格撑满页宽以减少折行行数；内容少的按内容长度收缩 */
  let maxCell = 0, totalVis = 0;
  for (const r of parsed) {
    for (const c of r.cells) { const v = vis(c.text); if (v > maxCell) maxCell = v; totalVis += v; }
  }
  const heavy = maxCell >= 60 || totalVis >= 600;

  /* 每列「真实需要」的视觉宽度（不设上限）；表头权重 0.6，避免表头把数字列撑空 */
  const raw = dataNeed.map((d, i) => Math.max(d, hdrNeed[i] * 0.6, 3));

  /* 总宽：内容多的撑满页宽，内容少的按需收缩 */
  const CHAR = 132, PAD = 170, FLOOR = 520;
  const wantedAll = raw.reduce((a, n) => a + Math.round(n * CHAR) + PAD, 0);
  const TW = heavy ? TOTAL : Math.min(TOTAL, Math.max(wantedAll, Math.round(TOTAL * 0.42)));

  /* 「根据内容自动调整」：短标签列（故障 / 角色 / 序 / 步骤…）按内容定宽，
     余下的宽度全部让给内容最多的列（处置 / 动作 / 说明），以减少折行行数。 */
  const narrowSet = new Set();
  const NARROW = 20;                    // ≤10 个汉字视为短标签列（项目 / 场景 / 序号 / 数值列）
  const wideIdx = [], W = new Array(nCols).fill(0);
  let fixed = 0;
  for (let i = 0; i < nCols; i++) {
    if (raw[i] <= NARROW) {
      narrowSet.add(i);
      W[i] = Math.max(FLOOR, Math.round(raw[i] * CHAR) + PAD + 160);  // +160：覆盖单元格左右边距，短标签不折行
      fixed += W[i];
    } else wideIdx.push(i);
  }
  if (!wideIdx.length) {                // 全是短列：按内容比例铺开
    const sum0 = raw.reduce((a, b) => a + b, 0);
    for (let i = 0; i < nCols; i++) W[i] = Math.floor(TW * raw[i] / sum0);
  } else {
    let rest = TW - fixed;
    if (rest < wideIdx.length * 1600) { // 短列占太多：整体按比例回退
      const sum0 = raw.reduce((a, b) => a + b, 0);
      for (let i = 0; i < nCols; i++) W[i] = Math.floor(TW * raw[i] / sum0);
    } else {
      const sumW = wideIdx.reduce((a, i) => a + raw[i], 0);
      for (const i of wideIdx) W[i] = Math.floor(rest * raw[i] / sumW);
    }
  }
  /* 取整误差补给最后一个宽列（补给短列会让它无谓变宽） */
  const fixIdx = wideIdx.length ? wideIdx[wideIdx.length - 1] : nCols - 1;
  W[fixIdx] += TW - W.reduce((a, b) => a + b, 0);

  /* ---- 自动缩排：估算表格高度，超过一页时逐级缩小字号，尽量整表放在同一页 ---- */
  const PAGE_H = 11906 - 1800;          // A4 横版可用高度（DXA）
  const BUDGET = PAGE_H - 900;          // 留出小节标题与段间距
  const estimate = (sz) => {
    const unit = 132 * sz / 20;         // 每「视觉单位」宽度
    let h = 0;
    parsed.forEach((r, ri) => {
      let maxLines = 1;
      r.cells.forEach((c, ck) => {
        const ci = startCol[ri][ck];
        let w = 0;
        for (let k = 0; k < c.colspan; k++) w += W[Math.min(ci + k, nCols - 1)];
        const perLine = Math.max(4, (w - 180) / unit);
        let lines = 0;
        for (const seg of String(c.text).split(/<br\s*\/?>/)) lines += Math.max(1, Math.ceil(vis(seg.trim()) / perLine));
        if (lines > maxLines) maxLines = lines;
      });
      h += maxLines * Math.round(17.5 * sz) + 120 + 40;   // 行高 + 单元格上下边距 + 段前后
    });
    return h;
  };
  let FS = 18, LN = 270, CM = 60;
  if (estimate(FS) > BUDGET) {
    /* 只在「缩到某一号真的能塞进一页」时才缩；否则保持正常字号，让它自然分页
       ——避免既缩成小字又照样跨页的最差结果 */
    const fit = [17, 16, 15, 14].find(c => estimate(c) <= BUDGET);
    if (fit) { FS = fit; LN = 250; CM = 40; }
  }
  /* 表格整体能放进一页时，让 Word 尽量不要在中间断开 */
  /* ---- 逐列判定是否居中：数字列、项目 / 参数名称等短文本列一律水平居中 ---- */
  const NUMRE = /^[\s\d.,:;%~～×+\-—–/()（）°'"≤≥＜＞<>=A-Za-z①-⑳ⅠⅡⅢⅣⅤ一二三四五六七八九十]+$/;
  const centerCols = new Set();
  for (let ci = 0; ci < nCols; ci++) {
    let maxV = 0, n = 0, num = 0;
    parsed.forEach((r, ri) => {
      if (r.cls.includes('note') || r.cls.includes('premise') || r.cls.includes('hdr')) return;
      r.cells.forEach((c, ck) => {
        if (startCol[ri][ck] !== ci || c.colspan !== 1) return;
        const plain = unesc(String(c.text)).replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
        n++; const v = vis(plain); if (v > maxV) maxV = v;
        if (NUMRE.test(plain) || plain === '' || plain === '—' || plain === '/') num++;
      });
    });
    if (!n) continue;
    if (maxV <= 24 || num / n >= 0.7) centerCols.add(ci);   // ≤12 汉字的短列，或以数字为主的列
  }

  /* 表格绝不跨页：能放进一页的一律整表保持（放不下时才允许自然分页） */
  const keepTogether = estimate(FS) <= BUDGET;
  if (process.env.FIT_LOG) console.error('TBL rows=%d est=%d fs=%d keep=%s', parsed.length, estimate(FS), FS, keepTogether);

  const trs = parsed.map((r, ri) => {
    const isHdr = r.cls.includes('hdr');
    const isNote = r.cls.includes('note');
    const isPre = r.cls.includes('premise');
    const isWarn = r.cls.includes('warn');
    /* 警告行：白底 + 左侧红竖条。不用红底——红色是强调色，不是背景色 */
    const fill = isHdr ? HDR : isWarn ? 'FFFFFF' : isPre ? PREMISE : isNote ? 'FFFFFF' : (ri % 2 ? ALT : 'FFFFFF');
    const cells = r.cells.map((c, ck) => {
      const ci = startCol[ri][ck];
      let w = 0;
      for (let k = 0; k < c.colspan; k++) w += W[Math.min(ci + k, nCols - 1)];
      const isFirstCol = (ci === 0) && !isPre && !isNote && !isWarn;
      /* 表头、数字列、项目 / 参数名称等短文本列整列居中；说明类长列左对齐 */
      const center = isHdr || c.head
                   || (c.colspan === 1 && (centerCols.has(ci) || narrowSet.has(ci)))
                   || (isFirstCol && c.colspan === 1);
      const paras = String(c.text).split(/<br\s*\/?>/).map(seg =>
        new Paragraph({
          children: runs(seg.trim(), { bold: isHdr || c.head, size: FS }),
          spacing: { before: 20, after: 20, line: LN },
          keepNext: keepTogether && ri < parsed.length - 1,
          alignment: center ? AlignmentType.CENTER : undefined
        }));
      const borders = {
        top: { style: BorderStyle.SINGLE, size: 2, color: LINE },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
        left: isWarn ? { style: BorderStyle.SINGLE, size: 14, color: RED }
             : isPre  ? { style: BorderStyle.SINGLE, size: 14, color: GRAY }
                      : { style: BorderStyle.SINGLE, size: 2, color: LINE },
        right: { style: BorderStyle.SINGLE, size: 2, color: LINE }
      };
      return new TableCell({
        width: { size: w, type: WidthType.DXA },
        columnSpan: c.colspan > 1 ? c.colspan : undefined,
        rowSpan: c.rowspan > 1 ? c.rowspan : undefined,
        shading: { type: ShadingType.CLEAR, color: 'auto', fill },
        borders,
        margins: { top: CM, bottom: CM, left: 90, right: 90 },
        verticalAlign: (center && c.rowspan > 1) ? VerticalAlign.CENTER : VerticalAlign.TOP,
        children: paras.length ? paras : [new Paragraph('')]
      });
    });
    return new TableRow({ tableHeader: isHdr, cantSplit: true, children: cells });
  });
  return new Table({ columnWidths: W, width: { size: W.reduce((a,b)=>a+b,0), type: WidthType.DXA }, rows: trs });
}

function mdTable(rows) {
  const n = Math.max(...rows.map(r => r.length));
  const W = []; for (let i = 0; i < n; i++) W.push(Math.floor(TOTAL / n));
  W[n - 1] = TOTAL - W.slice(0, n - 1).reduce((a, b) => a + b, 0);
  const trs = rows.map((cells, ri) => new TableRow({
    tableHeader: ri === 0, cantSplit: true,
    children: Array.from({ length: n }, (_, ci) => new TableCell({
      width: { size: W[ci], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: ri === 0 ? HDR : (ri % 2 ? ALT : 'FFFFFF') },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 2, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
        left: { style: BorderStyle.SINGLE, size: 2, color: LINE }, right: { style: BorderStyle.SINGLE, size: 2, color: LINE }
      },
      margins: { top: 60, bottom: 60, left: 90, right: 90 },
      verticalAlign: VerticalAlign.TOP,
      children: [new Paragraph({
        children: runs(cells[ci] || '', { bold: ri === 0, size: 18 }),
        spacing: { before: 20, after: 20, line: 270 }
      })]
    }))
  }));
  return new Table({ columnWidths: W, width: { size: TOTAL, type: WidthType.DXA }, rows: trs });
}

/* ---------- 解析 markdown ---------- */
let rawSrc = fs.readFileSync(SRC, 'utf8');
if (rawSrc.startsWith('---')) {                 // 跳过 YAML front matter
  const end = rawSrc.indexOf('\n---', 3);
  if (end > 0) rawSrc = rawSrc.slice(rawSrc.indexOf('\n', end + 1) + 1);
}
const src = rawSrc.split(/\r?\n/);
const body = [];
let i = 0, docTitle = '', pendingBreak = false;

while (i < src.length) {
  let ln = src[i];

  if (/^```/.test(ln)) {                       // 代码块
    const buf = []; i++;
    while (i < src.length && !/^```/.test(src[i])) buf.push(src[i++]);
    i++; body.push(codeBlock(buf)); body.push(P('', { before: 0, after: 40 })); continue;
  }
  if (/^\s*<table/.test(ln)) {                 // 内嵌 HTML 表格
    const buf = [];
    while (i < src.length && !/<\/table>/.test(src[i])) buf.push(src[i++]);
    buf.push(src[i++]);
    const tb = htmlTable(buf.join('\n'));
    if (tb) { body.push(tb); const g = tableGap(src, i); if (g) body.push(g); }
    continue;
  }
  if (/^\s*\|/.test(ln)) {                     // markdown 表格
    const raw = [];
    while (i < src.length && /^\s*\|/.test(src[i])) raw.push(src[i++]);
    const rows = raw
      .filter(r => !/^\s*\|[\s:|-]+\|\s*$/.test(r))
      .map(r => r.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim()));
    body.push(mdTable(rows)); { const g = tableGap(src, i); if (g) body.push(g); } continue;
  }
  if (/^%%PAGEBREAK%%\s*$/.test(ln)) { pendingBreak = true; i++; continue; }
  if (/^%%NAV%%\s*$/.test(ln)) {
    i++; const rows = [];
    while (i < src.length && !/^%%ENDNAV%%\s*$/.test(src[i])) {
      const t = src[i++].trim();
      if (t) { const p = t.split('|'); rows.push({ id: p[0].trim(), title: p.slice(1).join('|').trim() }); }
    }
    i++; body.push(navTable(rows)); continue;
  }
  if (/^#\s+/.test(ln)) { docTitle = ln.replace(/^#\s+/, ''); i++; continue; }
  if (/^##\s+/.test(ln)) { body.push(H(ln.replace(/^##\s+/, ''), 1, pendingBreak)); pendingBreak = false; i++; continue; }
  if (/^#####\s+/.test(ln)) { body.push(H(ln.replace(/^#####\s+/, ''), 4, pendingBreak)); pendingBreak = false; i++; continue; }
  if (/^####\s+/.test(ln)) { body.push(H(ln.replace(/^####\s+/, ''), 3, pendingBreak)); pendingBreak = false; i++; continue; }
  if (/^###\s+/.test(ln)) { body.push(H(ln.replace(/^###\s+/, ''), 2, pendingBreak)); pendingBreak = false; i++; continue; }
  if (/^---\s*$/.test(ln)) { i++; continue; }
  if (/^>\s?/.test(ln)) {
    const buf = [];
    while (i < src.length && /^>\s?/.test(src[i])) buf.push(src[i++].replace(/^>\s?/, ''));
    body.push(new Paragraph({
      children: runs(buf.join(' '), { size: 19 }),
      spacing: { before: 80, after: 120, line: 290 },
      indent: { left: 260 },
      border: { left: { style: BorderStyle.SINGLE, size: 12, color: GRAY, space: 10 } }
    }));
    continue;
  }
  {
    const mb = ln.match(/^(\s*)-\s+(.*)$/);
    if (mb) { body.push(bullet(mb[2], Math.min(2, Math.floor(mb[1].length / 2)))); i++; continue; }
    const mn = ln.match(/^(\s*)\d+\.\s+(.*)$/);
    if (mn) { body.push(numbered(mn[2], Math.min(2, Math.floor(mn[1].length / 3)))); i++; continue; }
  }
  if (!ln.trim()) { i++; continue; }
  if (/^详见\s/.test(ln.trim())) {              // 回查入口行：小号灰字
    body.push(new Paragraph({
      children: [new TextRun({ text: ln.trim(), font: { ascii: EN, eastAsia: CN }, size: 16, color: GRAY })],
      spacing: { before: 20, after: 160 }
    }));
    i++; continue;
  }
  body.push(P(ln)); i++;
}


/* ---------- 封面 + 目录 ---------- */
const AUTHOR = process.env.DOC_AUTHOR || '';
const rule = (o) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: o.before || 0, after: o.after || 0, line: 20 },
  children: [new TextRun({ text: '', size: 2 })],
  border: { bottom: { style: BorderStyle.SINGLE, size: o.size || 6, color: o.color || LINE, space: 2 } }
});
const front = [
  /* ---------- 封面 ---------- */
  new Paragraph({ spacing: { before: 2200 }, children: [] }),
  rule({ size: 18, color: '000000', after: 60 }),          // 主标题上方粗线
  new Paragraph({
    children: [new TextRun({ text: docTitle, font: { ascii: EN, eastAsia: CN }, size: 72, bold: true, color: '000000' })],
    alignment: AlignmentType.CENTER, spacing: { before: 420, after: 300, line: 520 }
  }),
  new Paragraph({
    children: [new TextRun({ text: process.env.DOC_SUBTITLE || '737-NG / 737-8　系统 · 运行 · 训练', font: { ascii: EN, eastAsia: CN }, size: 26, color: GRAY, characterSpacing: 30 })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 300, line: 320 }
  }),
  rule({ size: 6, color: LINE, after: 0 }),                 // 副标题下方细线
  ...(AUTHOR ? [new Paragraph({
    children: [new TextRun({ text: AUTHOR, font: { ascii: EN, eastAsia: CN }, size: 22, color: GRAY })],
    alignment: AlignmentType.CENTER, spacing: { before: 520, after: 0, line: 300 }
  })] : []),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({
    children: [new TextRun({ text: '目　　录', font: { ascii: EN, eastAsia: CN }, size: 36, bold: true, characterSpacing: 40 })],
    alignment: AlignmentType.CENTER, spacing: { before: 200, after: 360 }
  }),
  new TableOfContents('目录', { hyperlink: true, headingStyleRange: '1-2' }),
  new Paragraph({ children: [new PageBreak()] })
];

const doc = new Document({
  features: { updateFields: true },
  numbering: {
    config: [{
      reference: 'num',
      levels: [0, 1, 2].map(l => ({
        level: l, format: 'decimal', text: `%${l + 1}.`, alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: 340 + l * 340, hanging: 300 } } }
      }))
    }]
  },
  styles: {
    default: { document: { run: { font: { ascii: EN, eastAsia: CN }, size: 20 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 30, bold: true, color: '000000', font: { ascii: EN, eastAsia: CN } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, color: '000000', font: { ascii: EN, eastAsia: CN } } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 21, bold: true, color: '000000', font: { ascii: EN, eastAsia: CN } } },
      { id: 'Heading4', name: 'Heading 4', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 20, bold: true, color: '000000', font: { ascii: EN, eastAsia: CN } } }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
        margin: { top: 900, bottom: 900, left: 900, right: 900 }
      },
      titlePage: true
    },
    footers: {
      /* 封面不显示页码；正文页码在右下角，小五号（9pt），仅「第 X 页」 */
      first: new Footer({ children: [new Paragraph({ children: [] })] }),
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ children: ['第 ', PageNumber.CURRENT, ' 页'], font: { ascii: EN, eastAsia: CN }, size: 18, color: GRAY })]
        })]
      })
    },
    children: front.concat(body)
  }]
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('written ' + OUT + '  (' + body.length + ' blocks)'); });
