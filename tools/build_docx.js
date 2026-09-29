/*  build_docx.js —— markdown 源 → 带目录的 A4 横版 Word
 *  用法：node build_docx.js <input.md> [output.docx]
 *  支持：# / ## / ### 标题、段落、- 项目符号、1. 编号、表格、```代码块```、> 引用、**加粗**、`代码`
 *  目录用 Word TOC 域，打开文档后按 F9 或「更新域」刷新页码。
 */
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, TableOfContents,
  WidthType, ShadingType, BorderStyle, AlignmentType, VerticalAlign, HeadingLevel,
  PageBreak, Footer, PageNumber, PageOrientation, Bookmark, PageReference, InternalHyperlink, TableLayoutType
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
const HIDE_TBD = process.env.SHOW_TBD !== '1';   // 成品默认不显示〔待补来源〕（用户要求：表格与正文内不标来源）
// DOC_PORTRAIT=1：竖版 A4（iPad 阅读版）；默认横版
const PORTRAIT = process.env.DOC_PORTRAIT === '1';
const TOTAL = PORTRAIT ? 10000 : 14400; // A4 减页边距，DXA
const SC = (w) => Math.round(w * TOTAL / 14400);

/* ---------- 行内解析：**bold** `code` <em>红</em> <strong>粗</strong> ---------- */
const unesc = (t) => t.replace(/&lt;/g,'<').replace(/&gt;/g,'>')
  .replace(/&nbsp;/g,' ').replace(/&quot;/g,'"').replace(/&amp;/g,'&');
function runs(text, o = {}) {
  text = unesc(text);
  const out = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`|<em>[\s\S]*?<\/em>|<strong>[\s\S]*?<\/strong>|<small>[\s\S]*?<\/small>|〔待补来源〕|<br\s*\/?>)/g;
  let last = 0, m;
  const push = (t, kind) => {
    if (!t) return;
    /* 圈码 ①–⑳ 用中文字体排（西文字体没有这些字形，回退字体常没有粗体，导致首列序号不加粗） */
    if (kind !== 'code' && /[\u2460-\u2473]/.test(t) && !/^[\u2460-\u2473]+$/.test(t)) {
      t.split(/([\u2460-\u2473]+)/).forEach(seg => push(seg, kind)); return;
    }
    const circ = /^[\u2460-\u2473]+$/.test(t);
    if (kind === 'red' && o.noRed) kind = 'bold';
    const grayK = /^gray/.test(kind || '');
    /* 比较符 / 「约」与后面的数值之间用不换行空格，避免「<」在行尾、数值折到下一行 */
    t = t.replace(/【[^】]*】/g, m0 => m0.replace(/ /g, '\u00A0')).replace(/(\d{1,2}:\d{2}) ([—–-]) (\d{1,2}:\d{2})/g, '$1\u00A0$2\u00A0$3');
    t = t.replace(/([<>≤≥=＜＞≈约±]) (?=[\d−\-+.])/g, '$1\u00A0').replace(/(\d) (?=(kg|ft|kt|nm|NM|m|km|psi|psid|fpm|min|s|h|%|°|℃)(?![A-Za-z]))/g, '$1\u00A0');
    out.push(new TextRun({
      text: t,
      font: kind === 'code' ? { ascii: MONO, eastAsia: CN } : circ ? { ascii: CN, hAnsi: CN, eastAsia: CN } : { ascii: EN, eastAsia: CN },
      size: grayK && !o.inTable ? (o.size || 20) - 2 : (o.size || 20),
      bold: kind === 'bold' || kind === 'red' || kind === 'graybold' || o.bold,
      color: (kind === 'red' && !o.noRed) ? RED : (kind === 'code' ? '9C2A00' : (grayK ? (o.inTable ? '4A4A4A' : GRAY) : (o.color || '000000')))
    }));
  };
  /* 嵌套标记：<em> 与 <strong> 可互相嵌套，红色优先（红色本身已是粗体） */
  const walk = (s, kind) => {
    const r = /(\*\*[^*]+\*\*|`[^`]+`|<em>[\s\S]*?<\/em>|<strong>[\s\S]*?<\/strong>|<small>[\s\S]*?<\/small>|〔待补来源〕|<br\s*\/?>)/g;
    let l = 0, mm;
    while ((mm = r.exec(s)) !== null) {
      push(s.slice(l, mm.index), kind);
      const tk = mm[0];
      if (tk.startsWith('**')) walk(tk.slice(2, -2), kind === 'red' ? 'red' : (/^gray/.test(kind || '') ? 'graybold' : 'bold'));
      else if (tk.startsWith('`')) push(tk.slice(1, -1), 'code');
      else if (tk.startsWith('<em>')) {
        /* 红色只标数值与关键禁令短语：整句（去标记后超过 28 字）改为黑色加粗 */
        const inner = tk.slice(4, -5), plainLen = unesc(inner.replace(/<[^>]+>/g, '')).length;
        const plainIn = unesc(inner.replace(/<[^>]+>/g, ''));
        const hasNum = /\d/.test(plainIn);
        const ban = /不得|禁止|严禁|不要|必须|不能|不可|只能|仅|立即|切勿|不准/.test(plainIn);
        /* 红色只给：含数值的短语（≤ 20 字）或禁令短语（≤ 12 字）；解释段里不用红色 */
        const redOk = (hasNum && plainLen <= 20) || (ban && plainLen <= 12);
        walk(inner, /^gray/.test(kind || '') ? 'graybold' : (redOk ? 'red' : 'bold'));
      }
      else if (/^<br/.test(tk)) out.push(new TextRun({ break: 1 }));
      else if (tk.startsWith('<small>')) walk(tk.slice(7, -8), 'gray');      /* SD-33 学习解释：灰色小字 */
      else if (tk === '〔待补来源〕') { if (!HIDE_TBD) push(tk, 'gray'); }      /* 来源待补标记：成品不显示，汇总到条目来源行 */
      else walk(tk.slice(8, -9), kind === 'red' ? 'red' : (/^gray/.test(kind || '') ? 'graybold' : 'bold'));
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
  alignment: o.align,
  keepNext: o.keepNext || undefined
});

/* 「0.1　飞机尺寸…」→ 书签名 SEC_0_1，供导航表 PAGEREF 引用 */
function bmk(text) {
  const m = String(text).match(/^(\d+(?:\.\d+)*)[\s\u3000]/);
  return m ? 'SEC_' + m[1].replace(/\./g, '_') : null;
}
function H(text, level, brk, forceId) {
  const sizes = { 1: 30, 2: 24, 3: 21, 4: 20 };
  const id = forceId || (level <= 2 ? bmk(text) : null);
  text = unesc(String(text));
  const tr = new TextRun({ text, font: { ascii: EN, eastAsia: CN }, size: sizes[level], bold: true, color: '000000' });
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : level === 3 ? HeadingLevel.HEADING_3 : HeadingLevel.HEADING_4,
    children: id ? [new Bookmark({ id, children: [tr] })] : [tr],
    spacing: { before: level === 1 ? 320 : 220, after: level === 1 ? 140 : 100 },
    border: level === 1 ? { bottom: { style: BorderStyle.SINGLE, size: 8, color: GRAY, space: 6 } } : undefined,
    keepNext: true,                                  // 标题永远与下文同页
    /* 速查区竖版：排版预检发现会被拆页的表，其条目标题另起一页（BREAK_BEFORE=21,35） */
    pageBreakBefore: !!brk || (COMPACT && BREAKS.has((String(text).match(/^(\d+)\. /) || [])[1])) || itemBreak(text)
  });
}
/* 全书预排版（book_break.py）：条目标题按出现顺序编号，BREAK_IDX 里的条目另起一页 */
let ITEM_IDX = 0;
const BREAK_IDX = new Set(String(process.env.BREAK_IDX || '').split(',').filter(Boolean).map(Number));
function itemBreak(text) {
  if (!/^(\d+\. |[A-Z]-\d+\u3000)/.test(String(text))) return false;
  ITEM_IDX += 1;
  return BREAK_IDX.has(ITEM_IDX) && !ITEM_SKIP.has(ITEM_IDX);
}
const BREAKS = new Set(String(process.env.BREAK_BEFORE || '').split(',').filter(Boolean));

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
/* ---------- 目录页下方的「章节快速跳转」表：点章名跳到该章首页 ---------- */
function chapterJump(mods) {
  const W = [SC(11400), TOTAL - SC(11400)];
  const cell = (children, o = {}) => new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: o.fill || 'FFFFFF' },
    margins: { top: 40, bottom: 40, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 2, color: LINE }, right: { style: BorderStyle.SINGLE, size: 2, color: LINE }
    },
    children
  });
  const hdr = new TableRow({
    tableHeader: true,
    children: ['章　节', '页　码'].map((t, k) => cell(
      [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 6, after: 6, line: 220 },
        children: [new TextRun({ text: t, font: { ascii: EN, eastAsia: CN }, size: 19, bold: true })] })],
      { w: W[k], fill: HDR }))
  });
  const rows = mods.map((m, ri) => {
    const fill = ri % 2 ? ALT : 'FFFFFF';
    return new TableRow({
      cantSplit: true,
      children: [
        cell([new Paragraph({ spacing: { before: 6, after: 6, line: 220 }, children: [
          new InternalHyperlink({ anchor: m.id, children: [
            new TextRun({ text: m.text, font: { ascii: EN, eastAsia: CN }, size: 20, bold: true, color: '000000', underline: {} })
          ] })
        ] })], { w: W[0], fill }),
        cell([new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 6, after: 6, line: 220 }, children: [
          new PageReference(m.id)
        ] })], { w: W[1], fill })
      ]
    });
  });
  return new Table({ rows: [hdr, ...rows], width: { size: W[0] + W[1], type: WidthType.DXA }, columnWidths: W });
}

function navTable(rows) {
  const W = [SC(1500), TOTAL - SC(1500) - SC(2600), SC(2600)];
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

/* 单元格按 <br> 分段排版前，把跨段的 <em>/<strong>/<small> 在每段首尾补齐，避免标签原样漏到成品里 */
function balanceBr(text) {
  const segs = String(text).split(/<br\s*\/?>/);
  if (segs.length < 2) return text;
  let open = [];
  return segs.map(seg => {
    const pre = open.map(t => '<' + t + '>').join('');
    const st = open.slice();
    const re = /<(\/?)(em|strong|small)>/g; let m;
    while ((m = re.exec(seg)) !== null) {
      if (!m[1]) st.push(m[2]);
      else { const k = st.lastIndexOf(m[2]); if (k >= 0) st.splice(k, 1); }
    }
    open = st;
    return pre + seg + st.slice().reverse().map(t => '</' + t + '>').join('');
  }).join('<br>');
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
        text: balanceBr(c[3])
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
  if (/^(解释：|公司差异：|注：|<strong>注|出处：)/.test(nxt)) return null;   // 表后解释 / 差异 / 注紧跟表格
  return P('', { before: 0, after: 60 });
}

/* 窄表 + 长通栏说明：说明行若在窄表里要折成 4 行以上，就移到表外（前提行放表前、注释 / 警示行放表后），
   让表格保持按内容收窄，说明文字按正文宽度排，不再被挤成一长条 */
let PROBE = false, TBL_IDX = -1;
const TBL_TARGET = [];
function htmlTable(html) {
  if (!PROBE) TBL_IDX++;
  const parsed = parseHtmlTable(html);
  if (!parsed.length || !(COMPACT || FIT_ALL)) return [htmlTableCore(html)];
  const visC = (t) => { let n = 0; for (const ch of unesc(String(t).replace(/<[^>]+>/g, '').replace(HIDE_TBD ? /〔待补来源〕/g : /(?!)/g, ''))) n += /[\u2E80-\u9FFF\uFF00-\uFFEF]/.test(ch) ? 2 : 1.05; return n; };
  const nCols = parsed[0].cells.reduce((a, c) => a + c.colspan, 0) || 2;
  const need = new Array(nCols).fill(0);
  parsed.forEach(r => {
    if (/note|premise|warn/.test(r.cls)) return;
    let ci = 0;
    r.cells.forEach(c => { if (c.colspan === 1 && ci < nCols)
      need[ci] = Math.max(need[ci], ...String(c.text).split(/<br\s*\/?>/).map(visC)); ci += c.colspan; });
  });
  const natural = need.reduce((a, n) => a + Math.min(n, 64) * 96 + 370, 0);
  const tw = Math.min(TOTAL, natural);
  const isFull = r => /note|premise|warn/.test(r.cls) && r.cells.length === 1;
  const lines = r => String(r.cells[0].text).split(/<br\s*\/?>/).reduce((a, sg) => a + Math.max(1, Math.ceil(visC(sg) * 96 / Math.max(900, tw - 180))), 0);
  /* 通栏前提行一律移到表前作正文段落；较长的通栏注释行（约 40 字以上）或窄表里要折 4 行以上的注释移到表后。
     警示行（warn）保留在表内。用户要求：文字单独列出，下面附表格，不要为迁就表格把文字堆在一起（2026-09-29） */
  const fd = parsed.findIndex(r => !/hdr|premise/.test(r.cls));
  const lastD = parsed.length - 1 - [...parsed].reverse().findIndex(r => !/note|warn|premise/.test(r.cls));
  const out = parsed.filter((r, k) => isFull(r) &&
    ((r.cls.includes('premise') && k < fd) || (/note|warn/.test(r.cls) && k > lastD)));
  /* 单列表（只有一个表头 + 内容格）不是表格结构：排成「表头：」引导的正文段落 */
  if (nCols === 1) {
    const outP = [];
    parsed.forEach(r => r.cells.forEach(c => {
      const segs = balanceBr(String(c.text)).split(/<br\s*\/?>/).filter(x => x.trim());
      if (r.cls.includes('hdr') || c.head) outP.push(new Paragraph({ children: runs('<strong>' + segs.join(' ') + '</strong>', { size: 19 }), spacing: { before: 60, after: 20, line: 290 }, keepNext: true }));
      else segs.forEach(sg => outP.push(new Paragraph({ children: runs(sg.trim(), { size: 19 }), spacing: { before: 20, after: 20, line: 290 } })));
    }));
    return outP;
  }
  /* 表格中间的通栏注释 / 警示行：把表拆成「表 → 段落 → 表（带表头）」，不在表中间夹整行说明（2026-09-29 终检）。
     跨越该行的合并单元格存在时不拆。 */
  const spans = [];
  { const occ2 = []; parsed.forEach((r, ri) => { let ci = 0; r.cells.forEach(c => { while (occ2[ri] && occ2[ri][ci]) ci++;
      if (c.rowspan > 1) spans.push([ri, ri + c.rowspan - 1]);
      for (let rr = ri; rr < ri + c.rowspan; rr++) { occ2[rr] = occ2[rr] || []; for (let cc = ci; cc < ci + c.colspan; cc++) occ2[rr][cc] = true; }
      ci += c.colspan; }); }); }
  const mids = parsed.map((r, k) => k).filter(k => isFull(parsed[k]) && /note|warn/.test(parsed[k].cls) && k > fd && k < lastD
    && !spans.some(([a0, b0]) => a0 < k && b0 > k));
  if (!out.length && !mids.length) return [htmlTableCore(html)];
  const firstData = parsed.findIndex(r => !/hdr|premise/.test(r.cls));
  const pre = [], post = [];
  const trRe = /<tr([^>]*)>([\s\S]*?)<\/tr>/g;
  let idx = -1;
  const kept = html.replace(trRe, (all) => { idx++; const r = parsed[idx];
    if (!out.includes(r)) return all;
    (idx < firstData ? pre : post).push(r); return ''; });
  const para = (r) => new Paragraph({
    children: runs(String(r.cells[0].text).replace(/<br\s*\/?>/g, '\n'), { size: 18 }).map(x => x),
    spacing: { before: 60, after: 60, line: 270 }, indent: { left: 120 },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: r.cls.includes('warn') ? RED : GRAY, space: 8 } },
    keepNext: pre.includes(r)
  });
  const paras = (arr) => arr.flatMap(r => balanceBr(String(r.cells[0].text)).split(/<br\s*\/?>/).filter(x => x.trim()).map((sg, k, a) => new Paragraph({
    children: runs(sg.trim(), { size: 19 }),
    spacing: { before: k ? 20 : 100, after: k === a.length - 1 ? 100 : 20, line: 290 },
    /* 警示行移出表格后保留左侧红竖条 */
    indent: r.cls.includes('warn') ? { left: 120 } : undefined,
    border: r.cls.includes('warn') ? { left: { style: BorderStyle.SINGLE, size: 14, color: RED, space: 8 } } : undefined,
    keepNext: pre.includes(r)
  })));
  if (!mids.length) return [...paras(pre), htmlTableCore(kept), ...paras(post)];
  const rowsHtml = []; html.replace(trRe, (all) => { rowsHtml.push(all); return all; });
  const hdrIdx = parsed.map((r, k) => k).filter(k => k < firstData && parsed[k].cls.includes('hdr'));
  const hdrHtml = hdrIdx.map(k => rowsHtml[k]).join('\n');
  const partsOut = [...paras(pre)];
  let cur = [];
  const flush = () => { if (cur.length) partsOut.push(htmlTableCore('<table class="ftn">\n' + hdrHtml + '\n' + cur.join('\n') + '\n</table>')); cur = []; };
  parsed.forEach((r, k) => {
    if (k < firstData || out.includes(r)) return;
    if (mids.includes(k)) { flush(); partsOut.push(...paras([r])); return; }
    cur.push(rowsHtml[k]);
  });
  flush();
  return [...partsOut, ...paras(post)];
}

function htmlTableCore(html) {
  const parsed = parseHtmlTable(html);
  if (!parsed.length) return null;
  const nCols = parsed[0].cells.reduce((a, c) => a + c.colspan, 0) || 2;

  /* 列宽按内容长度加权：取每列最长单元格的视觉宽度（中日韩字符算 2） */
  const vis = (t) => {
    const s = unesc(String(t).replace(/<[^>]+>/g, '').replace(HIDE_TBD ? /〔待补来源〕/g : /(?!)/g, '')).replace(/\*\*/g, '');
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
  /* 速查区：短列（最长一行 ≤ 15 个汉字的短标签、短语）给足宽度、一律不折行；
     剩余宽度分给长句列：先按各列最长一行（封顶约 28 字）给，放不下再按比例压，但每列不低于约 11 字 */
  if (COMPACT || FIT_ALL) {
    /* 紧凑模式按实际字宽估算（9pt：汉字约 180 DXA；大写拉丁字母略宽、小写 / 数字约半个汉字、
       空格与标点更窄），让列宽「刚好放下」而不是按粗估留大片空白 */
    const CHAR = 96;
    const vis = (t) => {
      const s = unesc(String(t).replace(/<[^>]+>/g, '').replace(HIDE_TBD ? /〔待补来源〕/g : /(?!)/g, '')).replace(/\*\*/g, '');
      let n = 0;
      for (const ch of s) n += /[\u2E80-\u9FFF\uFF00-\uFFEF]/.test(ch) ? 2
        : /[A-Z]/.test(ch) ? 1.35 : /[a-z0-9]/.test(ch) ? 1.05 : /\s/.test(ch) ? 0.6 : 1.1;
      return Math.ceil(n);
    };
    const SHORT = 36, CAP = 64, MINL = 22, EXTRA = PAD + 200;
    /* 跨行合并的格子纵向空间多，可折成几行而不增加表高：按「最长一行 ÷ 跨行数」计 */
    const need2 = new Array(nCols).fill(0), segL = [];
    parsed.forEach((r, ri) => {
      if (r.cls.includes('hdr') || r.cls.includes('note') || r.cls.includes('premise')) return;
      r.cells.forEach((c, ck) => {
        if (c.colspan !== 1) return;
        const ci2 = startCol[ri][ck];
        const sgs = String(c.text).split(/<br\s*\/?>/);
        const hasMain = sgs.some(sg => !/^\s*(<[^>]+>)*\s*[（(]/.test(sg) && vis(sg) > 0);
        sgs.forEach(sg => { let v = vis(sg); if (v <= 0) return;
          /* 名称下另起一行的括注（补充说明）允许折行：按不超过约 10 个汉字计宽 */
          if (hasMain && /^\s*(<[^>]+>)*\s*[（(]/.test(sg)) v = Math.min(v, 20);
          /* 「A / B」并列的较长格：允许在「 / 」处折行，按最长一段计宽（用户 2026-09-29：1.3 A-2 数值列过宽） */
          if (v > 30 && / \/ /.test(sg)) v = Math.min(v, Math.max(...sg.split(/ \/ /).map(x => vis(x))) + 2);
          (segL[ci2] = segL[ci2] || []).push(v); });
      });
    });
    /* 每列所需宽度：以多数行一行排下为准；个别特别长的行（超过该列第 75 百分位长度的 1.5 倍）允许折行，
       避免一两行长括注把整列撑宽、其余行两侧留白（用户 2026-09-29：第 36 条首列仍有空白） */
    for (let k = 0; k < nCols; k++) {
      const a = (segL[k] || []).slice().sort((x, y) => x - y);
      if (!a.length) continue;
      const mx = a[a.length - 1], p75 = a[Math.min(a.length - 1, Math.floor(a.length * 0.75))];
      need2[k] = mx;   // （曾试过按第 75 百分位封顶，会把个别长格挤成细长条，已取消；只对名称下的括注放宽）
    }
    /* 表头：短表头（如「737-NG」）不折行；长表头按 0.6 计，允许折行，避免把数字列撑空 */
    const hdrV = new Array(nCols).fill(0);
    parsed.forEach((r, ri) => { if (r.cls.includes('hdr')) r.cells.forEach((c, ck) => {
      if (c.colspan === 1) hdrV[startCol[ri][ck]] = Math.max(hdrV[startCol[ri][ck]], ...String(c.text).split(/<br\s*\/?>/).map(vis)); }); });
    const full = need2.map((n, i2) => Math.max(n, hdrV[i2] <= 14 ? hdrV[i2] : hdrV[i2] * 0.6, 3));
    const isShort = full.map(n => n <= SHORT);
    /* 短列不折行，但不能挤占长句列：有长句列时，单个短列 ≤ 25% 页宽、短列合计 ≤ 50% 页宽，
       超出的从最宽的短列起改按长列参与分宽（允许在 <br> 之外折行） */
    const demoted = new Set();
    if (isShort.some(s => !s)) {
      const sw = i2 => Math.round(full[i2] * CHAR) + EXTRA;
      for (;;) {
        const sIdx = full.map((n, i2) => i2).filter(i2 => isShort[i2]);
        if (!sIdx.length) break;
        const tot = sIdx.reduce((a2, i2) => a2 + sw(i2), 0);
        const wid = sIdx.reduce((a2, i2) => (sw(i2) > sw(a2) ? i2 : a2), sIdx[0]);
        if (sw(wid) > TOTAL * 0.25 || tot > TOTAL * 0.5) { isShort[wid] = false; demoted.add(wid); } else break;
      }
    }
    const w = full.map((n, i2) => isShort[i2] ? Math.max(FLOOR, Math.round(n * CHAR) + EXTRA) : 0);
    const used = w.reduce((a2, b2) => a2 + b2, 0);
    const longIdx = full.map((n, i2) => i2).filter(i2 => !isShort[i2]);
    if (longIdx.length) {
      /* 每列内容总量（各数据格视觉宽度之和）：内容越多的长句列分到越宽，减少折行行数 */
      const amount = new Array(nCols).fill(0);
      parsed.forEach((r, ri) => {
        if (r.cls.includes('hdr') || r.cls.includes('note') || r.cls.includes('premise')) return;
        r.cells.forEach((c, ck) => { if (c.colspan === 1) amount[startCol[ri][ck]] += vis(String(c.text).replace(/<br\s*\/?>/g, '')); });
      });
      const oneLine = i2 => Math.round(full[i2] * CHAR) + EXTRA;            // 该列最长一格排成一行所需
      const room = TOTAL - used;
      const heavy = longIdx.some(i2 => full[i2] > CAP);                   // 有格子一行放不下：可以用满页宽
      const fits = longIdx.reduce((a2, i2) => a2 + Math.min(oneLine(i2), Math.round(CAP * CHAR) + EXTRA), 0);
      const target = heavy ? room : Math.min(room, fits);
      /* 贪心分宽：每次把一小段宽度给「能让整表总行数减少最多」的那一列（没有能减少的就给内容最多的列），
         直到用完目标宽度或各列都一行放得下——效果是：短语列尽量一行排完，长句列分到剩下的宽度 */
      const UNIT = CHAR;
      const rowsC = [];
      parsed.forEach((r, ri) => {
        if (r.cls.includes('hdr') || r.cls.includes('note') || r.cls.includes('premise')) return;
        rowsC.push(r.cells.map((c, ck) => ({ ci: startCol[ri][ck], span: c.colspan, rs: c.rowspan || 1,
          segs: String(c.text).split(/<br\s*\/?>/).map(sg => vis(sg.trim())) })));
      });
      /* 评分 = 行高总和 × 10 + 各格折行数之和：先压表格高度，高度不变时也尽量少折行 */
      const lines = (ww) => rowsC.reduce((h, row) => {
        const ls = row.map(c => {
          let wd = 0; for (let k = 0; k < c.span; k++) wd += ww[Math.min(c.ci + k, nCols - 1)];
          const per = Math.max(4, (wd - 180) / UNIT);
          return c.segs.reduce((a2, v) => a2 + Math.max(1, Math.ceil(v / per)), 0);
        });
        const hs = row.map((c, k) => ls[k] / c.rs);          // 合并格的行数摊到所跨各行
        return h + Math.max(1, ...hs) * 10 + ls.reduce((a2, x) => a2 + x, 0) * 2;
      }, 0);
      /* 被降级的短列（标签列）：括号前的主名称仍保证一行排下，只让括注部分折行 */
      const headVis = new Array(nCols).fill(0);
      parsed.forEach((r, ri) => {
        if (r.cls.includes('hdr') || r.cls.includes('note') || r.cls.includes('premise')) return;
        r.cells.forEach((c, ck) => { const ci3 = startCol[ri][ck]; if (c.colspan === 1 && demoted.has(ci3))
          String(c.text).split(/<br\s*\/?>/).forEach(sg => { headVis[ci3] = Math.max(headVis[ci3], vis(sg.trim().split(/[（(]/)[0])); }); });
      });
      longIdx.forEach(i2 => { w[i2] = Math.min(oneLine(i2), Math.round(TOTAL * 0.23), Math.round(Math.max(8, demoted.has(i2) ? headVis[i2] : 0) * CHAR) + EXTRA); });    // 起点：约 4 个汉字，其余交给贪心分配
      /* 长列下限：约 10 个汉字（放得下一行时取一行宽），避免某列被挤成每行几个字的细长条 */
      longIdx.forEach(i2 => { w[i2] = Math.max(w[i2], Math.min(oneLine(i2), 28 * CHAR + EXTRA)); });
      const start = longIdx.reduce((a2, i2) => a2 + w[i2], 0);
      if (start > target) longIdx.forEach(i2 => { w[i2] = Math.max(FLOOR, Math.floor(w[i2] * target / start)); });
      let left = target - longIdx.reduce((a2, i2) => a2 + w[i2], 0);
      const STEP = 150;
      while (left >= STEP) {
        const cand = longIdx.filter(i2 => w[i2] + STEP <= oneLine(i2) + STEP);
        if (!cand.length) break;
        const base = lines(w);
        let best = -1, gain = 0;
        for (const i2 of cand) { w[i2] += STEP; const g = base - lines(w); w[i2] -= STEP; if (g > gain) { gain = g; best = i2; } }
        if (best < 0) best = cand.reduce((a2, i2) => (amount[i2] / w[i2] > amount[a2] / w[a2] ? i2 : a2), cand[0]);
        w[best] += STEP; left -= STEP;
      }
    }
    /* 跨列合并格文字多、而所跨各列按短内容定宽时：把所跨列加宽到该格最长一行放得下（不超过页宽），
       减少合并格的折行行数、避免右侧大片空白（用户 2026-09-29：1.2 B-1「失速警告逻辑」） */
    {
      const spanCells = [];
      parsed.forEach((r, ri) => {
        if (r.cls.includes('hdr') || r.cls.includes('note') || r.cls.includes('premise')) return;
        r.cells.forEach((c, ck) => { if (c.colspan > 1) {
          const need = Math.max(...String(c.text).split(/<br\s*\/?>/).map(sg => vis(sg.trim())));
          spanCells.push({ ci: startCol[ri][ck], span: c.colspan, need }); } });
      });
      spanCells.sort((x, y) => y.need - x.need);
      for (const sc of spanCells) {
        const idx = []; for (let k = 0; k < sc.span; k++) idx.push(Math.min(sc.ci + k, nCols - 1));
        const cur = idx.reduce((a2, k) => a2 + w[k], 0);
        const others = w.reduce((a2, b2) => a2 + b2, 0) - cur;
        const want = Math.min(TOTAL - others, Math.round(sc.need * CHAR) + EXTRA);
        if (want > cur + 300) { const add = want - cur; idx.forEach(k => { w[k] += Math.floor(add * w[k] / cur); }); }
      }
    }
    const s = w.reduce((a2, b2) => a2 + b2, 0);
    for (let k2 = 0; k2 < nCols; k2++) W[k2] = s > TOTAL ? Math.floor(w[k2] * TOTAL / s) : w[k2];
  }
  /* 取整误差补给最后一个宽列（补给短列会让它无谓变宽） */
  const fixIdx = wideIdx.length ? wideIdx[wideIdx.length - 1] : nCols - 1;
  if (!(COMPACT || FIT_ALL)) W[fixIdx] += TW - W.reduce((a, b) => a + b, 0);
  if (PROBE) return { probeW: W.reduce((a2, b2) => a2 + b2, 0) };
  /* 同一条目里有多张表：统一到其中最宽者的宽度（多出的宽度按比例分给首列以外的各列） */
  {
    const tgt = TBL_TARGET[TBL_IDX] || 0, sumW = W.reduce((a2, b2) => a2 + b2, 0);
    /* 只做小幅对齐（本表已达目标宽度 85% 以上时）；差得多就保持按内容的宽度——表格不留空白优先（用户 2026-09-29） */
    if (tgt > sumW + 200 && sumW >= 0.85 * tgt) {
      const idxs = nCols > 1 ? [...Array(nCols).keys()].slice(1) : [0];
      const base = idxs.reduce((a2, k2) => a2 + W[k2], 0) || 1;
      idxs.forEach(k2 => { W[k2] += Math.floor((tgt - sumW) * W[k2] / base); });
    }
  }
  if (process.env.W_LOG && COMPACT) console.error('W', nCols, JSON.stringify(W), String(parsed[0].cells.map(c => c.text).join('/')).slice(0, 40));

  /* ---- 自动缩排：估算表格高度，超过一页时逐级缩小字号，尽量整表放在同一页 ---- */
  const PAGE_H = (PORTRAIT ? 16838 : 11906) - 1800;   // 可用高度（DXA）
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
    const fit = [17, 16].find(c => estimate(c) <= BUDGET);   // 最多缩一号半，不把整表缩成难读的小字
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
        const plain = unesc(String(c.text).replace(/<[^>]+>/g, '')).replace(/\s+/g, ' ').trim();
        n++; const v = vis(plain); if (v > maxV) maxV = v;
        if (NUMRE.test(plain) || plain === '' || plain === '—' || plain === '/') num++;
      });
    });
    if (!n) continue;
    if (maxV <= 24 || num / n >= 0.7) centerCols.add(ci);   // ≤12 汉字的短列，或以数字为主的列
  }

  /* 表格绝不跨页：能放进一页的一律整表保持（放不下时才允许自然分页） */
  /* 速查区：一列里过半的格是长句（很长或含多个分句）才整列左对齐；其余列（文字、数字、短句）居中 */
  /* 对齐（按内容判定）：一列里只要有格子要折行（首列标签除外：多个短标签分行仍居中），
     或有多段且含长句的格子，整列左对齐；其余（序号、数值、短标签、短语）居中 */
  const longCols = new Set(), longCell = new Set(), colN = [], colL = [], colV = [], colT = [], paraCells = [];
  if (COMPACT || FIT_ALL) {
    const unitW = 96 * FS / 18;   // 与列宽模型同一套字宽（紧凑模式 9pt 汉字 ≈ 2 × 96 DXA）
    const vis = (t) => {
      const s0 = unesc(String(t).replace(/<[^>]+>/g, '').replace(HIDE_TBD ? /〔待补来源〕/g : /(?!)/g, '')).replace(/\*\*/g, '');
      let n = 0;
      for (const ch of s0) n += /[\u2E80-\u9FFF\uFF00-\uFFEF]/.test(ch) ? 2
        : /[A-Z]/.test(ch) ? 1.35 : /[a-z0-9]/.test(ch) ? 1.05 : /\s/.test(ch) ? 0.6 : 1.1;
      return n;
    };
    parsed.forEach((r, ri) => {
      if (r.cls.includes('note') || r.cls.includes('premise') || r.cls.includes('hdr') || r.cls.includes('warn')) return;
      r.cells.forEach((c, ck) => {
        if (c.colspan !== 1 || c.head) return;
        const ci2 = startCol[ri][ck];
        const per = Math.max(4, (W[ci2] - 180) / unitW);
        const segs = String(c.text).split(/<br\s*\/?>/).map(sg => vis(sg.trim())).filter(v => v > 0);
        /* ===== 表格对齐固定规则（2026-09-29 用户：按内容判断，写成固定规则，全书统一） =====
           R1 表头一律居中。R2 首列标签列整列居中加粗。
           R3 「段落型」格子左对齐：分条列举（≥ 2 条 ①② / 1. 2.），或估算排版后 ≥ 4 行，
              或去标记后总长 > 90（约 45 个汉字），或含 ≥ 2 个句读（，。；）且总长 > 60。
              其余（数值、短语、一两行的短句）居中。
           R4 一列里段落型格子过半 → 整列左对齐（长文字列统一左齐）；否则整列居中，个别段落型格子单独左对齐。
           R5 「—」占位一律居中；所有格子纵向居中。 */
        const rawSegs = String(c.text).split(/<br\s*\/?>/).map(sg => unesc(sg.replace(/<[^>]+>/g, '')).replace(/〔待补来源〕/g, '').trim()).filter(Boolean);
        const enumN = rawSegs.filter(sg => /^([①-⑳]|\d+[.、)）]|[a-z][)）]|[A-Z]-\d+\s|第 ?\d+ ?[条步])/.test(sg)).length;
        const listy = rawSegs.length >= 3 && (enumN >= 2 || rawSegs.filter(sg => /[；;]$/.test(sg)).length >= 2);
        const lines = segs.reduce((a2, v) => a2 + Math.max(1, Math.ceil(v / (per * 1.04))), 0);
        const total = segs.reduce((a2, v) => a2 + v, 0);
        const puncts = (rawSegs.join('').match(/[，。；]/g) || []).length;
        /* 2026-09-29 终检：防同列「锯齿」——多行且带句读的也算段落型；段落型占三分之一以上整列左齐 */
        /* 段落型：分条列举（≥ 3 行）、排版后 ≥ 4 行、或总长超过约 70 个汉字；三行以内的短句（如 1.4 C-1 俯仰 / 横滚方式）居中 */
        const para = listy || lines >= 4 || total > 140;
        colN[ci2] = (colN[ci2] || 0) + 1;
        if (total > 30) colT[ci2] = (colT[ci2] || 0) + 1;   // 句子型格（约 15 字以上）
        if (para) { colL[ci2] = (colL[ci2] || 0) + 1; paraCells.push([ci2, c, listy || lines >= 4]); }
      });
    });
    /* 有段落型格子，且（段落型占三分之一以上，或多数格子是句子）→ 整列左齐，防同列锯齿 */
    for (let k3 = 0; k3 < nCols; k3++) if (colL[k3] && ((colL[k3] || 0) * 3 >= (colN[k3] || 1) || (colT[k3] || 0) * 2 > (colN[k3] || 1))) longCols.add(k3);
    /* 居中列里只有很长（≥ 4 行）或分条列举的格子单独左齐 */
    paraCells.forEach(([k3, c]) => longCell.add(c));
  }
  if (process.env.A_LOG && !PROBE) {
    const cols = [];
    for (let k = 0; k < nCols; k++) {
      let mx = 0, n = 0, multi = 0;
      parsed.forEach((r, ri) => { if (/hdr|note|premise|warn/.test(r.cls)) return;
        r.cells.forEach((c, ck) => { if (startCol[ri][ck] !== k || c.colspan !== 1 || c.head) return; n++;
          const segs = String(c.text).split(/<br\s*\/?>/);
          if (segs.length > 1) multi++;
          segs.forEach(sg => { let v = 0; for (const ch of unesc(sg.replace(/<[^>]+>/g, ''))) v += /[\u2E80-\u9FFF\uFF00-\uFFEF]/.test(ch) ? 2 : /[A-Z]/.test(ch) ? 1.35 : /[a-z0-9]/.test(ch) ? 1.05 : /\s/.test(ch) ? 0.6 : 1.1; mx = Math.max(mx, v); }); }); });
      cols.push({ w: W[k], need: Math.round(mx * 96 * FS / 18 + 180), n, multi, left: longCols.has(k) });
    }
    console.error('ALOG ' + JSON.stringify({ t: TBL_IDX, h: (hdrRowA => hdrRowA ? hdrRowA.cells.map(c => unesc(String(c.text).replace(/<[^>]+>/g, ''))).join('|') : '')(parsed.find(r => r.cls.includes('hdr'))), cols }));
  }
  /* 速查区：有合并单元格的表、或不到半页的表整表同页；其余大表允许分页（表头重复），避免整页留白 */
  const hasRowspan = parsed.some(r => r.cells.some(c => c.rowspan > 1));
  /* 一页放得下的表一律整表同页（宁可上一页留白，也不要拆开后多出一行重复表头） */
  const keepTogether = estimate(FS) <= BUDGET;
  if (process.env.FIT_LOG) console.error('TBL rows=%d est=%d fs=%d keep=%s', parsed.length, estimate(FS), FS, keepTogether);

  /* 顶部连续的通栏前提行 + 表头行一起作「重复标题行」（Word 要求标题行从第一行起连续） */
  let lead = 0;
  while (lead < parsed.length && (parsed[lead].cls.includes('hdr') || parsed[lead].cls.includes('premise'))) lead++;
  if (!parsed.slice(0, lead).some(r => r.cls.includes('hdr'))) lead = 0;
  /* 并列对比表（如「系统 A 供压组件 | 系统 B 供压组件」）：首列不是标签列，不加粗 */
  const hdrRow = parsed.find(r => r.cls.includes('hdr'));
  const normH = s => unesc(String(s).replace(/<[^>]+>/g, '')).replace(/[A-Za-z0-9\-（）()\s项个]/g, '');
  const parallel = !!(hdrRow && hdrRow.cells.length >= 2 && normH(hdrRow.cells[0].text) &&
    normH(hdrRow.cells[0].text) === normH(hdrRow.cells[1].text));
  /* 末尾的通栏注释 / 警示行不单独落到下一页：最后一行数据行与它们连在一起 */
  let lastData = parsed.length - 1;
  while (lastData > 0 && (parsed[lastData].cls.includes('note') || parsed[lastData].cls.includes('warn'))) lastData--;
  const tailNote = lastData < parsed.length - 1;
  if (process.env.B_LOG && COMPACT && (parallel || longCols.has(0))) console.error('NOBOLD', parallel ? 'parallel' : 'long', String(hdrRow ? hdrRow.cells.map(c => c.text).join('/') : '').slice(0, 40));
  if (process.env.P_LOG && COMPACT) { const tw = W.reduce((a2, b2) => a2 + b2, 0);
    parsed.forEach(r => { if (/premise|note|warn/.test(r.cls)) console.error('ROWLINES', r.cls, Math.ceil(vis(r.cells[0].text) * 90 / (tw - 180)), Math.round(tw * 100 / TOTAL) + '%', unesc(String(r.cells[0].text)).replace(/<[^>]+>/g, '').trim().slice(0, 24)); }); }
  /* 首列是否为标签列（整列判定，避免同一列有的加粗有的不加粗）：六成以上的格是短标签即整列加粗 */
  const shortTxt = (txt) => String(txt).split(/<br\s*\/?>/).every(seg => {
    const pl = unesc(seg.replace(/<[^>]+>/g, '')).replace(/〔待补来源〕/g, '').trim();
    const main = /^[（(]/.test(pl) ? '' : pl.split(/[（(]/)[0];   // 只看括号前的名称；括注允许折行
    const v = vis(main);
    return v <= 30 && !(/[，。；]/.test(main) && v > 20);
  });
  let fcN = 0, fcS = 0;
  parsed.forEach((r, ri) => { if (/hdr|note|premise|warn/.test(r.cls)) return;
    r.cells.forEach((c, ck) => { if (startCol[ri][ck] === 0 && c.colspan === 1 && !c.head) { fcN++; if (shortTxt(c.text)) fcS++; } }); });
  const firstIsLabel = fcN > 0 && fcS >= 0.6 * fcN;
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
      const shortCell = (txt) => String(txt).split(/<br\s*\/?>/).every(seg => {
        const pl = unesc(seg.replace(/<[^>]+>/g, '')).trim();
        const v = vis(pl);
        return v <= 30 && !(/[，。；]/.test(pl) && v > 20);
      });
      const placeholder = /^\s*(—|－|-|\/|)\s*$/.test(unesc(String(c.text).replace(/<[^>]+>/g, '')));
      /* 首列标签格：加粗、不标红（红色只留给数值与禁令） */
      const labelCol = isFirstCol && !parallel && c.colspan === 1 && nCols > 1 && firstIsLabel;
      /* R2：首列短标签（括号前名称短）一律居中，括注折成多行也不改左对齐（用户 2026-09-29：速查区第 93 条 MOC / OCA） */
      const labelShort = labelCol && (() => {
        let t = unesc(String(c.text).replace(/<br\s*\/?>/g, '').replace(/<[^>]+>/g, '')).replace(/〔待补来源〕/g, '');
        for (let k = 0; k < 3; k++) t = t.replace(/[（(][^（）()]*[）)]/g, '');
        return vis(t.trim()) <= 30 && !/[，。；]/.test(t); })();
      if (process.env.L_LOG && labelShort && longCell.has(c)) console.error('LBLFIX', unesc(String(c.text).replace(/<[^>]+>/g, '')).slice(0, 40));
      const center = placeholder || labelShort || (labelCol && !longCell.has(c)) || ((COMPACT || FIT_ALL)
        ? (isHdr || c.head || (!isNote && !isPre && !isWarn &&
             (c.colspan === 1 ? (!longCols.has(ci) && !longCell.has(c)) : shortCell(c.text))))
        : (isHdr || c.head
           || (c.colspan === 1 && (centerCols.has(ci) || narrowSet.has(ci)))
           || (isFirstCol && c.colspan === 1)));
      /* 首列序号格（只有一个圈码）：圈码字形在 Word 里常回退到无粗体的字体，改排为粗体阿拉伯数字 */
      const serial = isFirstCol && /^\s*(<strong>)?\s*[\u2460-\u2473]\s*(<\/strong>)?\s*$/.test(String(c.text));
      if (serial) c = Object.assign({}, c, { text: String(String(c.text).replace(/<[^>]+>/g, '').trim().charCodeAt(0) - 0x245F) });
      const paras = String(c.text).split(/<br\s*\/?>/).map(seg =>
        new Paragraph({
          /* 第一列（项目名 / 标签列）加粗，让表头行与首列都醒目；首列为长句列时不加粗 */
          children: runs(seg.trim(), { inTable: true, noRed: labelCol, bold: isHdr || c.head || labelCol, size: FS }),
          spacing: { before: 20, after: 20, line: LN },
          keepNext: ((isHdr || isPre) && ri < parsed.length - 1) || (keepTogether && ri < parsed.length - 1) || (tailNote && ri >= lastData && ri < parsed.length - 1),
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
        verticalAlign: VerticalAlign.CENTER,   // 用户 2026-09-29：内容尽量靠表格中心（纵向居中），排版更舒服
        children: paras.length ? paras : [new Paragraph('')]
      });
    });
    return new TableRow({ tableHeader: false, cantSplit: true, children: cells });   // 不重复表头（用户要求，2026-09-28）
  });
  /* 固定列宽：Word 默认「根据内容自动调整」会改写按内容算好的列宽，导致首列等短列被拉宽 */
  return new Table({ layout: TableLayoutType.FIXED, columnWidths: W, width: { size: W.reduce((a,b)=>a+b,0), type: WidthType.DXA }, rows: trs });
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
/* 另起一页的条目若紧跟在节 / 块标题（及其下的短说明）之后，把分页提到那个标题前，免得标题孤零零留在上一页底部 */
const ITEM_SKIP = new Set(), HEAD_BREAK = new Set();
if (BREAK_IDX.size) {
  const isHead = s => /^#{2,6}\s+/.test(s), isItem = s => /^#{2,6}\s+(\d+\. |[A-Z]-\d+\u3000)/.test(s);
  let k = 0;
  for (let L = 0; L < src.length; L++) {
    if (!isItem(src[L])) continue;
    k += 1; if (!BREAK_IDX.has(k)) continue;
    let target = -1, j = L - 1, paraOk = true;
    while (j >= 0) {
      const s = src[j];
      if (!s.trim()) { j--; continue; }
      if (isHead(s) && !isItem(s) && !/^#{2,6}\s+块索引/.test(s)) { target = j; paraOk = false; j--; continue; }
      if (paraOk && !/^\s*(<table|<tr|<\/table|\||%%)/.test(s) && s.length < 400) { j--; continue; }
      break;
    }
    if (target >= 0) { ITEM_SKIP.add(k); HEAD_BREAK.add(target); }
  }
}
const FIT_ALL = process.env.FIT_ALL !== '0';   // SD-27：全书表格按内容定宽
let COMPACT = false;          // 第零章速查区：表格按内容收宽、逐格判定对齐（SD-23）
/* 预排：先量出每张表按内容算的宽度，供同一条目内多表统一宽度 */
{
  PROBE = true;
  const widths = [], items = []; let item = 0, inCode = false;
  for (let L = 0; L < src.length; L++) {
    const s0 = src[L];
    if (/^```/.test(s0)) { inCode = !inCode; continue; }
    if (inCode) continue;
    if (/^%%COMPACT%%\s*$/.test(s0)) { COMPACT = true; continue; }
    if (/^%%ENDCOMPACT%%\s*$/.test(s0)) { COMPACT = false; continue; }
    if (/^#{1,6}\s+/.test(s0)) { item++; continue; }
    if (/^\s*<table/.test(s0)) {
      const buf = [];
      while (L < src.length && !/<\/table>/.test(src[L])) buf.push(src[L++]);
      buf.push(src[L]);
      const res = htmlTable(buf.join('\n'));
      const pw = res.find(x => x && x.probeW);
      widths.push(pw ? pw.probeW : 0); items.push(item);
    }
  }
  PROBE = false; COMPACT = false;
  const mx = {}, cnt = {};
  widths.forEach((w, k) => { if (w) { mx[items[k]] = Math.max(mx[items[k]] || 0, w); cnt[items[k]] = (cnt[items[k]] || 0) + 1; } });
  widths.forEach((w, k) => { TBL_TARGET[k] = (cnt[items[k]] > 1 && w && mx[items[k]] > w) ? mx[items[k]] : 0; });
}
const body = [];
const modules = [];          // 分册（## 级标题），供目录页的「章节快速跳转」表使用
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
    const tbs = htmlTable(buf.join('\n')).filter(Boolean);
    if (tbs.length) { body.push(...tbs); const g = tableGap(src, i); if (g) body.push(g); }
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
  if (/^%%COMPACT%%\s*$/.test(ln)) { COMPACT = true; i++; continue; }
  if (/^%%ENDCOMPACT%%\s*$/.test(ln)) { COMPACT = false; i++; continue; }
  if (/^%%NAV%%\s*$/.test(ln)) {
    i++; const rows = [];
    while (i < src.length && !/^%%ENDNAV%%\s*$/.test(src[i])) {
      const t = src[i++].trim();
      if (t) { const p = t.split('|'); rows.push({ id: p[0].trim(), title: p.slice(1).join('|').trim() }); }
    }
    i++; body.push(navTable(rows)); continue;
  }
  if (/^#\s+/.test(ln)) { docTitle = ln.replace(/^#\s+/, ''); i++; continue; }
  if (/^##\s+/.test(ln)) {
    const t = ln.replace(/^##\s+/, '');
    const cid = 'CHAP_' + modules.length;
    modules.push({ id: cid, text: t });
    body.push(H(t, 1, pendingBreak || HEAD_BREAK.has(i), cid)); pendingBreak = false; i++; continue;
  }
  if (/^#####\s+/.test(ln)) { body.push(H(ln.replace(/^#####\s+/, ''), 4, pendingBreak || HEAD_BREAK.has(i))); pendingBreak = false; i++; continue; }
  if (/^####\s+/.test(ln)) { body.push(H(ln.replace(/^####\s+/, ''), 3, pendingBreak || HEAD_BREAK.has(i))); pendingBreak = false; i++; continue; }
  if (/^###\s+/.test(ln)) { body.push(H(ln.replace(/^###\s+/, ''), 2, pendingBreak || HEAD_BREAK.has(i))); pendingBreak = false; i++; continue; }
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
  if (/^(解释：|公司差异：)/.test(ln.trim())) {    // SD-33 学习解释（灰色小字）/ SD-34 公司差异（标签加粗）
    let t = ln.trim();
    const exp = /^解释：/.test(t);
    { const m = t.match(/^(解释：|公司差异：)([^——]{1,30})——\s*(?:<[^>]+>)*\2[：:]/); if (m) t = t.replace(m[2] + '——', ''); }
    body.push(new Paragraph({
      children: exp ? runs('<small>' + t + '</small>') : runs('<strong>公司差异：</strong>' + t.slice(5), { size: 18 }),
      spacing: { before: 20, after: 100 }, indent: { left: 200 }
    }));
    i++; continue;
  }
  if (/^出处：/.test(ln.trim())) {                 // 表后出处（手册佐证）：灰色小字，紧跟表格
    const t = ln.trim().replace(/<[^>]+>/g, '');
    body.push(new Paragraph({
      children: [new TextRun({ text: t, font: { ascii: EN, eastAsia: CN }, size: 16, color: GRAY })],
      spacing: { before: 20, after: 140 }, indent: { left: 200 }
    }));
    i++; continue;
  }
  if (/^(详见\s|来源：)/.test(ln.trim())) {     // 回查入口行 / 速查区标题下的出处行：小号灰字
    const under = /^来源：/.test(ln.trim()) || (COMPACT && /^详见\s/.test(ln.trim()));
    body.push(new Paragraph({
      children: [new TextRun({ text: (() => {
        let t = ln.trim().replace(/^来源：部分内容待补来源$/, '来源：待补');
        /* 「详见 x.y A-n」内部不断行 */
        t = t.replace(/详见 ([^｜]+)$/, (m0, r) => '详见\u00A0' + r.replace(/ /g, '\u00A0'));
        return t; })(), font: { ascii: EN, eastAsia: CN }, size: 16, color: GRAY })],
      spacing: under ? { before: 0, after: 80 } : { before: 20, after: 160 },
      keepNext: true          // 来源 / 详见行与下文同页，不孤立在页底
    }));
    i++; continue;
  }
  if (COMPACT) {                                   // 速查区表外说明：与表格同字号，表前说明与表格同页
    let j = i + 1; while (j < src.length && !src[j].trim()) j++;
    body.push(P(ln, { size: 18, before: 40, after: 80, keepNext: j < src.length && /^<table/.test(src[j].trim()) }));
    i++; continue;
  }
  {                                                // 表格前的短说明（≤ 两行）与表格同页，不单独留在上一页
    let j = i + 1; while (j < src.length && !src[j].trim()) j++;
    const beforeTable = j < src.length && /^<table/.test(src[j].trim());
    let k = i - 1; while (k >= 0 && !src[k].trim()) k--;
    const beforeHead = j < src.length && /^#{2,6}\s+/.test(src[j]) && k >= 0 && /^#{2,6}\s+/.test(src[k]);   // 夹在节标题与条目标题之间的节首短说明：与条目同页，不让节标题孤悬页底
    const len = ln.replace(/<[^>]+>/g, '').length;
    body.push(P(ln, { keepNext: (beforeTable && len <= 120) || (beforeHead && len <= 200) }));
  }
  i++;
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
  ...(modules.length > 1 ? [
    new Paragraph({
      children: [new TextRun({ text: '章节快速跳转', font: { ascii: EN, eastAsia: CN }, size: 22, bold: true })],
      spacing: { before: 420, after: 60 }
    }),
    new Paragraph({
      children: [new TextRun({ text: '点击章名直接跳到该章第一页；页码在 Word 中按 Ctrl+A 再按 F9 可刷新。',
                               font: { ascii: EN, eastAsia: CN }, size: 18, color: GRAY })],
      spacing: { before: 0, after: 160 }
    }),
    chapterJump(modules)
  ] : []),
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
        size: PORTRAIT ? { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT }
                       : { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
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
