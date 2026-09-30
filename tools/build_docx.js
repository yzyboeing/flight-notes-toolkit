/*  build_docx.js —— markdown 源 → 带目录的 A4 横版 Word
 *  用法：node build_docx.js <input.md> [output.docx]
 *  支持：# / ## / ### 标题、段落、- 项目符号、1. 编号、表格、```代码块```、> 引用、**加粗**、`代码`
 *  目录按大纲自行生成（PAGEREF 域），打开文档后全选按 F9 或「更新域」刷新页码。
 */
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, TableOfContents,
  WidthType, ShadingType, BorderStyle, AlignmentType, VerticalAlign, HeadingLevel,
  PageBreak, Footer, Header, SimpleField, TabStopType, SectionType, PageNumber, PageOrientation, Bookmark, PageReference, InternalHyperlink, TableLayoutType, LeaderType, TableBorders
} = require('docx');
const fs = require('fs');

const SRC = process.argv[2] || 'flight_theory_notes_prompt_v5.md';
const OUT = process.argv[3] || 'prompt.docx';

// 字体：LibreOffice 导出 PDF 时，苹方 / Helvetica Neue 的粗体和数学符号会被拆成多种替代字体。
// 默认改用本机自带且同时覆盖中英文、数学符号和真实粗体的 Songti SC；需要时仍可用环境变量覆盖。
const MAIN = process.env.DOC_FONT || 'Songti SC',
      CN   = process.env.DOC_FONT_CN   || MAIN,
      EN   = process.env.DOC_FONT_EN   || MAIN,
      MONO = process.env.DOC_FONT_MONO || 'Menlo';
const FF = { ascii: EN, hAnsi: EN, eastAsia: CN, cs: EN };
const FF_CN = { ascii: CN, hAnsi: CN, eastAsia: CN, cs: CN };
const FF_MONO = { ascii: MONO, hAnsi: MONO, eastAsia: CN, cs: MONO };
/* DOC_PALETTE：整体配色方案比选（2026-09-29）。未设置＝现行方案。
   key＝数值字色，keyBg＝数值底色（荧光笔），note/noteBar＝注解条底色/竖条，warn＝警示条底色，hdr＝表头底，alt＝斑马纹，tail＝表后「注：」行底色 */
const PALETTES = {
  hl:   { key: '000000', keyBg: 'FFEB9C', note: 'F3F3F3', noteBar: '808080', warn: 'FDECEA', hdr: 'E7E6E6', alt: 'FAFAFA', tail: '' },
  blue: { key: '0B5CAD', keyBg: '',       note: 'EEF4FB', noteBar: '0B5CAD', warn: 'FDECEA', hdr: 'DCE6F1', alt: 'F7F9FC', tail: '' },
  teal: { key: '00806A', keyBg: 'E3F4EF', note: 'F1F7F5', noteBar: '00806A', warn: 'FDECEA', hdr: 'E4EDEA', alt: 'F8FBFA', tail: '' },
};
/* SD-68（2026-09-29 用户选 B）：默认配色＝亮蓝统一方案。
   DOC_PALETTE=green 可回到 SD-65 墨绿方案；hl / teal 为比选留档。 */
PALETTES.green = { key: '1B6B4C', keyBg: '', note: 'EAF1F8', noteBar: '2E74B5', warn: 'FDECEA', hdr: 'EBEBEB', alt: 'F7F7F7', tail: '' };
const PAL = PALETTES[process.env.DOC_PALETTE || 'blue'] || PALETTES.blue;
const GRAY = '595959', LINE = 'BFBFBF', ALT = PAL.alt || 'F7F7F7', CODE = 'F2F2F2';
const RED = 'C00000';
/* ---------- 彩色版配色（2026-09-29 用户：整本按彩色设计）----------
   只用三个色相，各自含义固定，避免「五颜六色但看不出层级」：
     红  C00000  限制值、禁令、必须句（正文强调色，仅此一处用红）
     蓝  1F4E79 / 2E74B5  结构色：标题、分隔线、注解条（不表示危险，只表示层级）
     琥珀 BF8F00  前提 / 适用条件（提醒「先看条件再看结论」）
   底色一律取同色相的最浅一档，打印不糊、iPad 上长时间看不刺眼。 */
const INK   = '1F4E79',            // 标题与粗分隔线（深蓝）
      INK2  = '2E74B5',            // 次级线条、注解竖条（中蓝）
      HDR   = PAL.hdr || 'EBEBEB',            // 表头底（中性灰，结构不占色相）
      NOTE_BG = PAL.note || 'EAF1F8',          // 注 / 补充说明底
      PRIORITY_BG = 'FFF2CC',                  // 最高优先级整行 / 整项底色（浅黄）
      PRE_BG  = 'FBF2E3',          // 前提 / 适用条件底
      PRE_BAR = 'BF8F00',          // 前提竖条（琥珀）
      PRIORITY_BAR = 'BF8F00';     // 最高优先级竖条（深黄）
const PREMISE = NOTE_BG;   /* SD-64：注解条由三色收敛为两色，前提条并入注解蓝 */
/* 印刷选项（2026-09-29 用户：黑白双面印刷）
   DOC_BW=1：黑白印刷——限制值由红色改为黑色加粗 + 下划线，警告条改黑色（红色在黑白印刷中与黑色几乎无法区分）
   DOC_DUPLEX=1：双面印刷——镜像页边距（内侧加宽装订）、奇偶页页眉页脚左右对调（页码在外侧）、封面与每章从右页（奇数页）开始 */
const BW = process.env.DOC_BW === '1';
const DUPLEX = process.env.DOC_DUPLEX === '1';
/* 黑白印刷时所有底色回落为灰阶，线条回落为黑 / 深灰 */
const C = (color, bw) => (BW ? bw : color);
const HDR_F = C(HDR, 'D9D9D9'), NOTE_F = C(NOTE_BG, 'FFFFFF'), PRIORITY_F = C(PRIORITY_BG, 'EDEDED'), PRE_F = C(NOTE_BG, 'EDEDED');
const NOTE_BAR = C(PAL.noteBar || INK2, GRAY), PRE_BAR_C = C(PAL.noteBar || INK2, GRAY), PRIORITY_BAR_C = C(PRIORITY_BAR, '000000');
/* SD-63「颜色让给内容」：标题与线条一律黑 / 深灰，层级靠字号、字重、线条、缩进表达；
   蓝 INK 专门留给「要记的数值」（行内 <b>），红留给限制与禁令（<em>）。 */
const H1_C = '000000', H2_C = '000000', H3_C = '000000';
const H1_LINE = '000000', H2_LINE = '404040', H3_BAR = C('595959', '404040');
/* SD-65 数值强调色＝墨绿 1B6B4C（2026-09-29 用户选定）。
   深蓝 1F4E79 在正文字号下与黑太接近；橙与红在同页并存时几乎分不开，会稀释红色的警示作用。
   墨绿与黑、红、链接蓝三者都分得开，形成四路分工：黑＝正文、红＝禁令、绿＝数值、蓝＝可点击。
   可用 DOC_KEYCOLOR 覆盖；黑白版回落为黑体加粗。 */
const KEY = '1B6B4C';
const KEY_C = C(process.env.DOC_KEYCOLOR || PAL.key || KEY, '000000');
const KEY_BG = process.env.DOC_KEYBG || PAL.keyBg || '';   // 数值强调的底色（荧光笔式），默认无
/* 2026-09-30 用户：页边距再小一点，让表格不那么拥挤——左右上下 720 DXA（12.7mm），页眉页脚距页边 360（双面仍内宽外窄） */
const M_IN = DUPLEX ? 1000 : 720, M_OUT = DUPLEX ? 620 : 720;   // 内侧 / 外侧页边距（DXA）
const M_TOP = 720, M_BOT = 720, M_HDR = 360, M_FTR = 360;
const HIDE_TBD = process.env.SHOW_TBD !== '1';   // 成品默认不显示〔待补来源〕（用户要求：表格与正文内不标来源）
// DOC_PORTRAIT=1：竖版 A4（iPad 阅读版）；默认横版
const PORTRAIT = process.env.DOC_PORTRAIT === '1';
/* DOC_SINGLE=1：单册（单章成书，如「机型基础知识速查」）——封面重排、章首页改为纯目录页 */
const SINGLE = process.env.DOC_SINGLE === '1';
const PAGE_W = PORTRAIT ? 11906 : 16838, PAGE_HT = PORTRAIT ? 16838 : 11906;
const TOTAL = PAGE_W - M_IN - M_OUT;   // 表格最大宽度＝版心宽度，随页边距自适应（2026-09-30 用户）
const SC = (w) => Math.round(w * TOTAL / 14400);

/* ---------- 行内解析：**bold** `code` <em>红</em> <strong>粗</strong> ---------- */
const unesc = (t) => t.replace(/&lt;/g,'<').replace(/&gt;/g,'>')
  .replace(/&nbsp;/g,' ').replace(/&quot;/g,'"').replace(/&amp;/g,'&');
function runs(text, o = {}) {
  text = unesc(text);
  const out = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`|<em>[\s\S]*?<\/em>|<strong>[\s\S]*?<\/strong>|<b>[\s\S]*?<\/b>|<small>[\s\S]*?<\/small>|〔待补来源〕|<br\s*\/?>)/g;
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
    t = t.replace(/【[^】]*】/g, m0 => m0.replace(/ /g, '\u00A0')).replace(/(\d{1,2}:\d{2}) ([—–-]) (\d{1,2}:\d{2})/g, '$1\u00A0$2\u00A0$3').replace(/(\d{1,2}:\d{2}) ?～ ?(\d{1,2}:\d{2})/g, '$1⁠～⁠$2');
    t = t.replace(/([<>≤≥=＜＞≈约±]) (?=[\d−\-+.])/g, '$1\u00A0').replace(/(\d) (?=(kg|ft|kt|nm|NM|m|km|psi|psid|fpm|min|s|h|%|°|℃)(?![A-Za-z]))/g, '$1\u00A0');
    t = t.replace(/以(?=[上下内外])/g, '以\u2060');   // 避免「以上 / 以下 / 以内 / 以外」在单元格行尾拆成孤字
    out.push(new TextRun({
      text: t,
      font: kind === 'code' ? FF_MONO : circ ? FF_CN : FF,
      size: grayK && !o.inTable ? (o.size || 20) - 2 : (o.size || 20),
      bold: kind === 'bold' || kind === 'red' || kind === 'key' || kind === 'graybold' || o.bold,
      underline: (BW && kind === 'red' && !o.noRed) ? {} : undefined,
      ...(kind === 'key' && KEY_BG && !BW ? { shading: { type: ShadingType.CLEAR, color: 'auto', fill: KEY_BG } } : {}),
      color: (kind === 'red' && !o.noRed) ? (BW ? '000000' : RED) : (kind === 'key' ? KEY_C : (kind === 'code' ? '9C2A00' : (grayK ? (o.inTable ? '4A4A4A' : GRAY) : (o.color || '000000'))))
    }));
  };
  /* 嵌套标记：<em> 与 <strong> 可互相嵌套，红色优先（红色本身已是粗体） */
  const walk = (s, kind) => {
    const r = /(\*\*[^*]+\*\*|`[^`]+`|<em>[\s\S]*?<\/em>|<strong>[\s\S]*?<\/strong>|<b>[\s\S]*?<\/b>|<small>[\s\S]*?<\/small>|〔待补来源〕|<br\s*\/?>)/g;
    let l = 0, mm;
    while ((mm = r.exec(s)) !== null) {
      push(s.slice(l, mm.index), kind);
      const tk = mm[0];
      if (tk.startsWith('**')) walk(tk.slice(2, -2), kind === 'red' ? 'red' : (/^gray/.test(kind || '') ? 'graybold' : 'bold'));
      else if (tk.startsWith('`')) push(tk.slice(1, -1), 'code');
      else if (tk.startsWith('<em>')) {
        /* SD-75：颜色由源标记的语义决定，不再仅因含数字就自动改蓝。
           <em>＝选定的边界 / 警戒 / 关键动作（红）；<b>＝选定的记忆值（蓝）；
           <strong>＝一般黑粗。这样同一句里可按逻辑有选择地安排红蓝重点。 */
        const inner = tk.slice(4, -5);
        walk(inner, /^gray/.test(kind || '') ? 'graybold' : 'red');
      }
      /* <b>…</b>＝要记的数值：深蓝加粗（SD-63）。红色优先，灰色解释段内不变蓝 */
      else if (/^<b>/.test(tk)) walk(tk.slice(3, -4), kind === 'red' ? 'red' : (/^gray/.test(kind || '') ? 'graybold' : 'key'));
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
  return out.length ? out : [new TextRun({ text: '', font: FF, size: 20 })];
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
let AFTER_H1 = false;
const NO_SEC_BREAK = !!process.env.NO_SEC_BREAK;   // 速查区单独成册等场合可关掉「每节另起一页」
function H(text, level, brk, forceId) {
  const sizes = { 1: 30, 2: 24, 3: 21, 4: 20 };
  const id = forceId || (level <= 2 ? bmk(text) : null);
  text = unesc(String(text));
  const isItem = /^(\d+\. |[A-Z]-\d+\u3000)/.test(text);
  const secBreak = level === 2 && /^\d+\.\d+[\s\u3000]/.test(text) && !AFTER_H1 && !NO_SEC_BREAK;
  AFTER_H1 = level === 1;
  const HC = { 1: H1_C, 2: H2_C, 3: H3_C, 4: H3_C }[level];
  const tr = new TextRun({ text, font: FF, size: sizes[level], bold: true, color: HC });
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : level === 3 ? HeadingLevel.HEADING_3 : HeadingLevel.HEADING_4,
    children: id ? [new Bookmark({ id, children: [tr] })] : [tr],
    spacing: { before: level === 1 ? 320 : 220, after: level === 1 ? 140 : 100 },
    /* 印刷版层级（2026-09-29）：章标题下粗黑线；节标题下细线；条目标题左侧竖条 */
    border: level === 1 ? { bottom: { style: BorderStyle.SINGLE, size: 12, color: H1_LINE, space: 6 } }
          : level === 2 ? { bottom: { style: BorderStyle.SINGLE, size: 6, color: H2_LINE, space: 4 } }
          : (level === 3 && isItem) ? { left: { style: BorderStyle.SINGLE, size: 18, color: H3_BAR, space: 6 } } : undefined,
    indent: (level === 3 && isItem) ? { left: 60 } : undefined,
    keepNext: true,                                  // 标题永远与下文同页
    /* 速查区竖版：排版预检发现会被拆页的表，其条目标题另起一页（BREAK_BEFORE=21,35）；
       印刷版：每个 x.y 节另起一页（紧跟章标题的第一节除外） */
    pageBreakBefore: !!brk || secBreak || (COMPACT && BREAKS.has((String(text).match(/^(\d+)\. /) || [])[1])) || itemBreak(text)
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
          children: [new TextRun({ text: l || ' ', font: FF_MONO, size: 17 })],
          spacing: { before: 0, after: 0, line: 250 }
        }))
      })]
    })]
  });
}


/* ---------- 模块速查导航表（页码用 PAGEREF 域，F9 刷新） ---------- */
/* ---------- 目录页下方的「章节快速跳转」表：点章名跳到该章首页 ---------- */
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
        children: [new TextRun({ text: t, font: FF, size: 18, bold: true })] })],
      { w: W[i], fill: HDR_F }))
  });
  const body = rows.map((r, ri) => {
    const fill = ri % 2 ? ALT : 'FFFFFF';
    const id = 'SEC_' + r.id.replace(/\./g, '_');
    return new TableRow({
      cantSplit: true,
      children: [
        cell([new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 4, after: 4, line: 205 },
          children: [new TextRun({ text: r.id, font: FF, size: 18, bold: true })] })], { w: W[0], fill }),
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
        cls: (at.match(/class="([^"]*)"/) || [, ''])[1],
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
let PROBE = false, TBL_IDX = -1, KEEP_LAST = false;
/* KEEP_FORCE 文件（fit_fix.py 维护）：一行一个表签名＝强制整表同页；「~N:签名」＝整表同页并按第 N 级压缩（SD-80）；「!签名」＝已放弃（fit_fix 自用） */
const KEEP_FORCE = new Set(), SHRINK = new Map(), WFIX = new Map();
if (process.env.KEEP_FORCE && fs.existsSync(process.env.KEEP_FORCE))
  for (const l0 of fs.readFileSync(process.env.KEEP_FORCE, 'utf8').split('\n')) {
    const l = l0.trim(); if (!l || l.startsWith('!')) continue;
    const w = l.match(/^W:(\d+):(\d+):(.+)$/);   // SD-85 实测加宽：W:列号:加宽DXA:签名
    if (w) { if (!WFIX.has(w[3])) WFIX.set(w[3], []); WFIX.get(w[3]).push([+w[1], +w[2]]); continue; }
    if (/^WB?:/.test(l)) continue;
    const m = l.match(/^~(\d):(.+)$/);
    if (m) { SHRINK.set(m[2], +m[1]); KEEP_FORCE.add(m[2]); } else KEEP_FORCE.add(l);
  }
const TBL_DUMPS = [];
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
  const isFull = r => (/note|warn/.test(r.cls) && r.cells.length === 1) || r.cls.includes('premise');   // 前提行不论几格都移出表格
  const rowText = r => r.cells.map(c => String(c.text)).filter(x => x.trim()).join('<br>');
  const lines = r => rowText(r).split(/<br\s*\/?>/).reduce((a, sg) => a + Math.max(1, Math.ceil(visC(sg) * 96 / Math.max(900, tw - 180))), 0);
  /* 通栏前提行一律移到表前作正文段落；较长的通栏注释行（约 40 字以上）或窄表里要折 4 行以上的注释移到表后。
     警示行（warn）保留在表内。用户要求：文字单独列出，下面附表格，不要为迁就表格把文字堆在一起（2026-09-29） */
  const fd = parsed.findIndex(r => !/hdr|premise/.test(r.cls));   // 第一行数据行（表头前后的前提行都算表前）
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
  const mids = parsed.map((r, k) => k).filter(k => isFull(parsed[k]) && /note|warn|premise/.test(parsed[k].cls) && k > fd && k < lastD
    && !spans.some(([a0, b0]) => a0 < k && b0 > k));
  if (!out.length && !mids.length) return [htmlTableCore(html)];
  const firstData = parsed.findIndex(r => !/hdr|premise/.test(r.cls));
  const pre = [], post = [];
  const trRe = /<tr([^>]*)>([\s\S]*?)<\/tr>/g;
  let idx = -1;
  const kept = html.replace(trRe, (all) => { idx++; const r = parsed[idx];
    if (!out.includes(r)) return all;
    (idx < firstData ? pre : post).push(r); return ''; });
  /* 表外注解段：警示 / 最高优先级用黄底，文字中的关键动作仍由 <em> 标红。 */
  const barOf = (cls) => /warn|priority/.test(cls) ? PRIORITY_BAR_C : cls.includes('premise') ? PRE_BAR_C : NOTE_BAR;
  const fillOf = (cls) => /warn|priority/.test(cls) ? PRIORITY_F : cls.includes('premise') ? PRE_F : NOTE_F;
  const para = (r) => new Paragraph({
    children: runs(String(r.cells[0].text).replace(/<br\s*\/?>/g, '\n'), { size: 18 }).map(x => x),
    spacing: { before: 60, after: 60, line: 270 }, indent: { left: 120, right: 80 },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: fillOf(r.cls) },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: barOf(r.cls), space: 8 } },
    keepNext: pre.includes(r)
  });
  const paras = (arr, kn) => arr.flatMap(r => balanceBr(rowText(r)).split(/<br\s*\/?>/).filter(x => x.trim()).map((sg, k, a) => new Paragraph({
    children: runs(sg.trim(), { size: 19 }),
    spacing: { before: k ? 20 : 100, after: k === a.length - 1 ? 100 : 20, line: 290 },
    /* 警示 / 前提 / 注 行移出表格后保留左侧竖条与淡底色 */
    indent: /warn|premise|note/.test(r.cls) ? { left: 120, right: 80 } : undefined,
    shading: /warn|premise|note/.test(r.cls) ? { type: ShadingType.CLEAR, color: 'auto', fill: fillOf(r.cls) } : undefined,
    border: /warn|premise|note/.test(r.cls) ? { left: { style: BorderStyle.SINGLE, size: 14, color: barOf(r.cls), space: 8 } } : undefined,
    keepNext: kn || pre.includes(r)
  })));
  /* SD-71：表末通栏注解行（移出表格成段落）与表格末行同页 */
  if (!mids.length) { const kl = KEEP_LAST; if (post.length) KEEP_LAST = true; const core = htmlTableCore(kept); KEEP_LAST = kl; return [...paras(pre), core, ...paras(post)]; }
  const rowsHtml = []; html.replace(trRe, (all) => { rowsHtml.push(all); return all; });
  const hdrIdx = parsed.map((r, k) => k).filter(k => k < firstData && parsed[k].cls.includes('hdr'));
  const hdrHtml = hdrIdx.map(k => rowsHtml[k]).join('\n');
  const partsOut = [...paras(pre)];
  let cur = [];
  /* SD-71：表中注解行把表拆成几段时，前一段末行、注解段与后一段首行互相「与下段同页」，不在注解处断开 */
  const flush = (bindNext) => { if (cur.length) { const kl = KEEP_LAST; if (bindNext) KEEP_LAST = true;
    partsOut.push(htmlTableCore('<table class="ftn">\n' + hdrHtml + '\n' + cur.join('\n') + '\n</table>')); KEEP_LAST = kl; } cur = []; };
  parsed.forEach((r, k) => {
    if (k < firstData || out.includes(r)) return;
    if (mids.includes(k)) { flush(true); partsOut.push(...paras([r], true)); return; }
    cur.push(rowsHtml[k]);
  });
  flush(post.length > 0);
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

  /* SD-78 语义分层（2026-09-30 用户：「结合语义去判断是否有多个并列的语义，通过加点可以让思路和语义更明确，不一定那么死板」）
     只改成品排版，notes_src 一字不动；放在列宽 / 表高估算之前，估算按拆分后的段落计算。
     ① 同一格多段（<br>）且语义并列 → 每段「•」；后段以「但 / 因此 / 即 / 其中 / 否则 / 此时 / 然后 / →」等开头的是延续，不拆；
     ② 一段里以「；」连起的几个完整长句（各 ≥ 约 10 字）→ 拆成并列「•」；
     ③ 「引语：子项1；子项2；…」→ 引语一行 + 「– 子项」缩进；
     已有 ①②③ / 编号 / 符号、表头、首列标签、注解行不动；【机型】开头的平行段不再加点（标签本身即标记）。 */
  {
    const MK_B = '', MK_C = '', MK_P = '', MK_Q = '\uE004';
    const plainOf = (t) => unesc(String(t).replace(/<[^>]+>/g, '')).trim();
    const pv = (t) => { let n = 0; for (const ch of plainOf(t)) n += /[⺀-鿿＀-￯]/.test(ch) ? 2 : /[A-Z]/.test(ch) ? 1.35 : /[a-z0-9]/.test(ch) ? 1.05 : /\s/.test(ch) ? 0.6 : 1.1; return n; };
    const balanced = (t) => ['em', 'strong', 'b', 'i', 'small', 'span', 'sup', 'sub', 'u'].every(g =>
      (String(t).match(new RegExp('<' + g + '(\\s[^>]*)?>', 'g')) || []).length === (String(t).match(new RegExp('</' + g + '>', 'g')) || []).length);
    const splitOut = (t, sepRe) => { const out = []; let cur = '';
      for (let i = 0; i < t.length; i++) { const ch = t[i];
        if (ch === '<') { const j = t.indexOf('>', i); if (j < 0) { cur += t.slice(i); break; } cur += t.slice(i, j + 1); i = j; continue; }
        cur += ch; if (sepRe.test(ch)) { out.push(cur); cur = ''; } }
      if (plainOf(cur)) out.push(cur); else if (out.length) out[out.length - 1] += cur;
      return out.every(balanced) ? out : null; };
    const CONT = /^(但|但是|因此|所以|即|其中|否则|此时|然后|随后|并且|而且|→|（|\()/;
    const STRUCT = /^([①-⑳]|\d+[.、)）]\s|[A-Z]-\d+|第 ?\d+ ?[条步]|注[：:]|[▪•·–—-]\s)/;
    /* 不分条的列（2026-09-30 用户，速查区第 23 条：「定义……根本不需要加圆点，只需要把这个定义居中」）：
       表头显式 col-center 的列、表头为「定义 / 含义 / 释义 / 概念」的列——一格就是一个完整概念，不拆 */
    const noSemCols = new Set(), bulletCols = new Set();   // 表头 class="col-bullet"：整列每格按项加「•」（SD-84 显式标注）
    { const hr = parsed.find(r => r.cls.includes('hdr'));
      if (hr) { const hri = parsed.indexOf(hr);
        hr.cells.forEach((hc, hk) => { if (hc.colspan !== 1) return;
          if (/(^|\s)col-center(\s|$)/.test(hc.cls || '') || /^(定义|含义|释义|概念)$/.test(plainOf(hc.text).replace(/\s+/g, '')))
            noSemCols.add(startCol[hri][hk]);
          if (/(^|\s)col-bullet(\s|$)/.test(hc.cls || '')) bulletCols.add(startCol[hri][hk]); }); } }
    parsed.forEach((r, ri) => {
      if (/hdr|note|premise|warn/.test(r.cls)) return;
      r.cells.forEach((c, k) => {
        if (c.head || startCol[ri][k] === 0 || noSemCols.has(startCol[ri][k])) return;
        if (bulletCols.has(startCol[ri][k]) && c.colspan === 1) {   // 显式整列加点：每段按「；」拆项（各 ≥ 约 6 字），延续句悬挂对齐
          const s1 = String(c.text).split(/<br\s*\/?>/).filter(x => plainOf(x));
          const p1 = plainOf(c.text).replace(/\s+/g, '');
          if (!s1.length || /^[—－\-–\/／无空×✕✓√?？…]*$/.test(p1) || s1.some(x => STRUCT.test(plainOf(x)))) return;
          const out1 = [];
          s1.forEach((x, xi) => { const pp = splitOut(x, /[；;]/);
            const pieces = pp && pp.length >= 2 && pp.every(y => pv(String(y).replace(/[；;，,。]/g, '')) >= 12) ? pp : [x];
            pieces.forEach(y => out1.push(((xi > 0 || out1.length) && CONT.test(plainOf(y)) ? MK_Q : MK_B) + y)); });
          c.text = out1.join('<br>'); return;
        }
        const segs = String(c.text).split(/<br\s*\/?>/).filter(x => plainOf(x));
        /* 源文件用「- 」写的子项（接在一句话之后）：统一排成「引语 + – 子项」，不再原样印出连字符 */
        const DASH = /^\s*[-–—]\s+/;
        if (segs.length >= 2 && segs.every(x => DASH.test(plainOf(x)))) {   // 整格每段都以「- 」开头：就是分条列表，排成「•」
          c.text = segs.map(x => MK_B + x.replace(/^\s*[-–—]\s+/, '')).join('<br>'); return;
        }
        if (segs.length >= 2 && !DASH.test(plainOf(segs[0])) && segs.slice(1).some(x => DASH.test(plainOf(x)))) {
          c.text = segs.map(x => DASH.test(plainOf(x)) ? MK_C + x.replace(/^\s*[-–—]\s+/, '') : MK_P + x).join('<br>'); return;
        }
        if (!segs.length || segs.some(s0 => STRUCT.test(plainOf(s0)))) return;
        /* 引语段 + 子项段（2026-09-30 用户，速查区第 54 条：「……超出以下任意限制：」前面加圆点，下面各限制加短线）：
           首段以「：」结尾、后面 ≥ 2 段，且后面各段都不再以「：」结尾、不是延续句 → 「• 引语」+「– 子项」 */
        if (segs.length >= 3 && /[：:]$/.test(plainOf(segs[0])) && pv(segs[0]) >= 4
            && segs.slice(1).every(x => !/[：:]$/.test(plainOf(x)) && !CONT.test(plainOf(x)))) {
          c.text = [MK_B + segs[0], ...segs.slice(1).map(x => MK_C + MK_C + x)].join('<br>'); return;
        }
        const items = [];
        for (const s0 of segs) {
          const ci = s0.search(/[：:]/);
          if (ci > 0) {
            const head = s0.slice(0, ci + 1), rest = s0.slice(ci + 1), hp = plainOf(head);
            const kids = balanced(head) && balanced(rest) ? splitOut(rest, /[；;]/) : null;
            if (kids && kids.length >= 2 && pv(hp) >= 4 && pv(hp) <= 60 && !/^注/.test(hp) && kids.every(x => pv(x) >= 6)) {
              items.push({ t: head, kind: 'intro', kids }); continue; }
          }
          const parts = splitOut(s0, /[；;]/);
          /* 整格 ≥ 约 35 字才按「；」拆（2026-09-30 用户，速查区第 34 条：两句短要求不加点） */
          if (parts && parts.length >= 2 && parts.every(x => pv(x) >= 20 && !CONT.test(plainOf(x))) && pv(s0) >= 70) { parts.forEach(x => items.push({ t: x, kind: 'item' })); continue; }
          items.push({ t: s0, kind: 'item' });
        }
        const parallel = items.length >= 2
          && items.slice(1).every(it => !CONT.test(plainOf(it.t)))
          && !(items[0].kind === 'item' && /[：:]$/.test(plainOf(items[0].t)))
          && items.every(it => it.kind === 'intro' || pv(it.t) >= 12)
          && items.some(it => it.kind === 'intro' || pv(it.t) > 20)
          && !items.every(it => it.kind === 'item' && /^【/.test(plainOf(it.t)));
        const hasIntro = items.some(it => it.kind === 'intro');
        if (!parallel && !hasIntro) return;
        const out = [];
        for (const it of items) {
          if (it.kind === 'intro') { out.push((parallel ? MK_B : MK_P) + it.t); it.kids.forEach(x => out.push((parallel ? MK_C + MK_C : MK_C) + x)); }
          else out.push((parallel ? MK_B : '') + it.t);
        }
        c.text = out.join('<br>');
      });
    });
    /* SD-83 同列统一（2026-09-30 用户：「既然……都加了『原点』，那我建议前面的……也去加『原点』，这样整个一列就会统一」）：
       一列里只要有格子分条加了「•」，同列其余有实质内容的数据格也加「•」——整格作一项，首段加点，后续段悬挂对齐（\uE004）。
       占位格（— / 无 / ✓ 等）、已有 ①② / 编号 / 符号的格、跨列格、首列、表头与注解行不动。 */
    const colB = new Set(), colShort = new Set();
    parsed.forEach((r, ri) => { if (/hdr|note|premise|warn/.test(r.cls)) return;
      r.cells.forEach((c, k) => { if (c.colspan !== 1) return; const t = String(c.text), p1 = plainOf(t).replace(/\s+/g, '');
        if (t.includes(MK_B)) colB.add(startCol[ri][k]);
        else if (p1 && !/^[—－\-–\/／无空×✕✓√?？…]*$/.test(p1) && pv(t) < 12 && !/[。．]$/.test(p1)) colShort.add(startCol[ri][k]); }); });   // 带句号的短句（「中断起飞。」）算句子，不算取值
    colShort.forEach(k => colB.delete(k));   // 有数值 / 短词格的列是「取值列」，不做同列统一（避免「• V2」「• 持续」）
    parsed.forEach((r, ri) => { if (/hdr|note|premise|warn/.test(r.cls)) return;
      r.cells.forEach((c, k) => {
        const col = startCol[ri][k], t = String(c.text);
        if (c.head || col === 0 || c.colspan !== 1 || !colB.has(col) || t.includes(MK_B)) return;
        if (t.includes(MK_P)) {   // 「引语 + – 子项」格在加点列里：引语改作「•」项，子项多缩一级，其余段悬挂对齐
          c.text = t.split(/<br\s*\/?>/).map(x => x.startsWith(MK_P) ? MK_B + x.slice(1) : x.startsWith(MK_C + MK_C) ? x
            : x.startsWith(MK_C) ? MK_C + x : (plainOf(x) ? MK_Q + x : x)).join('<br>');
          return;
        }
        if (/[\uE001-\uE004]/.test(t)) return;
        const segs = t.split(/<br\s*\/?>/).filter(x => plainOf(x));
        const p0 = plainOf(t).replace(/\s+/g, '');
        if (!segs.length || /^[—－\-–\/／无空×✕✓√?？…]*$/.test(p0) || segs.some(s0 => STRUCT.test(plainOf(s0)))) return;
        c.text = segs.map((s0, i) => (i === 0 ? MK_B : MK_Q) + s0).join('<br>');
        if (process.env.UNI_LOG && !PROBE) console.error('UNI\t' + p0.slice(0, 40));
      }); });
    /* 一列里加了点的格如果每格都只有一项（一个「•」、没有子项），点就没有意义（用户第 25、52 条：单句 / 短语居中不加点）→ 整列撤点 */
    { const cells = new Map();
      parsed.forEach((r, ri) => { if (/hdr|note|premise|warn/.test(r.cls)) return;
        r.cells.forEach((c, k) => { if (c.colspan !== 1) return; const col = startCol[ri][k];
          if (!cells.has(col)) cells.set(col, []); cells.get(col).push(c); }); });
      cells.forEach((cs, col) => { if (bulletCols.has(col)) return;
        const marked = cs.filter(c => String(c.text).includes(MK_B));
        if (marked.length >= 2 && marked.every(c => (String(c.text).match(/\uE001/g) || []).length === 1 && !/[\uE002\uE003\uE004]/.test(String(c.text))))
          marked.forEach(c => { c.text = String(c.text).replace(/[\uE001]/g, ''); }); }); }
  }

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
    /* 不可断开的词（【机型】标签、英文 / 数字串、时刻）计入列的最小宽度：表头与数据格都不许把它从中间拆开（2026-09-29 全书排版复查） */
    const tokMin = new Array(nCols).fill(0);
    parsed.forEach((r, ri) => { if (/note|premise|warn/.test(r.cls)) return;
      r.cells.forEach((c, ck) => { if (c.colspan !== 1) return; const ci5 = startCol[ri][ck];
        const plain = unesc(String(c.text).replace(/<br\s*\/?>/g, ' ').replace(/<[^>]+>/g, '')).replace(/〔待补来源〕/g, '');
        (plain.match(/\d{1,2}:\d{2}\s*[—–\-~～]\s*\d{1,2}:\d{2}|(?:【[^】]{1,12}】)?[A-Za-z0-9][A-Za-z0-9.\/\-:+%°]*|【[^】]{1,12}】/g) || []).forEach(tk => { tokMin[ci5] = Math.max(tokMin[ci5], Math.ceil((vis(tk) + 1) * (ci5 === 0 || r.cls.includes('hdr') ? 1.12 : 1.02))); }); }); });   // 首列与表头为粗体，略宽
    const full = need2.map((n, i2) => Math.max(n, hdrV[i2] <= 14 ? hdrV[i2] : hdrV[i2] * 0.6, tokMin[i2], 3));
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
      /* 「轻列」：非首列、内容很少（不到最重长句列的 1/5）却因一两句短语被撑宽的列（如速查区第 47 条「目视警戒」列、
         跨行合并的一句话）——允许折成两行，把宽度让给说明列，减少说明列的折行（用户 2026-09-29） */
      const heaviest = Math.max(1, ...longIdx.map(i2 => amount[i2]));
      const light = new Set(heavy ? longIdx.filter(i2 => i2 > 0 && demoted.has(i2) && amount[i2] <= 0.2 * heaviest) : []);   // 只在有长句列（用满页宽）的表里做
      const twoLine = i2 => Math.round(Math.max(14, Math.ceil(full[i2] / 2) + 1, tokMin[i2]) * CHAR) + EXTRA;
      if (process.env.W_LOG && light.size) console.error('LIGHT', [...light].join(','), String(parsed[0].cells.map(c => c.text).join('/')).replace(/<[^>]+>/g, '').slice(0, 40));
      longIdx.forEach(i2 => { w[i2] = light.has(i2) ? Math.min(oneLine(i2), twoLine(i2)) : Math.min(oneLine(i2), Math.round(TOTAL * 0.23), Math.round(Math.max(8, demoted.has(i2) ? headVis[i2] : 0) * CHAR) + EXTRA); });    // 起点：约 4 个汉字，其余交给贪心分配
      /* 长列下限：约 10 个汉字（放得下一行时取一行宽），避免某列被挤成每行几个字的细长条；轻列下限为两行宽 */
      longIdx.forEach(i2 => { w[i2] = Math.max(w[i2], light.has(i2) ? Math.min(oneLine(i2), twoLine(i2)) : Math.min(oneLine(i2), 28 * CHAR + EXTRA)); });
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
  /* SD-72（2026-09-29 用户，速查区第 40 条）：表格没用满版面宽度时，不许让任何格子（尤其表头）无谓折行。
     把剩余宽度回填给正在折行的列：表头折行的列优先，其次按「补到一行放下所需宽度」由小到大，
     只补到刚好一行放下为止（不把短表硬撑满宽，仍守「按内容定宽、不留空白」）。 */
  if (COMPACT || FIT_ALL) {
    const fineVis = (t) => { const s0 = unesc(String(t).replace(/<[^>]+>/g, '').replace(/〔待补来源〕/g, '')).replace(/\*\*/g, '').trim(); let n = 0;
      for (const ch of s0) n += /[⺀-鿿＀-￯]/.test(ch) ? 2 : /[A-Z]/.test(ch) ? 1.35 : /[a-z0-9]/.test(ch) ? 1.05 : /\s/.test(ch) ? 0.6 : 1.1; return n; };
    const U = 96;   // 9pt 下每视觉单位约 96 DXA（与列宽模型同一套字宽）
    const req = new Array(nCols).fill(0), hdrReq = new Array(nCols).fill(0);
    parsed.forEach((r, ri) => { if (/note|premise|warn/.test(r.cls)) return;
      r.cells.forEach((c, ck) => { if (c.colspan !== 1) return; const k = startCol[ri][ck];
        const bold = r.cls.includes('hdr') || c.head || k === 0;
        const need = Math.max(0, ...String(c.text).split(/<br\s*\/?>/).map(sg => Math.ceil(fineVis(sg) * U * (bold ? 1.08 : 1.03)) + 260));
        req[k] = Math.max(req[k], need); if (r.cls.includes('hdr')) hdrReq[k] = Math.max(hdrReq[k], need); }); });
    let slack = TOTAL - W.reduce((a2, b2) => a2 + b2, 0);
    const order = [...Array(nCols).keys()].filter(k => W[k] < req[k])
      .sort((x, y) => ((W[y] < hdrReq[y]) - (W[x] < hdrReq[x])) || ((req[x] - W[x]) - (req[y] - W[y])));
    let filled = 0;
    for (const k of order) { const gap = req[k] - W[k]; if (gap <= slack) { W[k] += gap; slack -= gap; filled++; } }
    if (process.env.W_LOG && filled) console.error('BACKFILL', filled, JSON.stringify(W), String(parsed[0].cells.map(c => c.text).join('/')).replace(/<[^>]+>/g, '').slice(0, 40));
    /* SD-84 末行孤字（2026-09-30 用户，速查区第 55 条「警」字单独占一行）：某格折行后末行只剩约 1～2 个字、且共 2～4 行时，
       从最宽的列匀出宽度（不超过其 12%），让该格少折一行 */
    const tail = new Array(nCols).fill(0);
    parsed.forEach((r, ri) => { if (/note|premise|warn/.test(r.cls)) return;
      r.cells.forEach((c, ck) => { if (c.colspan !== 1) return; const k = startCol[ri][ck];
        const bold = r.cls.includes('hdr') || c.head || k === 0, avail = W[k] - 260;
        String(c.text).split(/<br\s*\/?>/).forEach(sg => { const L = fineVis(sg.replace(/^[\uE001-\uE004]+/, '')) * 90; if (avail <= 0 || !L) return;   // 按宋体 9pt 实际字宽（汉字 180 DXA，粗体同宽）
          const lines = Math.ceil(L / avail), rem = L - (lines - 1) * avail;
          if (lines >= 2 && lines <= 4 && rem <= 4.4 * 90) tail[k] = Math.max(tail[k], Math.ceil(L / (lines - 1)) + 260 + 60 - W[k]);
          /* 短标签 / 表头（约 20 字以内）折行：能一行放下就给够一行（SD-72 延伸：表内其他列有富余时也匀过来） */
          if (lines >= 2 && L <= 20 * 180 && !/[\uE001-\uE004]/.test(sg)) tail[k] = Math.max(tail[k], Math.ceil(L) + 260 + 60 - W[k]); }); }); });
    for (let k = 0; k < nCols; k++) if (tail[k] > 0) {
      let j = -1; for (let q = 0; q < nCols; q++) if (q !== k && (j < 0 || W[q] > W[j])) j = q;
      if (j >= 0 && tail[k] <= 0.12 * W[j] && W[j] - tail[k] > W[k] + tail[k]) { W[k] += tail[k]; W[j] -= tail[k];
        if (process.env.W_LOG) console.error('TAILFIX', k, tail[k]); }
    }
    /* SD-85 按 PDF 实测加宽（fit_fix.py 写入）：从最宽的另一列匀出，不超过其 15% */
    { const sig0 = parsed.map(r => r.cells.map(c => unesc(String(c.text).replace(/<[^>]+>/g, ''))).join('')).join('').replace(/[^\p{L}\p{N}]/gu, '').slice(0, 80);
      for (const [k, dd] of (WFIX.get(sig0) || [])) { if (k >= nCols) continue;
        let j = -1; for (let q = 0; q < nCols; q++) if (q !== k && (j < 0 || W[q] > W[j])) j = q;
        const d0 = j >= 0 ? Math.min(dd, Math.floor(0.15 * W[j])) : 0;
        if (d0 > 0) { W[k] += d0; W[j] -= d0; } } }
  }
  if (process.env.W_LOG && COMPACT) console.error('W', nCols, JSON.stringify(W), String(parsed[0].cells.map(c => c.text).join('/')).slice(0, 40));

  /* ---- 自动缩排：估算表格高度，超过一页时逐级缩小字号，尽量整表放在同一页 ---- */
  const PAGE_H = PAGE_HT - M_TOP - M_BOT;   // 可用高度（DXA），随页边距自适应
  const BUDGET = PAGE_H - 900;          // 留出小节标题与段间距
  const estimate = (sz) => {
    const unit = 132 * sz / 20;         // 每「视觉单位」宽度
    let h = 0;
    parsed.forEach((r, ri) => {
      let maxH = Math.round(17.5 * sz);
      r.cells.forEach((c, ck) => {
        const ci = startCol[ri][ck];
        let w = 0;
        for (let k = 0; k < c.colspan; k++) w += W[Math.min(ci + k, nCols - 1)];
        let lines = 0, np = 0;
        for (const seg0 of String(c.text).split(/<br\s*\/?>/)) {
          /* SD-78：加点 / 子项有悬挂缩进，可用行宽变窄；每多一段多一份段前后间距 */
          const mk = (seg0.match(/^[\uE001-\uE004]+/) || [''])[0];
          const ind = (mk === '\uE001' || mk === '\uE004') ? 200 : mk === '\uE002\uE002' ? 560 : mk === '\uE002' ? 360 : 0;
          const perLine = Math.max(4, (w - 180 - ind) / unit);
          lines += Math.max(1, Math.ceil(vis(seg0.slice(mk.length).trim()) / perLine)); np++;
        }
        const ch = lines * Math.round(17.5 * sz) + Math.max(0, np - 1) * 40;
        if (ch > maxH) maxH = ch;
      });
      h += maxH + 120 + 40;   // 行高 + 单元格上下边距 + 段前后
    });
    return h;
  };
  let FS = 18, LN = 270, CM = 60;
  /* SD-80（2026-09-30 用户批准）：略超一页的表按级压缩后整表同页，级别由 fit_fix.py 按 PDF 实测逐级试出：
     1 级＝9pt 不变、收紧行距与单元格边距；2 级＝8.5pt；3 级＝8pt（底线，不再往下缩）。3 级仍放不下的照常分页。 */
  const tblText = parsed.map(r => r.cells.map(c => unesc(String(c.text).replace(/<[^>]+>/g, ''))).join('')).join('').replace(/[^\p{L}\p{N}]/gu, '');
  const tblSig = tblText.slice(0, 80);
  const shrinkLv = SHRINK.get(tblSig) || 0;
  if (shrinkLv) { FS = [18, 18, 17, 16][shrinkLv]; LN = shrinkLv === 3 ? 240 : 250; CM = 30; }
  else if (process.env.ALLOW_SHRINK && estimate(FS) > BUDGET) {   // 旧开关，仅手动调试用
    const fit = [17, 16].find(c => estimate(c) <= BUDGET);
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
  const longCols = new Set(), semanticCenterCols = new Set();
  const colN = [], colL = [], colV = [], colT = [], paraCells = [];
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
           R3 「段落型」格子左对齐：分条列举（≥ 3 行 ①② / 1. 2.），或估算排版后 ≥ 4 行，
              或去标记后总长 > 100（约 50 个汉字），或含句读（，。；）且总长 > 40（约 20 个汉字）。
              其余（数值、短语、一两行的短句）居中。
           R4 一列里段落型格子过半 → 整列左对齐（长文字列统一左齐）；否则整列居中，个别段落型格子单独左对齐。
           R5 「—」占位一律居中；所有格子纵向居中。 */
        const rawSegs = String(c.text).split(/<br\s*\/?>/).map(sg => unesc(sg.replace(/<[^>]+>/g, '')).replace(/〔待补来源〕/g, '').trim()).filter(Boolean);
        const enumN = rawSegs.filter(sg => /^([①-⑳]|\d+[.、)）]|[a-z][)）]|[A-Z]-\d+\s|第 ?\d+ ?[条步])/.test(sg)).length;
        const listy = rawSegs.length >= 3 && (enumN >= 2 || rawSegs.filter(sg => /[；;]$/.test(sg)).length >= 2);
        const hierarchyHint = rawSegs.length >= 3 && rawSegs.some((sg, si) =>
          si < rawSegs.length - 1 && /[：:]$/.test(sg) && !/[，。；]/.test(sg.replace(/[：:]$/, '')));
        const lines = segs.reduce((a2, v) => a2 + Math.max(1, Math.ceil(v / (per * 1.04))), 0);
        const total = segs.reduce((a2, v) => a2 + v, 0);
        const puncts = (rawSegs.join('').match(/[，。；]/g) || []).length;
        /* 2026-09-29 终检：防同列「锯齿」——多行且带句读的也算段落型；段落型占三分之一以上整列左齐 */
        /* 段落型：分条列举（≥ 3 行）、排版后 ≥ 4 行、或总长超过约 70 个汉字；三行以内的短句（如 1.4 C-1 俯仰 / 横滚方式）居中 */
        /* 2026-09-29 印刷版：带句读、约 20 字以上的整句，或约 50 字以上的长格，也算段落型（左齐）；
           长句居中在印刷品里显得散乱（2.2 C-2 例）。短语、数值、一两行短句仍居中（1.4 C-1 例） */
        const para = hierarchyHint || listy || lines >= 4 || total > 100 || (puncts >= 1 && total > 40);
        colN[ci2] = (colN[ci2] || 0) + 1;
        if (total > 30) colT[ci2] = (colT[ci2] || 0) + 1;   // 句子型格（约 15 字以上）
        if (para) { colL[ci2] = (colL[ci2] || 0) + 1; paraCells.push([ci2, c, hierarchyHint || listy || lines >= 4 || total > 100]); }
      });
    });
    /* 有段落型格子，且（段落型占三分之一以上，或多数格子是句子）→ 整列左齐，防同列锯齿 */
    for (let k3 = 0; k3 < nCols; k3++) if (colL[k3] && ((colL[k3] || 0) * 3 >= (colN[k3] || 1) || (colT[k3] || 0) * 2 > (colN[k3] || 1))) longCols.add(k3);
    /* 只要一列出现真正的段落 / 列举格，就把整列左齐；不再只改那个格子，
       避免同列短格居中、长格左齐形成锯齿。显式 col-center 可覆盖这一自动判定。 */
    paraCells.forEach(([k3, c, strong]) => { if (strong) longCols.add(k3); });
    /* 块索引 / 章索引的「条目」列一律左对齐（全书统一，不随条目多少变化） */
    { const hr = parsed.find(r => r.cls.includes('hdr'));
      if (hr) hr.cells.forEach((c, k4) => { if (/^(条目|本节条目|条目与主题)$/.test(unesc(String(c.text).replace(/<[^>]+>/g, '')).trim())) longCols.add(k4); }); }
    /* SD-75：原理 / 含义 / 方式等栏目若每格都是单一紧凑短句，即使自然折成 2～3 行也整列居中。
       这类文字承担「并列定义 / 机制」而非连续叙述；限制、说明、处置等长文列仍左对齐。 */
    { const hr = parsed.find(r => r.cls.includes('hdr'));
      const hri = hr ? parsed.indexOf(hr) : -1;
      const semanticHdr = /^(原理|工作原理|作用|方法|方式|状态|现象|结果|目的|含义|逻辑|动作与原理|触发条件|发生条件|进入条件|启动条件|激活条件|工作条件)$/;
      if (hr) hr.cells.forEach((hc, hk) => {
        if (hc.colspan !== 1) return;
        const ci4 = startCol[hri][hk];
        const ht = unesc(String(hc.text).replace(/<[^>]+>/g, '')).trim();
        if (!semanticHdr.test(ht)) return;
        let n4 = 0, ok4 = true;
        parsed.forEach((r, ri) => {
          if (/hdr|note|premise|warn/.test(r.cls)) return;
          r.cells.forEach((c, ck) => {
            if (startCol[ri][ck] !== ci4 || c.colspan !== 1 || c.head) return;
            const ss = String(c.text).split(/<br\s*\/?>/).map(x => unesc(x.replace(/<[^>]+>/g, '')).trim()).filter(Boolean);
            if (!ss.length) return;
            n4++;
            const joined = ss.join('');
            const enumN = ss.filter(x => /^([①-⑳]|\d+[.、)）]|[a-zA-Z][)）]|第 ?\d+ ?[条步])/.test(x)).length;
            const conditionColumn = /条件$/.test(ht);
            if (ss.length > (conditionColumn ? 3 : 2) || enumN >= 2 || (joined.match(/[，。；]/g) || []).length > 2 || vis(joined) > (conditionColumn ? 140 : 105)) ok4 = false;
          });
        });
        if (n4 && ok4) semanticCenterCols.add(ci4);
      });
    }
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
  /* SD-71（2026-09-29 用户）：一页放得下的表一律整表同页——宁可上一页留白，也不要页底只剩表头和一两行。
     （此前「超过半页即允许分页」的规则作废。）超过一页的大表才在行间分页，见下方 orphan 规则。 */
  /* 估算值贴近一页上限（> 95%）的表不整表绑定：估算有几个百分点误差（标题占位、换字体后行高变化），
     绑定后放不下时排版软件会在末行前硬断、留下孤行（2026-09-30 Songti SC 实测：3.8 分类表估 9325 / 预算 9566，实际放不下） */
  /* SD-79 两遍排版：第一遍按估算（留 12% 余量）；成品 PDF 里「断开但两截加起来放得进一页」的表，由 fit_fix.py 记下签名，
     第二遍强制整表同页（KEEP_FORCE 文件，一行一个签名）。签名＝表内文字只留字母数字后的前 80 字。 */
  if (!PROBE && process.env.TBL_DUMP) TBL_DUMPS.push({ sig: tblSig, text: tblText, ratio: +(estimate(FS) / BUDGET).toFixed(3), W: W.slice(),
    br: parsed.map(r => r.cells.map(c => unesc(String(c.text).replace(/<br\s*\/?>/g, '\u0001').replace(/<[^>]+>/g, '')).replace(/[^\p{L}\p{N}\u0001]/gu, '').replace(/\u0001/g, '|')).join('')).join('') });
  const keepTogether = estimate(FS) <= BUDGET * 0.88 || KEEP_FORCE.has(tblSig);   // SD-78 实测：Songti SC 下个别表实际比估算高约 12%（4.20 A/P 可用性表 估 8575 / 实 ≈ 9640），留 15% 余量
  if (process.env.FIT_LOG) console.error('TBL rows=%d est=%d fs=%d keep=%s', parsed.length, estimate(FS), FS, keepTogether);

  /* 顶部连续的通栏前提行 + 表头行一起作「重复标题行」（Word 要求标题行从第一行起连续） */
  let lead = 0;
  while (lead < parsed.length && (parsed[lead].cls.includes('hdr') || parsed[lead].cls.includes('premise'))) lead++;
  if (!parsed.slice(0, lead).some(r => r.cls.includes('hdr'))) lead = 0;
  /* 防孤行（SD-71 修订）：大表首尾各保留若干行同页，但要按数据行数缩放，保证中间至少留一处可断开的位置；
     每一截至少 2 行（行少而高的块索引表，原「首 3 尾 3」等于全表绑定，放不下时只能在末行前硬断） */
  const dataN = parsed.length - lead;
  const orphHead = lead + (dataN >= 8 ? 2 : 1);          // 首截：表头 + 至少 2 行数据（大表 3 行）
  const orphTail = dataN >= 5 ? 3 : 2;                    // 尾截：至少 3 行（数据行少时 2 行）
  /* 并列对比表（如「系统 A 供压组件 | 系统 B 供压组件」）：首列不是标签列，不加粗 */
  const hdrRow = parsed.find(r => r.cls.includes('hdr'));
  /* 表头可用 class="col-left" / class="col-center" 显式声明整列语义。
     这是列级规则，不是逐格特例：适用于「定义」等应整列居中的内容，
     以及「说明 / 条件 / 结果」等应整列左齐的内容。表头本身仍一律居中。 */
  const forcedLeftCols = new Set(), forcedCenterCols = new Set();
  if (hdrRow) {
    const hri = parsed.indexOf(hdrRow);
    hdrRow.cells.forEach((hc, hk) => {
      if (hc.colspan !== 1) return;
      const ci4 = startCol[hri][hk];
      if (/(^|\s)col-left(\s|$)/.test(hc.cls || '')) forcedLeftCols.add(ci4);
      if (/(^|\s)col-center(\s|$)/.test(hc.cls || '')) forcedCenterCols.add(ci4);
    });
  }
  /* SD-84 同列统一·单句列居中（2026-09-30 用户，速查区第 24 条「渗漏场景这一列都居中」、第 25 条「第一列『动作』这一列都需要居中，
     因为后面都居中了」）：一列（含首列）每个数据格都只有一段、是一句话（句中无「。；」）、不超过约 40 字、没有分条符号 → 整列居中；
     覆盖 col-left 与「长句左齐」判定。多段、多句、列举、分条的格仍左齐。 */
  const sentCols = new Set();
  {
    const pvS = (t) => { let n = 0; for (const ch of unesc(String(t).replace(/<[^>]+>/g, '')).trim()) n += /[\u2E80-\u9FFF\uFF00-\uFFEF]/.test(ch) ? 2 : 1.1; return n; };
    const ok = new Array(nCols).fill(true), seen = new Array(nCols).fill(0);
    parsed.forEach((r, ri) => { if (/hdr|note|premise|warn/.test(r.cls)) return;
      r.cells.forEach((c, ck) => { if (c.head) return; const k = startCol[ri][ck];
        const t = String(c.text), pl = unesc(t.replace(/<[^>]+>/g, '')).trim();
        if (/^[—－\-–\/／无空×✕✓√?？…]*$/.test(pl.replace(/\s+/g, ''))) return;   // 占位格不参与判定
        const body = pl.replace(/[。．]$/, '');
        /* 短词列举（2026-09-30 用户，速查区第 50 条「触发音响」：都是短词、不是句子 → 居中）：顿号 / 逗号 / 「或」分隔、每项 ≤ 约 15 字宽，不论总长 */
        const terms = body.split(/[、，,]|\s或\s/).map(x => x.trim()).filter(Boolean);
        const termList = terms.length >= 2 && terms.every(x => pvS(x) <= 30);
        const bad = /<br\s*\/?>/.test(t) || /[\uE001-\uE004]/.test(t) || /[。]/.test(body) || (pvS(t) > 80 && !termList) || /^([①-⑳]|\d+[.、)）]\s)/.test(pl);
        if (c.colspan !== 1) { if (bad) for (let j = k; j < k + c.colspan; j++) ok[j] = false; return; }   // 跨列格只在「不是单句」时否决所跨各列
        seen[k]++;
        if (bad) ok[k] = false; }); });
    for (let k = 0; k < nCols; k++) if (ok[k] && seen[k] >= 2) sentCols.add(k);
    if (process.env.SENT_LOG && !PROBE && sentCols.size) console.error('SENT\t' + [...sentCols].join(',') + '\t' + parsed[0].cells.map(c => unesc(String(c.text).replace(/<[^>]+>/g, ''))).join('|').slice(0, 50));
  }
  /* SD-78（2026-09-30 用户，速查区第 5 条）：整列每一格都只有一段、且按最终列宽一行排得下时，整列居中 */
  const oneLineCols = new Set();
  {
    const fv = (t) => { let n = 0; for (const ch of unesc(String(t).replace(/<[^>]+>/g, '')).trim()) n += /[\u2E80-\u9FFF\uFF00-\uFFEF]/.test(ch) ? 2 : /[A-Z]/.test(ch) ? 1.35 : /[a-z0-9]/.test(ch) ? 1.05 : /\s/.test(ch) ? 0.6 : 1.1; return n; };
    const okc = new Array(nCols).fill(true), seen = new Array(nCols).fill(0);
    parsed.forEach((r, ri) => { if (/hdr|note|premise|warn/.test(r.cls)) return;
      r.cells.forEach((c, ck) => { if (c.colspan !== 1 || c.head) return; const k = startCol[ri][ck]; seen[k]++;
        const segs = String(c.text).split(/<br\s*\/?>/).filter(x => x.replace(/<[^>]+>/g, '').trim());
        if (segs.length !== 1 || fv(segs[0]) * 96 * FS / 18 + 180 > W[k]) okc[k] = false; }); });
    for (let k = 1; k < nCols; k++) if (okc[k] && seen[k] >= 2) oneLineCols.add(k);
  }
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
    const isPriority = r.cls.includes('priority');
    /* 黄色只用于显式标记的最高优先级整行 / 整项；红色仍只用于行内限制、警戒和关键动作。 */
    const fill = isHdr ? HDR_F : (isWarn || isPriority) ? PRIORITY_F : isPre ? PRE_F : isNote ? NOTE_F : (ri % 2 ? ALT : 'FFFFFF');
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
      const rawParas = String(c.text).split(/<br\s*\/?>/);
      const plainParas = rawParas.map(seg => unesc(seg.replace(/<[^>]+>/g, '')).trim());
      const contentParas = plainParas.filter(Boolean);
      const enumN2 = contentParas.filter(p => /^([①-⑳]|\d+[.、)）]|[a-zA-Z][)）]|第 ?\d+ ?[条步])/.test(p)).length;
      const parentFlags = plainParas.map((p, pi) => {
        const body = p.replace(/[：:]$/, '');
        return pi < plainParas.length - 1 && /[：:]$/.test(p) && vis(p) <= 70 && !/[，。；]/.test(body);
      });
      const semMarked = /[\uE001-\uE004]/.test(String(c.text));
      const hierarchy = !semMarked && !isHdr && !c.head && rawParas.length >= 3 && parentFlags.some(Boolean)
        && !plainParas.slice(1).some(p => /^[①-⑳]\s*/.test(p));
      /* SD-78（2026-09-30 用户，速查区第 8 / 12 / 15 条）：同一格内 ≥ 2 条彼此独立的并列句（原文以 <br> 分开）→ 每条前加「•」、悬挂缩进、左齐。
         已有编号（①②③ / 1. / A-1）、【机型】开头、首段以「：」引出（SD-75 父子层级）、括注续行的，不加；只有一句的格子不加。 */
      const bulletCell = false && !hierarchy && !isHdr && !c.head && !labelCol && !isNote && !isPre && !isWarn
        && contentParas.length >= 2 && contentParas.length === plainParas.length
        && !contentParas.some(p => /^([①-⑳]|\d+[.、)）]\s|【|[A-Z]-\d+|第 ?\d+ ?[条步]|注[：:]|[（(]|[▪•·–—-]\s)/.test(p))
        && !/[：:]$/.test(contentParas[0])
        && contentParas.some(p => vis(p) > 20)
        && contentParas.every(p => vis(p) >= 12);   // 用户 2026-09-30：加点只限长的、语义并列的句子——每条 ≥ 约 6 个汉字、至少一条 > 约 10 个汉字；「20s」这类标签 + 说明、短状态词不加   // 有特别短的段（「20s」这类时刻 / 标签，< 约 4 个汉字）时是「标签 + 说明」，不是并列句，不加点   // 每条都很短（≤ 约 8 个汉字，如「ALTN 电门亮 / ON 指示可见」）时不加点，保持居中
      if (process.env.HIER_LOG && hierarchy) console.error('HIERARCHY', TBL_IDX, ci + 1, contentParas[0] || '');
      /* 同一语义列保持同一种对齐，避免短格居中、长格左齐形成锯齿。
         父子层级本身依靠悬挂缩进表达，始终左齐；独立占位符仍居中。 */
      const forceLeft = c.colspan === 1 && forcedLeftCols.has(ci);
      const forceCenter = c.colspan === 1 && forcedCenterCols.has(ci);
      const center = isHdr || c.head || placeholder || (!hierarchy && forceCenter) || (!hierarchy && !semMarked && (() => { for (let j = ci; j < ci + c.colspan; j++) if (!sentCols.has(j)) return false; return true; })())
        || (!forceLeft && !hierarchy && ((labelShort && !longCols.has(ci)) || (c.colspan === 1 && (semanticCenterCols.has(ci) || oneLineCols.has(ci)))
          || (labelCol && !longCols.has(ci)) || ((COMPACT || FIT_ALL)
          ? (!isNote && !isPre && !isWarn &&
               (c.colspan === 1 ? !longCols.has(ci) : shortCell(c.text)))
          : ((c.colspan === 1 && (centerCols.has(ci) || narrowSet.has(ci)))
             || (isFirstCol && c.colspan === 1)))));
      /* 首列序号格（只有一个圈码）：圈码字形在 Word 里常回退到无粗体的字体，改排为粗体阿拉伯数字 */
      const serial = isFirstCol && /^\s*(<strong>)?\s*[\u2460-\u2473]\s*(<\/strong>)?\s*$/.test(String(c.text));
      if (serial) c = Object.assign({}, c, { text: String(String(c.text).replace(/<[^>]+>/g, '').trim().charCodeAt(0) - 0x245F) });
      const paras = rawParas.map((seg, pi) => {
        const hasParent = parentFlags.slice(0, pi + 1).some(Boolean);
        const mk = (seg.match(/^[\uE001-\uE004]+/) || [''])[0]; seg = seg.slice(mk.length);
        const prefix = hierarchy ? (parentFlags[pi] ? '▪ ' : (hasParent ? '– ' : '')) : mk === '\uE001' ? '• ' : mk.startsWith('\uE002') ? '– ' : '';
        return new Paragraph({
          /* 第一列（项目名 / 标签列）加粗，让表头行与首列都醒目；首列为长句列时不加粗 */
          children: runs(prefix + seg.trim(), { inTable: true, noRed: labelCol, bold: isHdr || c.head || labelCol, size: FS }),
          spacing: { before: 20, after: 20, line: LN },
          indent: hierarchy ? { left: parentFlags[pi] ? 180 : 360, hanging: 140 } : mk === '\uE001' ? { left: 200, hanging: 200 } : mk === '\uE004' ? { left: 200 } : mk === '\uE002\uE002' ? { left: 560, hanging: 180 } : mk === '\uE002' ? { left: 360, hanging: 180 } : undefined,
          keepNext: ((isHdr || isPre) && ri < parsed.length - 1) || (keepTogether && ri < parsed.length - 1) || (!keepTogether && ri < parsed.length - 1 && (ri < orphHead || ri >= parsed.length - orphTail)) || (tailNote && ri >= lastData && ri < parsed.length - 1) || (KEEP_LAST && ri === parsed.length - 1),
          alignment: (hierarchy || semMarked) ? undefined : (center ? AlignmentType.CENTER : undefined)
        });
      });
      const borders = {
        top: { style: BorderStyle.SINGLE, size: 2, color: LINE },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
        left: (isWarn || isPriority) ? { style: BorderStyle.SINGLE, size: 14, color: PRIORITY_BAR_C }
             : isPre  ? { style: BorderStyle.SINGLE, size: 14, color: PRE_BAR_C }
             : isNote ? { style: BorderStyle.SINGLE, size: 14, color: NOTE_BAR }
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
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: ri === 0 ? HDR_F : (ri % 2 ? ALT : 'FFFFFF') },
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
/* 正文段落里以 <br> 结尾的连续行合并成同一段（段内软回车），不再每行另起一段、多出空行（2026-09-29 全书排版复查） */
const src = (() => {
  const L0 = rawSrc.split(/\r?\n/), out = [];
  let inT = false, inC = false;
  const special = t => /^(#{1,6}\s|<table|\||```|%%|---|解释：|注：|出处：|公司差异：|来源：|详见\s|<tr|<\/table)/.test(t);
  for (let k = 0; k < L0.length; k++) {
    let t = L0[k];
    if (/^```/.test(t)) inC = !inC;
    if (!inC && /^\s*<table/.test(t)) inT = true;
    if (inT || inC) { out.push(t); if (/<\/table>/.test(t)) inT = false; continue; }
    while (/<br\s*\/?>\s*$/.test(t) && k + 1 < L0.length) {
      const nx = L0[k + 1];
      if (!nx.trim() || special(nx.trim())) { t = t.replace(/(<br\s*\/?>\s*)+$/, ''); break; }
      t = t.replace(/\s+$/, '') + nx.trim(); k++;
    }
    t = t.replace(/(<br\s*\/?>\s*)+$/, '');
    out.push(t);
  }
  return out;
})();
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

/* ---------- 全书大纲（2026-09-29 封面 / 目录 / 章首页改版）：章、节、速查主题块 ---------- */
const OUTLINE = [];
{
  let comp = false, inCode = false, q = 0, cur = null;
  for (const s of src) {
    if (/^```/.test(s)) { inCode = !inCode; continue; }
    if (inCode) continue;
    if (/^%%COMPACT%%\s*$/.test(s)) { comp = true; continue; }
    if (/^%%ENDCOMPACT%%\s*$/.test(s)) { comp = false; continue; }
    let m;
    if ((m = s.match(/^##\s+(.*)$/))) { cur = { id: 'CHAP_' + OUTLINE.length, text: m[1].trim(), desc: '', secs: [] }; OUTLINE.push(cur); continue; }
    if (cur && (m = s.match(/^%%CHAPDESC%%\s*(.*)$/))) { cur.desc = unesc(m[1].trim()); continue; }
    if (cur && (m = s.match(/^###\s+(.*)$/))) {
      const t = m[1].trim(), id = bmk(t) || (comp ? 'QRB_' + (q++) : null);   // 速查区主题块无编号，另起书签 QRB_n
      if (id) cur.secs.push({ id, text: unesc(t) });
    }
  }
}
/* 只有按「第X章」分章的笔记（全书 / 单章分册）才排章首页、总览封面和自生成目录；整理规范等其他文档仍用旧版式 */
const BOOK = OUTLINE.length > 0 && OUTLINE.every(ch => /^第.{1,3}章[\s\u3000]/.test(unesc(ch.text)));
const NB = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const splitChap = (t) => { const m = unesc(String(t)).match(/^(第.{1,3}章)[\s　]+(.*)$/); return m ? [m[1], m[2]] : ['', unesc(String(t))]; };
const splitSec = (t) => { const m = String(t).match(/^(\d+(?:\.\d+)*)[\s　]+(.*)$/); return m ? [m[1], m[2]] : ['', String(t)]; };
const chapNo = (cn) => { const k = '零一二三四五六七八九'.indexOf(String(cn).charAt(1)); return k >= 0 ? String(k).padStart(2, '0') : ''; };
/* 目录 / 章首页的一行：整行（编号 + 标题 + 点线 + 页码）都在一个内部超链接里，点哪儿都能跳；页码用 PAGEREF，Word 中 F9 刷新 */
function tocLine(e, w, o = {}) {
  const NUMW = o.numW || 640;
  if (e.chap) return new Paragraph({
    keepNext: true,
    tabStops: [{ type: TabStopType.RIGHT, position: w - 60 }],
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'E6E6E6' },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: '000000', space: 4 } },
    indent: { left: 80 },
    spacing: { before: o.first ? 0 : 200, after: 60, line: 280 },
    children: [
      new InternalHyperlink({ anchor: e.id, children: [
        new TextRun({ text: e.cn + (e.cn ? '　' : '') + e.ct, font: FF, size: 22, bold: true, color: '000000' }),
        new TextRun({ text: '\t', size: 22 }), new PageReference(e.id) ] })
    ]
  });
  /* 说明行：带 id 时整行做成内部超链接（例：总目录里第零章下的「速查主题」跳到本章首页的主题清单） */
  /* 2026-09-30 用户：总目录第零章下的「速查主题」要让人一眼看出能点击跳转——
     → 箭头 + 下划线蓝字 + 「点击跳转」提示 + 点线连页码 */
  if (e.note && e.jump) return new Paragraph({
    keepLines: true, indent: { left: NUMW }, spacing: { before: 40, after: 40, line: 260 },
    tabStops: [{ type: TabStopType.RIGHT, position: w, leader: LeaderType.DOT }],
    children: [new InternalHyperlink({ anchor: e.id, children: [
      new TextRun({ text: '→ ', font: FF, size: 20, bold: true, color: H2_C }),
      new TextRun({ text: e.note, font: FF, size: 20, bold: true, color: H2_C, underline: {} }),
      new TextRun({ text: '　点击跳转', font: FF, size: 16, color: GRAY }),
      new TextRun({ text: '\t', size: 20 }), new PageReference(e.id) ] })]
  });
  if (e.note) {
    const nr = new TextRun({ text: e.note, font: FF, size: 18, color: e.id ? H2_C : GRAY });
    return new Paragraph({
      indent: { left: NUMW }, spacing: { before: 20, after: 20, line: 240 },
      children: [e.id ? new InternalHyperlink({ anchor: e.id, children: [nr] }) : nr]
    });
  }
  return new Paragraph({
    keepLines: true,
    tabStops: [...(e.num ? [{ type: TabStopType.LEFT, position: NUMW }] : []),
               { type: TabStopType.RIGHT, position: w, leader: LeaderType.DOT }],
    indent: e.num ? { left: NUMW, hanging: NUMW } : undefined,
    spacing: o.loose ? { before: 60, after: 60, line: 300 } : { before: 20, after: 20, line: 240 },
    children: [
      new InternalHyperlink({ anchor: e.id, children: [
        ...(e.num ? [new TextRun({ text: e.num + '\t', font: FF, size: 20, color: '000000' })] : []),
        new TextRun({ text: e.text, font: FF, size: 20, color: '000000' }),
        new TextRun({ text: '\t', size: 20 }), new PageReference(e.id) ] })
    ]
  });
}
/* 无框多栏排版表：cols = [[Paragraph…], …]，栏宽 colW，栏距 gap，整表居中 */
function colsTable(cols, colW, gap) {
  const B = { top: NB, bottom: NB, left: NB, right: NB };
  const cell = (ch, w) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: B,
    margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: ch.length ? ch : [new Paragraph({ children: [] })] });
  const widths = [], cells = [];
  cols.forEach((ch, k) => {
    if (k) { widths.push(gap); cells.push(cell([], gap)); }
    widths.push(colW); cells.push(cell(ch, colW));
  });
  return new Table({ rows: [new TableRow({ children: cells })], alignment: AlignmentType.CENTER, borders: TableBorders.NONE,
    layout: TableLayoutType.FIXED, width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths });
}
/* 单册目录页（DOC_SINGLE=1）：不出章序号与章名，整页就是一张目录——
   「目录」标题 + 双线 → 简介一行 → 主题两栏、点线连页码。 */
function singleToc(ch, brk) {
  const n = ch.secs.length;
  const out = [new Paragraph({ pageBreakBefore: true, spacing: { before: PORTRAIT ? 600 : 200, after: 0 }, children: [] })];
  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, keepNext: true,
    spacing: { before: 0, after: 120 },
    children: [new Bookmark({ id: ch.id, children: [new TextRun({ text: '目录', font: FF, size: 40, bold: true, characterSpacing: 60, color: H1_C })] })]
  }));
  const hr = (sz, col, after) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after, line: 20 },
    children: [new TextRun({ text: '', size: 2 })],
    border: { bottom: { style: BorderStyle.SINGLE, size: sz, color: col, space: 2 } } });
  out.push(hr(24, H1_LINE, 40), hr(4, C(INK2, '000000'), 0));
  /* 2026-09-29 用户：目录页不出简介行；主题单栏竖排 */
  out.push(new Paragraph({ spacing: { before: 0, after: PORTRAIT ? 400 : 240 }, children: [] }));
  if (n) {
    /* SD-71 只出横版：横版一页高度排不下单栏 32 条，改为两栏（竖版仍单栏） */
    const nc = PORTRAIT ? 1 : 2, GAP = 900;
    const colW = nc === 1 ? Math.min(TOTAL, 7600) : Math.floor((TOTAL - GAP) / 2);
    const per = Math.ceil(n / nc), cols = [];
    /* 主题本身没有编号，这里按顺序补 01…N，便于口头指引「看第 12 条」 */
    const lines = ch.secs.map((sec, k) => {
      const [, t] = splitSec(sec.text);
      return { id: sec.id, num: String(k + 1).padStart(2, '0'), text: t };
    });
    for (let c = 0; c < nc; c++) cols.push(lines.slice(c * per, (c + 1) * per).map(l => tocLine(l, colW, { numW: 560, loose: true })));
    out.push(colsTable(cols, colW, GAP));
  }
  return out;
}
/* 章首页：大号章序号 → 章名（Heading 1，页眉 STYLEREF 取它）+ 粗线 → 本章简介 → 本章各节与页码 */
function chapterOpener(ch, brk) {
  const [cn] = splitChap(ch.text);
  const no = chapNo(cn), n = ch.secs.length;
  if (SINGLE) return singleToc(ch, brk);
  const out = [new Paragraph({ pageBreakBefore: !!brk, spacing: { before: 0, after: PORTRAIT ? 2400 : 500 }, children: [] })];
  if (no) out.push(new Paragraph({ spacing: { before: 0, after: 0 }, keepNext: true,
    children: [new TextRun({ text: no, font: FF, size: 150, bold: true, color: C('C9D8E8', '000000') })] }));
  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_1, keepNext: true,
    children: [new Bookmark({ id: ch.id, children: [new TextRun({ text: unesc(ch.text), font: FF, size: 44, bold: true, color: H1_C })] })],
    spacing: { before: 60, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: H1_LINE, space: 8 } }
  }));
  if (ch.desc) out.push(new Paragraph({ spacing: { before: 60, after: 0, line: 340 },
    children: [new TextRun({ text: ch.desc, font: FF, size: 22, color: '404040' })] }));
  if (n) {
    const quick = ch.secs.every(s => /^QRB_/.test(s.id));
    out.push(new Paragraph({ keepNext: true, spacing: { before: PORTRAIT ? 900 : 480, after: 120 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C(INK2, '000000'), space: 4 } },
      children: [new TextRun({ text: quick ? '本章速查主题' : '本章内容', color: H1_C, font: FF, size: 20, bold: true, characterSpacing: 40 })] }));
    const nc = n <= 10 ? 1 : (PORTRAIT ? 2 : (n > 20 ? 3 : 2)), GAP = 600;
    const colW = nc === 1 ? Math.min(TOTAL, 9000) : Math.floor((TOTAL - GAP * (nc - 1)) / nc);
    const per = Math.ceil(n / nc), cols = [];
    for (let c = 0; c < nc; c++) cols.push(ch.secs.slice(c * per, (c + 1) * per).map(s => {
      const [num, t] = splitSec(s.text); return tocLine({ id: s.id, num, text: t }, colW, { numW: 600 });
    }));
    const tb = colsTable(cols, colW, GAP);
    out.push(tb);
  }
  return out;
}

const body = [];
const CHAP_IDX = [];   // 双面印刷：每章起点（分节，从奇数页开始）
const modules = [];          // 分册（## 级标题），与 OUTLINE 一一对应
let i = 0, docTitle = '', pendingBreak = false, QRB_N = 0;

while (i < src.length) {
  let ln = src[i];

  /* HTML 注释整行：不渲染（速查区用它保留「详见」给校验脚本，成品里不出现） */
  if (/^\s*<!--[\s\S]*?-->\s*$/.test(ln)) { i++; continue; }

  if (/^```/.test(ln)) {                       // 代码块
    const buf = []; i++;
    while (i < src.length && !/^```/.test(src[i])) buf.push(src[i++]);
    i++; body.push(codeBlock(buf)); body.push(P('', { before: 0, after: 40 })); continue;
  }
  if (/^\s*<table/.test(ln)) {                 // 内嵌 HTML 表格
    const buf = [];
    while (i < src.length && !/<\/table>/.test(src[i])) buf.push(src[i++]);
    buf.push(src[i++]);
    { let j = i; while (j < src.length && !src[j].trim()) j++;
      KEEP_LAST = j < src.length && /^(注：|出处：|解释：|公司差异：)/.test(src[j].trim()); }   // 表后注 / 出处 / 解释与表格末行同页
    const tbs = htmlTable(buf.join('\n')).filter(Boolean);
    KEEP_LAST = false;
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
  if (/^%%CHAPDESC%%/.test(ln)) { i++; continue; }   // 章简介：已排在章首页
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
    const ch = OUTLINE[modules.length - 1];
    if (DUPLEX) CHAP_IDX.push(body.length);
    if (BOOK) {
      body.push(...chapterOpener(ch, !DUPLEX && (pendingBreak || HEAD_BREAK.has(i))));
      AFTER_H1 = true; pendingBreak = true;        // 章首页单独一页，正文从下一页开始
    } else { body.push(H(t, 1, !DUPLEX && (pendingBreak || HEAD_BREAK.has(i)), cid)); pendingBreak = false; }
    i++; continue;
  }
  if (/^#####\s+/.test(ln)) { body.push(H(ln.replace(/^#####\s+/, ''), 4, pendingBreak || HEAD_BREAK.has(i))); pendingBreak = false; i++; continue; }
  if (/^####\s+/.test(ln)) { body.push(H(ln.replace(/^####\s+/, ''), 3, pendingBreak || HEAD_BREAK.has(i))); pendingBreak = false; i++; continue; }
  if (/^###\s+/.test(ln)) {
    const t3 = ln.replace(/^###\s+/, '').trim();
    body.push(H(t3, 2, pendingBreak || HEAD_BREAK.has(i), (COMPACT && !bmk(t3)) ? 'QRB_' + (QRB_N++) : undefined));
    pendingBreak = false; i++; continue;
  }
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
  /* 表后说明块（2026-09-29 印刷版统一）：注 / 公司差异 / 解释 / 出处 同一缩进、同一字号、左侧细线，连续几行成一块 */
  const TAIL = { indent: { left: 240 }, border: { left: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 8 } } };
  if (/^(注：|解释：|公司差异：)/.test(ln.trim())) {    // SD-33 解释（灰色）/ SD-34 公司差异、注（标签加粗）
    let t = ln.trim();
    const exp = /^解释：/.test(t);
    { const m = t.match(/^(解释：|公司差异：|注：)([^——]{1,30})——\s*(?:<[^>]+>)*\2[：:]/); if (m) t = t.replace(m[2] + '——', ''); }
    const lab = t.match(/^(注：|公司差异：)/);
    body.push(new Paragraph({
      children: exp ? runs('<small>' + t + '</small>') : runs('<strong>' + lab[1] + '</strong>' + t.slice(lab[1].length), { size: 18 }),
      spacing: { before: 10, after: 50, line: 280 }, ...TAIL
    }));
    i++; continue;
  }
  if (/^出处：/.test(ln.trim())) {                 // 表后出处（手册佐证）：灰色小字，紧跟表格
    const t = unesc(ln.trim().replace(/<[^>]+>/g, ''));   // 实体（&lt; &gt;）还原
    body.push(new Paragraph({
      children: [new TextRun({ text: t, font: FF, size: 16, color: GRAY })],
      spacing: { before: 10, after: 50, line: 260 }, ...TAIL
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
        return t; })(), font: FF, size: 16, color: GRAY })],
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
const PREFACE_SIGNATURE = process.env.DOC_PREFACE_SIGNATURE || AUTHOR;
const editionLabel = (() => {
  const value = String(process.env.DOC_EDITION || '').trim();
  if (!value) return '';
  return /^版本号(?:\s|$)/.test(value)
    ? value.replace(/^版本号\s*/, '版本号 ')
    : `版本号 ${value}`;
})();
const noticeRuns = (size = 18, color = '404040') => {
  const value = String(process.env.DOC_NOTICE || '');
  const m = value.match(/^(特别提示[：:])([\s\S]*)$/);
  return m
    ? [new TextRun({ text: m[1], font: FF, size, bold: true, color: '000000' }),
       new TextRun({ text: m[2], font: FF, size, color })]
    : [new TextRun({ text: value, font: FF, size, color })];
};
const rule = (o) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: o.before || 0, after: o.after || 0, line: 20 },
  children: [new TextRun({ text: '', size: 2 })],
  border: { bottom: { style: BorderStyle.SINGLE, size: o.size || 6, color: o.color || LINE, space: 2 } }
});
/* 封面（2026-09-29 改版，黑白印刷）：上粗细双线 → 机型标签 → 书名 → 副标题 → 下细粗双线 → 各章总览 → 作者 / 版次 → 细框声明 */
const BOXW = PORTRAIT ? 8000 : 10000;
const coverTxt = (text, size, o = {}) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: o.before || 0, after: o.after || 0, line: o.line },
  outlineLevel: o.outline ? 0 : undefined,
  children: [new TextRun({ text, font: FF, size, bold: !!o.bold, color: o.color || '000000', characterSpacing: o.cs })]
});
function coverOverview() {
  if (!BOOK || OUTLINE.length < 2) return [];
  const W = [1100, BOXW - 1100];
  const B = (last) => ({ top: NB, left: NB, right: NB, bottom: last ? NB : { style: BorderStyle.SINGLE, size: 4, color: LINE } });
  const rows = OUTLINE.map((ch, k) => {
    const [cn, ct] = splitChap(ch.text), last = k === OUTLINE.length - 1;
    return new TableRow({ cantSplit: true, children: [
      new TableCell({ width: { size: W[0], type: WidthType.DXA }, borders: B(last), verticalAlign: VerticalAlign.CENTER,
        margins: { top: 20, bottom: 20, left: 0, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: chapNo(cn), font: FF, size: 36, bold: true })] })] }),
      new TableCell({ width: { size: W[1], type: WidthType.DXA }, borders: B(last), verticalAlign: VerticalAlign.CENTER,
        margins: { top: 20, bottom: 20, left: 200, right: 0 },
        children: [
          new Paragraph({ spacing: { line: 280 }, children: [
            new TextRun({ text: cn + (cn ? '　' : ''), font: FF, size: 20, color: GRAY }),
            new TextRun({ text: ct, font: FF, size: 24, bold: true }) ] }),
          ...(ch.desc && PORTRAIT ? [new Paragraph({ spacing: { before: 10, line: 240 }, children: [new TextRun({ text: ch.desc, font: FF, size: 16, color: GRAY })] })] : [])
        ] })
    ] });
  });
  return [new Table({ rows, alignment: AlignmentType.CENTER, borders: TableBorders.NONE, layout: TableLayoutType.FIXED,
                      width: { size: BOXW, type: WidthType.DXA }, columnWidths: W })];
}
const sideInd = Math.max(0, Math.round(((PORTRAIT ? 11906 : 16838) - M_IN - M_OUT - BOXW) / 2));
/* 单册封面（DOC_SINGLE=1）：题名组落在视觉中心偏上，版次与声明压到页面下部，
   声明不用方框，改细线 + 灰字，整页只有两组粗细双线作为骨架。 */
const NOTE_IND = sideInd + (PORTRAIT ? 900 : 1600);
/* 全书与速查册共用同一套封面提示：两条短灰线围住提示文字，版本号单独置于下方。
   不再使用四边框，避免封面底部像表格；版本号也不属于提示框内容。 */
const coverNoticeBlock = (before = 0) => [
  ...(process.env.DOC_NOTICE ? [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before, after: 0, line: 20 },
      indent: { left: NOTE_IND, right: NOTE_IND },
      children: [new TextRun({ text: '', font: FF, size: 2 })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '808080', space: 2 } } }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 140, after: 140, line: 290 },
      indent: { left: NOTE_IND, right: NOTE_IND },
      children: noticeRuns(18, '595959') }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0, line: 20 },
      indent: { left: NOTE_IND, right: NOTE_IND },
      children: [new TextRun({ text: '', font: FF, size: 2 })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '808080', space: 2 } } }),
  ] : []),
  ...(editionLabel ? [coverTxt(editionLabel, 18, { before: process.env.DOC_NOTICE ? 220 : before, color: '808080' })] : [])
];
/* 全书封面已有章节总览，底部空间比单册紧：用一个仅带上下边线的段落承载提示，
   再把版本号作为独立段落放在其下。视觉语言与单册一致，但不会把提示推到第 2 页。 */
const coverNoticeCompact = (before = 0) => [
  ...(process.env.DOC_NOTICE ? [new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before, after: 0, line: 260 },
    indent: { left: NOTE_IND, right: NOTE_IND },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: '808080', space: 5 },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: '808080', space: 5 }
    },
    children: noticeRuns(18, '595959')
  })] : []),
  ...(editionLabel ? [coverTxt(editionLabel, 18, { before: process.env.DOC_NOTICE ? 20 : before, color: '808080' })] : [])
];
const coverSingle = [
  new Paragraph({ spacing: { before: 0, after: PORTRAIT ? 2600 : 600 }, children: [] }),
  rule({ size: 24, color: '000000', after: 50 }),
  rule({ size: 4, color: '000000', after: 0 }),
  coverTxt(process.env.DOC_LABEL || 'B737-NG　/　B737-8', 24, { before: PORTRAIT ? 520 : 280, bold: true, color: '404040', cs: 80 }),
  coverTxt(docTitle, 72, { before: 140, after: 140, bold: true, outline: true }),
  coverTxt(process.env.DOC_SUBTITLE || '', 26, { after: PORTRAIT ? 520 : 280, color: GRAY, cs: 60 }),
  rule({ size: 4, color: '000000', after: 50 }),
  rule({ size: 24, color: '000000', after: 0 }),
  new Paragraph({ spacing: { before: 0, after: PORTRAIT ? 5200 : 1500 }, children: [] }),
  ...coverNoticeBlock(0),
];
const coverFull = [
  new Paragraph({ spacing: { before: 0, after: PORTRAIT ? 900 : 150 }, children: [] }),
  rule({ size: 24, color: '000000', after: 50 }),
  rule({ size: 4, color: '000000', after: 0 }),
  coverTxt(process.env.DOC_LABEL || 'B737-NG　/　B737-8', 24, { before: PORTRAIT ? 560 : 280, bold: true, color: '404040', cs: 80 }),
  coverTxt(docTitle, 76, { before: 120, after: 120, bold: true, outline: true }),
  coverTxt(process.env.DOC_SUBTITLE || '系统 · 运行 · 训练', 26, { after: PORTRAIT ? 560 : 120, color: GRAY, cs: 60 }),
  rule({ size: 4, color: '000000', after: 50 }),
  rule({ size: 24, color: '000000', after: 0 }),
  new Paragraph({ spacing: { before: 0, after: PORTRAIT ? 600 : 40 }, children: [] }),
  ...coverOverview(),
  ...(AUTHOR ? [coverTxt(AUTHOR, 22, { before: PORTRAIT ? 700 : 200, color: '404040' })] : []),
  ...coverNoticeCompact(AUTHOR ? 40 : (PORTRAIT ? 1200 : 40))
];
const cover = SINGLE ? coverSingle : coverFull;
/* 目录（2026-09-29 改版）：不用 Word TOC 域，按大纲自行生成——章为灰底粗体行（左粗竖条），节缩进、点线连页码；
   页码为 PAGEREF 域（Word 中全选 F9 刷新），标题可点击跳转。速查区只列章行，31 个主题清单在第零章首页。
   整章成栏：竖版每页一栏，横版每页两栏；一栏放不下的章换到下一栏。 */
function buildToc() {
  if (process.env.DOC_NOTOC === '1') return [];   /* 单章成册：章首页自带本章目录，不再出总目录页 */
  if (!BOOK) return [new TableOfContents('目录', { hyperlink: true, headingStyleRange: '1-2' })];
  const PER = PORTRAIT ? 1 : 2, CAP = PORTRAIT ? 44 : 27, GAP = 800;
  const colW = PORTRAIT ? Math.min(TOTAL, 9000) : Math.floor((TOTAL - GAP) / 2);
  const blocks = OUTLINE.map(ch => {
    const [cn, ct] = splitChap(ch.text);
    const quick = ch.secs.length && ch.secs.every(s => /^QRB_/.test(s.id));
    const ls = [{ chap: true, id: ch.id, cn, ct }];
    if (quick) ls.push({ note: '速查主题清单（共 ' + ch.secs.length + ' 项）', id: ch.id, jump: true });
    else ch.secs.forEach(s => { const [num, text] = splitSec(s.text); ls.push({ id: s.id, num, text }); });
    return ls;
  });
  const wt = ls => ls.reduce((a, l) => a + (l.chap ? 1.8 : 1), 0);
  const cols = []; let cur = [], cw = 0;
  blocks.forEach(b => { const w = wt(b); if (cur.length && cw + w > CAP) { cols.push(cur); cur = []; cw = 0; } cur = cur.concat(b); cw += w; });
  if (cur.length) cols.push(cur);
  const out = [];
  for (let pg = 0; pg * PER < cols.length; pg++) {
    /* 标题两页都写「总目录」（用户 2026-09-29）；只有第一页进 PDF 书签，避免重复 */
    out.push(new Paragraph({ pageBreakBefore: pg > 0 || (!DUPLEX && pg === 0), alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 },
      outlineLevel: pg ? undefined : 0,
      children: [new TextRun({ text: '总目录', font: FF, size: pg ? 28 : 40, bold: true, characterSpacing: pg ? 20 : 60, color: H1_C })] }));
    out.push(rule({ size: 24, color: H1_LINE, after: 40 }), rule({ size: 4, color: C(INK2, '000000'), after: 0 }));
    if (pg === 0) out.push(new Paragraph({ spacing: { before: 0, after: 120 }, children: [] }));
    const pc = cols.slice(pg * PER, pg * PER + PER).map(c => c.map((l, k) => tocLine(l, colW, { first: k === 0 })));
    out.push(pc.length === 1 && PORTRAIT ? colsTable(pc, colW, GAP) : colsTable(pc.length < PER ? pc.concat([[]]) : pc, colW, GAP));
  }
  return out;
}
/* 主题线索引（SD-61）：总目录之后、正文之前一页。文件由 DOC_TOPICINDEX 指定（gh-private/主题线索引.md）。
   内容用与正文相同的解析器渲染，因此可以直接写 <table class="ftn">。 */
function buildTopicIndex() {
  const tp = process.env.DOC_TOPICINDEX;
  if (!tp || !fs.existsSync(tp)) return [];
  const out = [new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 },
    outlineLevel: 0,
    children: [new TextRun({ text: '主题线索引', font: FF, size: 40, bold: true, characterSpacing: 60, color: H1_C })] })];
  out.push(rule({ size: 24, color: H1_LINE, after: 40 }), rule({ size: 4, color: C(INK2, '000000'), after: 0 }));
  out.push(new Paragraph({ spacing: { before: 0, after: 200 }, children: [] }));
  const lines = fs.readFileSync(tp, 'utf8').split(/\r?\n/);
  let k = 0;
  while (k < lines.length) {
    const ln = lines[k];
    if (/^\s*<table/.test(ln)) {
      const buf = [];
      while (k < lines.length && !/<\/table>/.test(lines[k])) buf.push(lines[k++]);
      buf.push(lines[k++]);
      out.push(...htmlTable(buf.join('\n')));
      out.push(new Paragraph({ spacing: { before: 0, after: 120 }, children: [] }));
      continue;
    }
    if (ln.trim()) out.push(new Paragraph({ spacing: { before: 60, after: 60, line: 300 },
      children: runs(ln.trim(), { size: 19 }) }));
    k++;
  }
  if (!DUPLEX) out.push(new Paragraph({ children: [new PageBreak()] }));
  return out;
}
const toc = [
  ...buildToc(),
  ...(DUPLEX ? [] : [new Paragraph({ children: [new PageBreak()] })])
];
/* 前言（SD-51）：封面之后、目录之前一页。文字来自 DOC_PREFACE 指向的文件（gh-private/前言.md），
   段落首行缩进两字；以「【特别提示】」开头的段落加粗；末尾右下角署名独立取 DOC_PREFACE_SIGNATURE，
   未设置时兼容回退到封面作者。 */
function buildPreface() {
  const pf = process.env.DOC_PREFACE;
  if (!pf || !fs.existsSync(pf)) return [];
  const paras = fs.readFileSync(pf, 'utf8').split(/\n\s*\n/).map(x => x.trim()).filter(Boolean);
  const out = [];
  out.push(new Paragraph({ pageBreakBefore: !DUPLEX, alignment: AlignmentType.CENTER, spacing: { before: PORTRAIT ? 600 : 200, after: 120 },
    outlineLevel: 0,
    children: [new TextRun({ text: '前言', font: FF, size: 40, bold: true, characterSpacing: 60, color: H1_C })] }));
  out.push(rule({ size: 24, color: H1_LINE, after: 40 }), rule({ size: 4, color: C(INK2, '000000'), after: 0 }));
  out.push(new Paragraph({ spacing: { before: 0, after: 360 }, children: [] }));
  const ind = PORTRAIT ? 600 : 1800;
  paras.forEach(t => {
    const tip = /^【特别提示】/.test(t);
    out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 200, line: 400 },
      indent: { left: ind, right: ind, firstLine: tip ? 0 : 440 },
      children: [new TextRun({ text: t, font: FF, size: 22, bold: tip })] }));
  });
  if (PREFACE_SIGNATURE) out.push(new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: PORTRAIT ? 2200 : 2400, after: 0 },
    indent: { right: ind },
    children: [new TextRun({ text: PREFACE_SIGNATURE, font: FF, size: 22, color: '404040' })]
  }));
  return out;
}
const preface = buildPreface();
const front = cover.concat(preface, toc, buildTopicIndex());

/* ---------- 分节与页眉页脚 ---------- */
const PAGE = {
  size: PORTRAIT ? { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT }
                 : { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
  margin: { top: M_TOP, bottom: M_BOT, left: M_IN, right: M_OUT, header: M_HDR, footer: M_FTR }
};
const CONTENT_W = (PORTRAIT ? 11906 : 16838) - M_IN - M_OUT;
const hdrFont = FF;
const emptyHF = () => ({ header: new Header({ children: [new Paragraph({ children: [] })] }),
                         footer: new Footer({ children: [new Paragraph({ children: [] })] }) });
/* 页眉：书名 + 当前章名（STYLEREF 域），下细线。单面：左书名右章名；双面：奇数页（右页）章名靠外（右），偶数页（左页）书名靠外（左） */
const hdrPara = (left, right) => new Paragraph({
  tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
  children: [
    ...(left ? [left()] : []),
    new TextRun({ text: '\t', size: 16 }),
    ...(right ? [right()] : [])
  ]
});
/* 页眉书名：页眉没有机型标签，写全机型（DOC_HEADER，缺省为书名） */
const titleRun = () => new TextRun({ text: process.env.DOC_HEADER || docTitle, font: hdrFont, size: 16, color: GRAY });
const chapRun = () => new TextRun({ children: [new SimpleField('STYLEREF "Heading 1"', '')], font: hdrFont, size: 16, color: GRAY });
const pageRun = () => new TextRun({ children: ['第 ', PageNumber.CURRENT, ' 页'], font: hdrFont, size: 18, color: GRAY });
const tocRun = () => new TextRun({ text: '目　录', font: hdrFont, size: 16, color: GRAY });
const prefRun = () => new TextRun({ text: '前　言', font: hdrFont, size: 16, color: GRAY });
const footPara = (align) => new Paragraph({ alignment: align, children: [pageRun()] });
let SECTIONS;
if (!DUPLEX) {
  SECTIONS = [{
    properties: { page: PAGE, titlePage: true },
    headers: { first: emptyHF().header, default: new Header({ children: [hdrPara(titleRun, chapRun)] }) },
    /* 封面不显示页码；正文页码在右下角，小五号（9pt），仅「第 X 页」 */
    footers: { first: emptyHF().footer, default: new Footer({ children: [footPara(AlignmentType.RIGHT)] }) },
    children: front.concat(body)
  }];
} else {
  const HF = {
    headers: { default: new Header({ children: [hdrPara(null, chapRun)] }),      // 奇数页（右页）：章名靠外侧
               even: new Header({ children: [hdrPara(titleRun, null)] }) },       // 偶数页（左页）：书名靠外侧
    footers: { default: new Footer({ children: [footPara(AlignmentType.RIGHT)] }),
               even: new Footer({ children: [footPara(AlignmentType.LEFT)] }) }
  };
  const e = emptyHF(), e2 = emptyHF();
  const starts = CHAP_IDX.length ? CHAP_IDX : [body.length];
  SECTIONS = [
    /* 封面：单独一节，无页眉页脚（背面空白页同样无页眉页脚） */
    { properties: { page: PAGE, titlePage: true },
      headers: { first: e.header, default: e.header, even: e2.header },
      footers: { first: e.footer, default: e.footer, even: e2.footer },
      children: cover },
    /* 前言：从右页开始，页眉写「前言」 */
    ...(preface.length ? [{ properties: { page: PAGE, type: SectionType.ODD_PAGE },
      headers: { default: new Header({ children: [hdrPara(null, prefRun)] }), even: new Header({ children: [hdrPara(prefRun, null)] }) },
      footers: HF.footers, children: preface }] : []),
    /* 目录：从右页开始；页眉外侧固定写「目录」（STYLEREF 在目录页会取到第零章，不用） */
    { properties: { page: PAGE, type: SectionType.ODD_PAGE },
      headers: { default: new Header({ children: [hdrPara(null, tocRun)] }), even: new Header({ children: [hdrPara(tocRun, null)] }) },
      footers: HF.footers, children: toc.concat(body.slice(0, starts[0])) },
    /* 每章一节，从右页开始；章首页（节内第一页）不排页眉页脚 */
    ...starts.map((st, k) => { const f = emptyHF();
      return { properties: { page: PAGE, type: SectionType.ODD_PAGE, titlePage: true },
               headers: { ...HF.headers, first: f.header }, footers: { ...HF.footers, first: f.footer },
               children: body.slice(st, k + 1 < starts.length ? starts[k + 1] : body.length) }; })
  ];
}

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
    default: { document: { run: { font: FF, size: 20 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        paragraph: { outlineLevel: 0 },
        run: { size: 30, bold: true, color: '000000', font: FF } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        paragraph: { outlineLevel: 1 },
        run: { size: 24, bold: true, color: '000000', font: FF } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        paragraph: { outlineLevel: 2 },
        run: { size: 21, bold: true, color: '000000', font: FF } },
      { id: 'Heading4', name: 'Heading 4', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        paragraph: { outlineLevel: 3 },
        run: { size: 20, bold: true, color: '000000', font: FF } }
    ]
  },
  evenAndOddHeaderAndFooters: DUPLEX,
  sections: SECTIONS
});

if (process.env.TBL_DUMP) fs.writeFileSync(process.env.TBL_DUMP, TBL_DUMPS.map(x => JSON.stringify(x)).join('\n') + '\n');
Packer.toBuffer(doc).then(async b => {
  {
    /* SD-71：rowSpan 生成的合并延续格是空段落 <w:p/>，不带「与下段同页」；LibreOffice 要求表内段落全都设了才整表不拆。
       同一行其他格已设 keepNext 时，给延续格补上。 */
    const JSZip0 = require('jszip');
    const z0 = await JSZip0.loadAsync(b);
    let dx = await z0.file('word/document.xml').async('string');
    let nfix = 0;
    dx = dx.replace(/<w:tr[ >][\s\S]*?<\/w:tr>/g, tr => {
      if (!tr.includes('<w:keepNext/>') || !tr.includes('w:vMerge w:val="continue"')) return tr;
      return tr.replace(/(<w:tc><w:tcPr><w:vMerge w:val="continue"\/>[\s\S]*?<\/w:tcPr>)<w:p\/>/g, (m0, pre) => { nfix++; return pre + '<w:p><w:pPr><w:keepNext/></w:pPr></w:p>'; });
    });
    if (nfix) { z0.file('word/document.xml', dx); b = await z0.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }); }
    if (process.env.FIT_LOG) console.error('vMerge keepNext 补 %d 格', nfix);
  }
  if (DUPLEX) {
    /* docx 库不支持镜像页边距：生成后在 settings.xml 里补 <w:mirrorMargins/>（内侧 = left，外侧 = right） */
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(b);
    let st = await zip.file('word/settings.xml').async('string');
    if (!/<w:mirrorMargins/.test(st)) {
      const after = /<w:(defaultTabStop|evenAndOddHeaders|characterSpacingControl|compat|updateFields|rsids|mathPr|themeFontLang|clrSchemeMapping|decimalSymbol|listSeparator|trackRevisions|doNotTrackMoves|documentProtection|autoFormatOverride|styleLockTheme|styleLockQFSet|defaultTableStyle|bookFoldPrinting|bookFoldRevPrinting|bookFoldPrintingSheets|gutterAtTop|hideSpellingErrors|hideGrammaticalErrors|activeWritingStyle|proofState|formsDesign|attachedTemplate|linkStyles|stylePaneFormatFilter|stylePaneSortMethod|documentType|mailMerge|revisionView)\b/;
      const m = st.match(after);
      st = m ? st.replace(m[0], '<w:mirrorMargins/>' + m[0]) : st.replace(/(<w:settings[^>]*>)/, '$1<w:mirrorMargins/>');
      zip.file('word/settings.xml', st);
    }
    b = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  }
  fs.writeFileSync(OUT, b); console.log('written ' + OUT + '  (' + body.length + ' blocks)' + (BW ? ' [黑白]' : '') + (DUPLEX ? ' [双面]' : ''));
});
