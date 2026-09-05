#!/usr/bin/env node
"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");

function loadDocx() {
  try {
    return require("docx");
  } catch (firstError) {
    const bundled = path.join(
      os.homedir(),
      ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/docx",
    );
    try {
      return require(bundled);
    } catch (_secondError) {
      process.stderr.write(
        "build_docx: cannot load the 'docx' package. Run with the Codex bundled Node runtime or make docx available through NODE_PATH.\n",
      );
      process.exit(1);
    }
  }
}

const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  HeadingLevel,
  PageBreak,
  PageNumber,
  PageOrientation,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableOfContents,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} = loadDocx();

const PAGE_WIDTH = 16838;
const PAGE_HEIGHT = 11906;
const PAGE_MARGIN = 567;
const CONTENT_WIDTH = PAGE_WIDTH - PAGE_MARGIN * 2;
// Hiragino Sans GB is a macOS CJK sans-serif family recognized by Word and
// common macOS renderers; Word can substitute an equivalent elsewhere.
const FONT = "Hiragino Sans GB";

function fail(message) {
  process.stderr.write(`build_docx: ${message}\n`);
  process.exit(1);
}

function decodeEntities(text) {
  const entities = {
    "&amp;": "&",
    "&lt;": "<",
    "&gt;": ">",
    "&quot;": '"',
    "&#39;": "'",
    "&nbsp;": " ",
  };
  return text.replace(/&(amp|lt|gt|quot|#39|nbsp);/g, (match) => entities[match] || match);
}

function plainText(html) {
  return decodeEntities(
    html
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div)>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/[ \t]+\n/g, "\n")
      .trim(),
  );
}

function parseFrontMatter(markdown) {
  if (!markdown.startsWith("---\n")) return { attributes: {}, body: markdown };
  const end = markdown.indexOf("\n---\n", 4);
  if (end === -1) fail("input has unclosed YAML front matter");
  const raw = markdown.slice(4, end);
  const attributes = {};
  for (const line of raw.split("\n")) {
    const match = line.match(/^([A-Za-z0-9_]+):\s*(.*)$/);
    if (!match) continue;
    let value = match[2].trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    attributes[match[1]] = value;
  }
  return { attributes, body: markdown.slice(end + 5) };
}

function parseAttributes(raw) {
  const attributes = {};
  const pattern = /([:\w-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let match;
  while ((match = pattern.exec(raw))) attributes[match[1]] = match[2] ?? match[3] ?? "";
  return attributes;
}

function visualLength(text) {
  let total = 0;
  for (const character of plainText(text)) total += character.codePointAt(0) > 255 ? 1.8 : 1;
  return Math.max(total, 1);
}

function richRuns(source, options = {}) {
  const normalized = source.replace(/<br\s*\/?>/gi, "\n");
  const pattern = /(<em>[\s\S]*?<\/em>|<strong>[\s\S]*?<\/strong>|\*\*[\s\S]*?\*\*)/gi;
  const runs = [];
  let cursor = 0;
  let match;

  function pushText(value, style = {}) {
    const decoded = plainText(value);
    const parts = decoded.split("\n");
    parts.forEach((part, index) => {
      if (index > 0) runs.push(new TextRun({ break: 1 }));
      if (part) {
        runs.push(new TextRun({
          text: part,
          font: FONT,
          size: options.size || 19,
          bold: Boolean(options.bold || style.bold),
          color: style.color || options.color || "000000",
        }));
      }
    });
  }

  while ((match = pattern.exec(normalized))) {
    pushText(normalized.slice(cursor, match.index));
    const token = match[0];
    if (/^<em>/i.test(token)) pushText(token.replace(/^<em>|<\/em>$/gi, ""), { bold: true, color: "C00000" });
    else if (/^<strong>/i.test(token)) pushText(token.replace(/^<strong>|<\/strong>$/gi, ""), { bold: true });
    else pushText(token.slice(2, -2), { bold: true });
    cursor = pattern.lastIndex;
  }
  pushText(normalized.slice(cursor));
  return runs.length ? runs : [new TextRun({ text: "", font: FONT, size: options.size || 19 })];
}

function parseTable(html) {
  const tableMatch = html.match(/<table\b([^>]*)>([\s\S]*?)<\/table>/i);
  if (!tableMatch) fail("malformed HTML table");
  const tableAttributes = parseAttributes(tableMatch[1]);
  const rows = [];
  const rowPattern = /<tr\b([^>]*)>([\s\S]*?)<\/tr>/gi;
  let rowMatch;
  while ((rowMatch = rowPattern.exec(tableMatch[2]))) {
    const rowAttributes = parseAttributes(rowMatch[1]);
    const cells = [];
    const cellPattern = /<(th|td)\b([^>]*)>([\s\S]*?)<\/\1>/gi;
    let cellMatch;
    while ((cellMatch = cellPattern.exec(rowMatch[2]))) {
      const attributes = parseAttributes(cellMatch[2]);
      cells.push({
        header: cellMatch[1].toLowerCase() === "th",
        attributes,
        html: cellMatch[3].trim(),
        rowSpan: Math.max(1, Number.parseInt(attributes.rowspan || "1", 10) || 1),
        columnSpan: Math.max(1, Number.parseInt(attributes.colspan || "1", 10) || 1),
      });
    }
    if (cells.length) rows.push({ attributes: rowAttributes, cells });
  }
  if (!rows.length) fail("HTML table contains no rows");
  return { attributes: tableAttributes, rows };
}

function mapGrid(rows) {
  const active = [];
  let columnCount = 0;
  for (const row of rows) {
    let column = 0;
    for (const cell of row.cells) {
      while ((active[column] || 0) > 0) column += 1;
      cell.columnStart = column;
      for (let offset = 0; offset < cell.columnSpan; offset += 1) {
        if (cell.rowSpan > 1) active[column + offset] = Math.max(active[column + offset] || 0, cell.rowSpan);
      }
      column += cell.columnSpan;
    }
    columnCount = Math.max(columnCount, column, active.length);
    for (let index = 0; index < active.length; index += 1) {
      if (active[index] > 0) active[index] -= 1;
    }
  }
  return Math.max(columnCount, 1);
}

function computeWidths(tableData) {
  const columnCount = mapGrid(tableData.rows);
  const totals = Array(columnCount).fill(0);
  const maxima = Array(columnCount).fill(1);
  const counts = Array(columnCount).fill(0);
  let totalText = 0;

  for (const row of tableData.rows) {
    for (const cell of row.cells) {
      const length = visualLength(cell.html);
      totalText += length;
      const share = length / cell.columnSpan;
      for (let offset = 0; offset < cell.columnSpan; offset += 1) {
        const index = cell.columnStart + offset;
        totals[index] += share;
        maxima[index] = Math.max(maxima[index], share);
        counts[index] += 1;
      }
    }
  }

  const dense = totalText >= 360 || tableData.rows.length >= 9 || (columnCount >= 4 && totalText >= 180);
  const fraction = dense
    ? 1
    : Math.min(0.96, Math.max(0.34, 0.24 + totalText / 520 + columnCount * 0.055));
  const target = Math.max(columnCount * 850, Math.round(CONTENT_WIDTH * fraction));
  const weights = totals.map((total, index) => {
    const average = total / Math.max(counts[index], 1);
    return Math.max(5, maxima[index] * 0.7 + average * 0.3);
  });
  const minimum = 650;
  const distributable = Math.max(0, target - minimum * columnCount);
  const weightSum = weights.reduce((sum, value) => sum + value, 0);
  const widths = weights.map((weight) => minimum + Math.round(distributable * weight / weightSum));
  widths[widths.length - 1] += target - widths.reduce((sum, value) => sum + value, 0);
  return { columnCount, target, widths };
}

const borders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
};

function makeTable(html) {
  const tableData = parseTable(html);
  const sizing = computeWidths(tableData);
  const rows = tableData.rows.map((row, rowIndex) => {
    const classes = (row.attributes.class || "").split(/\s+/);
    const isHeader = classes.includes("hdr") || row.cells.every((cell) => cell.header);
    const isPremise = classes.includes("premise");
    const isNote = classes.includes("note");
    const fill = isHeader ? "D9D9D9" : isPremise ? "EDEDED" : isNote ? "EAEFF3" : rowIndex % 2 ? "FFFFFF" : "F7F7F7";

    const cells = row.cells.map((cell) => {
      const textLength = visualLength(cell.html);
      const isFirstColumn = cell.columnStart === 0;
      const center = isHeader || (isFirstColumn && textLength <= 80);
      const cellWidth = sizing.widths
        .slice(cell.columnStart, cell.columnStart + cell.columnSpan)
        .reduce((sum, width) => sum + width, 0);
      return new TableCell({
        rowSpan: cell.rowSpan,
        columnSpan: cell.columnSpan,
        width: { size: cellWidth, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        shading: { type: ShadingType.CLEAR, fill },
        margins: { top: 90, bottom: 90, start: 100, end: 100 },
        borders,
        children: [new Paragraph({
          alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
          spacing: { before: 0, after: 0, line: 240 },
          children: richRuns(cell.html, { bold: isHeader }),
        })],
      });
    });
    return new TableRow({ children: cells, tableHeader: isHeader, cantSplit: true });
  });

  return new Table({
    rows,
    width: { size: sizing.target, type: WidthType.DXA },
    columnWidths: sizing.widths,
    layout: TableLayoutType.FIXED,
    alignment: AlignmentType.CENTER,
    borders,
    margins: { top: 90, bottom: 90, start: 100, end: 100 },
  });
}

function splitBlocks(markdown) {
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const blocks = [];
  let paragraph = [];
  let table = [];
  let inTable = false;

  function flushParagraph() {
    const text = paragraph.join(" ").trim();
    if (text) blocks.push({ type: "paragraph", value: text });
    paragraph = [];
  }

  for (const line of lines) {
    if (!inTable && /^\s*<table\b/i.test(line)) {
      flushParagraph();
      inTable = true;
      table = [line];
      if (/<\/table>\s*$/i.test(line)) {
        blocks.push({ type: "table", value: table.join("\n") });
        inTable = false;
      }
      continue;
    }
    if (inTable) {
      table.push(line);
      if (/<\/table>\s*$/i.test(line)) {
        blocks.push({ type: "table", value: table.join("\n") });
        inTable = false;
      }
      continue;
    }
    const heading = line.match(/^(#{1,3})\s+(.+)$/);
    if (heading) {
      flushParagraph();
      blocks.push({ type: "heading", level: heading[1].length, value: heading[2].trim() });
    } else if (!line.trim()) {
      flushParagraph();
    } else if (/^```/.test(line.trim())) {
      flushParagraph();
    } else {
      paragraph.push(line.trim());
    }
  }
  if (inTable) fail("input contains an unclosed <table>");
  flushParagraph();
  return blocks;
}

function bodyChildren(markdown) {
  const children = [];
  for (const block of splitBlocks(markdown)) {
    if (block.type === "table") {
      children.push(makeTable(block.value));
      children.push(new Paragraph({ spacing: { before: 0, after: 100 } }));
      continue;
    }
    if (block.type === "heading") {
      const levels = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3];
      children.push(new Paragraph({
        heading: levels[block.level - 1],
        spacing: { before: block.level === 1 ? 220 : 160, after: 100 },
        children: richRuns(block.value, { bold: true, size: block.level === 1 ? 24 : 22 }),
      }));
      continue;
    }
    children.push(new Paragraph({
      spacing: { before: 0, after: 100, line: 280 },
      children: richRuns(block.value, { size: 20 }),
    }));
  }
  return children;
}

async function main() {
  const [inputArg, outputArg] = process.argv.slice(2);
  if (!inputArg || !outputArg || process.argv.length !== 4) {
    fail("usage: node scripts/build_docx.js <input.md> <output.docx>");
  }
  const inputPath = path.resolve(inputArg);
  const outputPath = path.resolve(outputArg);
  if (!fs.existsSync(inputPath) || !fs.statSync(inputPath).isFile()) fail(`input file not found: ${inputPath}`);
  if (path.extname(inputPath).toLowerCase() !== ".md") fail("input file must use the .md extension");
  if (path.extname(outputPath).toLowerCase() !== ".docx") fail("output file must use the .docx extension");

  const markdown = fs.readFileSync(inputPath, "utf8");
  if (!markdown.trim()) fail("input file is empty");
  const { attributes, body } = parseFrontMatter(markdown);
  const title = attributes.title_cn || path.basename(inputPath, path.extname(inputPath));
  const productType = String(attributes.product_type || attributes.type || "study").toLowerCase();
  const isQuickReference = /quick|速查/.test(productType);
  const contentChildren = bodyChildren(body);
  if (!isQuickReference) {
    contentChildren.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 220, after: 0 },
      children: [new TextRun({ text: "— 笔记整理完毕 —", font: FONT, size: 18, color: "808080" })],
    }));
  }

  const footer = new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "第 ", font: FONT, size: 17, color: "808080" }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 17, color: "808080" }),
        new TextRun({ text: " 页　共 ", font: FONT, size: 17, color: "808080" }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 17, color: "808080" }),
        new TextRun({ text: " 页", font: FONT, size: 17, color: "808080" }),
      ],
    })],
  });

  const doc = new Document({
    styles: {
      default: { document: { run: { font: FONT, size: 20, color: "000000" } } },
      paragraphStyles: [
        { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 40, bold: true, color: "000000" },
          paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 200 } } },
        { id: "Heading1", name: "heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 24, bold: true, color: "000000" },
          paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 0 } },
        { id: "Heading2", name: "heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 22, bold: true, color: "000000" },
          paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 } },
        { id: "Heading3", name: "heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 20, bold: true, color: "000000" },
          paragraph: { spacing: { before: 140, after: 60 }, outlineLevel: 2 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          // docx swaps explicit width/height when LANDSCAPE is set, so provide
          // portrait-order dimensions here to obtain a landscape A4 pgSz.
          size: { width: PAGE_HEIGHT, height: PAGE_WIDTH, orientation: PageOrientation.LANDSCAPE },
          margin: { top: PAGE_MARGIN, right: PAGE_MARGIN, bottom: PAGE_MARGIN, left: PAGE_MARGIN },
        },
      },
      footers: { default: footer },
      children: [
        new Paragraph({
          text: title,
          style: "Title",
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "595959", space: 6 } },
        }),
        new TableOfContents("目录", { hyperlink: true, headingStyleRange: "1-2" }),
        new Paragraph({ children: [new PageBreak()] }),
        ...contentChildren,
      ],
    }],
  });

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, await Packer.toBuffer(doc));
  process.stdout.write(`${outputPath}\n`);
}

main().catch((error) => fail(error && error.stack ? error.stack : String(error)));
