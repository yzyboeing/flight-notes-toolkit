#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const outputPath = path.join(root, "dist", "flight_theory_notes_prompt.md");
const referenceOrder = [
  "references/fidelity.md",
  "references/decisions.md",
  "references/structure.md",
  "references/formats.md",
  "references/output.md",
  "references/source_material.md",
  "references/library.md",
];

function fail(message) {
  process.stderr.write(`export_prompt: ${message}\n`);
  process.exit(1);
}

function readRequired(relativePath) {
  const absolutePath = path.join(root, relativePath);
  if (!fs.existsSync(absolutePath)) fail(`missing required file: ${relativePath}`);
  const value = fs.readFileSync(absolutePath, "utf8").replace(/\r\n/g, "\n").trim();
  if (!value) fail(`required file is empty: ${relativePath}`);
  return value;
}

function stripFrontMatter(markdown) {
  if (!markdown.startsWith("---\n")) return markdown;
  const end = markdown.indexOf("\n---\n", 4);
  if (end === -1) fail("SKILL.md has unclosed YAML front matter");
  return markdown.slice(end + 5).trim();
}

function stripSkillOnly(markdown) {
  const startTag = "<!-- SKILL_ONLY_START -->";
  const endTag = "<!-- SKILL_ONLY_END -->";
  const start = markdown.indexOf(startTag);
  const end = markdown.indexOf(endTag);
  if (start === -1 || end === -1 || end < start) {
    fail("SKILL.md must contain one valid SKILL_ONLY marker pair");
  }
  return `${markdown.slice(0, start).trim()}\n\n${markdown.slice(end + endTag.length).trim()}`;
}

function adaptSkillForPrompt(markdown) {
  const start = markdown.indexOf("## 文件路由");
  const end = markdown.indexOf("## 两种工作模式");
  if (start === -1 || end === -1 || end < start) {
    fail("SKILL.md must contain 文件路由 before 两种工作模式");
  }
  const replacement = [
    "## 使用说明",
    "",
    "以下各节已经由导出脚本从唯一规则源完整展开。内容忠实原则优先级最高；长期决策每次都要遵守，已经确认的问题不得重复询问。",
    "",
  ].join("\n");
  return `${markdown.slice(0, start).trim()}\n\n${replacement}\n${markdown.slice(end).trim()}`;
}

function adaptReferencesForPrompt(markdown) {
  return markdown
    .replace(/`references\/structure\.md`/g, "「结构方法」一节")
    .replace(/`fidelity\.md`/g, "「内容忠实原则、主动判断与反馈」一节")
    .replace(/`structure\.md`/g, "「结构方法」一节")
    .replace(/当前 `documents` skill/g, "当前平台的文档生成与渲染工具")
    .replace(/按 `documents` skill/g, "按当前平台的文档生成与渲染流程");
}

function localDate() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

const version = readRequired("VERSION");
if (!/^\d+\.\d+(?:\.\d+)?$/.test(version)) {
  fail(`invalid VERSION value: ${JSON.stringify(version)}`);
}

const skill = adaptSkillForPrompt(stripSkillOnly(stripFrontMatter(readRequired("SKILL.md"))));
const references = referenceOrder.map(readRequired);
const iterationVariant = readRequired("tools/iteration_prompt_variant.md");

const banner = [
  `> **版本 ${version}　生成于 ${localDate()}**`,
  "> 本文由 `flight-theory-notes` skill 自动导出，**请勿手工修改**——改动会在下次导出时被覆盖。",
  "> 篇幅较长时请使用 Project 指令 / Gemini Gem 指令，或在会话开头整段粘贴。",
].join("\n");

const generated = adaptReferencesForPrompt([banner, skill, ...references, iterationVariant]
  .join("\n\n---\n\n")
  .replace(/\n{4,}/g, "\n\n\n")
  .trimEnd()) + "\n";

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, generated, "utf8");
process.stdout.write(`${outputPath}\n`);
