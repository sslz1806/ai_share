const pptxgen = require("pptxgenjs");
const fs = require("fs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "AI分享";
pres.title = "AI大矩阵分享";

// ========== COLOR PALETTE ==========
const C = {
  darkBg: "1B2A4A",
  primary: "2B5A8C",
  accent: "00A3C4",
  accent2: "0BC5EA",
  lightBg: "F5F7FA",
  white: "FFFFFF",
  textDark: "1A202C",
  textGray: "4A5568",
  textLight: "A0AEC0",
  cardBg: "FFFFFF",
  cardBorder: "E2E8F0",
  highlight: "EBF8FF",
  green: "38A169",
  orange: "DD6B20",
  red: "E53E3E",
};

// ========== HELPERS ==========
const cardShadow = () => ({
  type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.10,
});
const sectionShadow = () => ({
  type: "outer", color: "000000", blur: 12, offset: 3, angle: 135, opacity: 0.15,
});

function gifData(name) {
  const p = `C:/Users/20561/Desktop/extracted_${name}`;
  const b64 = fs.readFileSync(p).toString("base64");
  return `image/gif;base64,${b64}`;
}

const slideTitleOpts = (y = 0.3) => ({
  x: 0.6, y, w: 8.8, h: 0.7, fontSize: 26, fontFace: "Arial",
  color: C.textDark, bold: true, margin: 0,
});

function addSectionDivider(presObj, sectionNum, title, subtitle) {
  const slide = presObj.addSlide();
  slide.background = { color: C.darkBg };
  slide.addShape(presObj.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.04, fill: { color: C.accent },
  });
  slide.addText(`0${sectionNum}`, {
    x: 0.8, y: 1.2, w: 2, h: 1.2, fontSize: 72, fontFace: "Arial",
    color: C.accent, bold: true, margin: 0,
  });
  slide.addText(title, {
    x: 0.8, y: 2.5, w: 8.4, h: 1.0, fontSize: 36, fontFace: "Arial",
    color: C.white, bold: true, margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.8, y: 3.5, w: 8.4, h: 0.6, fontSize: 16, fontFace: "Calibri",
      color: C.textLight, margin: 0,
    });
  }
  slide.addShape(presObj.shapes.RECTANGLE, {
    x: 0.8, y: 4.8, w: 3, h: 0.04, fill: { color: C.accent },
  });
  return slide;
}

function addPageNumber(slide, num, total) {
  slide.addText(`${num} / ${total}`, {
    x: 8.2, y: 5.2, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Calibri",
    color: C.textLight, align: "right", margin: 0,
  });
}

// ========== CONSTANTS ==========
const totalSlides = 30;

// =====================================================================
// SLIDE 1: COVER
// =====================================================================
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent },
  });
  slide.addShape(pres.shapes.OVAL, {
    x: 7.5, y: -1.5, w: 4, h: 4, fill: { color: C.accent, transparency: 90 },
  });
  slide.addShape(pres.shapes.OVAL, {
    x: -1.5, y: 3.5, w: 3.5, h: 3.5, fill: { color: C.accent2, transparency: 85 },
  });
  slide.addText("AI大矩阵分享", {
    x: 0.8, y: 1.6, w: 8.4, h: 1.2, fontSize: 48, fontFace: "Arial",
    color: C.white, bold: true, margin: 0,
  });
  slide.addText("从大模型到智能体 · AI 全面赋能工作与生活", {
    x: 0.8, y: 2.9, w: 8.4, h: 0.5, fontSize: 18, fontFace: "Calibri",
    color: C.accent, margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 3.7, w: 2.5, h: 0.04, fill: { color: C.accent },
  });
  slide.addText("面向全体同事 · AI 知识分享", {
    x: 0.8, y: 4.5, w: 5, h: 0.4, fontSize: 13, fontFace: "Calibri",
    color: C.textLight, margin: 0,
  });
})();

// =====================================================================
// SLIDE 2: TABLE OF CONTENTS (5 chapters)
// =====================================================================
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("目  录", { ...slideTitleOpts(0.4), fontSize: 30 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.1, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  const chapters = [
    { num: "01", title: "AI 进化史速览", desc: "大模型发展历程 & 对工作的影响" },
    { num: "02", title: "实用干货分享", desc: "模型对比 · 工具推荐 · VibeCoding" },
    { num: "03", title: "实用命令 & 内置 Skills", desc: "Claude Code 命令速查 · 办公/开发 Skills" },
    { num: "04", title: "Web 自动化", desc: "BrowserUse · Playwright CLI" },
    { num: "05", title: "Agent 约束 (Harness)", desc: "记忆文件 · SubAgent · Superpowers 工作流" },
  ];

  chapters.forEach((ch, i) => {
    const baseY = 1.4 + i * 0.85;
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7, y: baseY, w: 0.55, h: 0.55, fill: { color: C.primary },
    });
    slide.addText(ch.num, {
      x: 0.7, y: baseY, w: 0.55, h: 0.55, fontSize: 12, fontFace: "Arial",
      color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(ch.title, {
      x: 1.5, y: baseY - 0.03, w: 6, h: 0.3, fontSize: 16, fontFace: "Arial",
      color: C.textDark, bold: true, margin: 0,
    });
    slide.addText(ch.desc, {
      x: 1.5, y: baseY + 0.28, w: 6, h: 0.25, fontSize: 11, fontFace: "Calibri",
      color: C.textGray, margin: 0,
    });
  });
  addPageNumber(slide, 2, totalSlides);
})();

// =====================================================================
// SECTION 01: AI 进化史速览 (Slides 3-5)
// =====================================================================

/* ---- SLIDE 3: 大模型近年发展 ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("大模型近年发展", { ...slideTitleOpts(0.4), fontSize: 28 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.1, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  const points = [
    "近几年 AI 发展迭代迅猛，从聊天大模型进阶到 AI Agent，行业持续迎来新突破。",
    "AI 已深度融入生产与日常生活，影响力还会持续扩大。",
    "当下职场，会用 AI 的人正在逐步领跑；主动学习 AI 能力、与自身技能相结合，已是必备素养。",
  ];
  points.forEach((t, i) => {
    const y = 1.5 + i * 1.1;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y, w: 0.06, h: 0.7, fill: { color: C.accent },
    });
    slide.addText(t, { x: 0.9, y, w: 5.2, h: 0.7, fontSize: 13, fontFace: "Calibri",
      color: C.textDark, valign: "top", margin: 0 });
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 6.5, y: 1.5, w: 3.0, h: 2.6, fill: { color: C.highlight }, shadow: cardShadow(),
  });
  slide.addText("核心观点", {
    x: 6.8, y: 1.7, w: 2.5, h: 0.4, fontSize: 14, fontFace: "Arial",
    color: C.primary, bold: true, margin: 0,
  });
  slide.addText("不是 AI 淘汰人，\n是会用 AI 的人，\n正在淘汰不会用 AI 的人。", {
    x: 6.8, y: 2.2, w: 2.5, h: 1.6, fontSize: 15, fontFace: "Calibri",
    color: C.textDark, valign: "top", margin: 0,
  });
  addPageNumber(slide, 3, totalSlides);
})();

/* ---- SLIDE 4: AI进化史速览 (GIF) ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("一. AI 进化史速览", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });
  try {
    slide.addImage({
      data: gifData("image1.gif"),
      x: 0.5, y: 1.3, w: 9.0, h: 3.8,
      sizing: { type: "contain", w: 9.0, h: 3.8 },
    });
  } catch (_) {
    slide.addText("[AI 进化史时间线示意图 - 源自笔记]", {
      x: 1, y: 2, w: 8, h: 2, fontSize: 14, fontFace: "Calibri",
      color: C.textGray, align: "center", valign: "middle", margin: 0,
    });
  }
  addPageNumber(slide, 4, totalSlides);
})();

/* ---- SLIDE 5: AI对开发者的影响 ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addText("AI 对开发者的影响", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.4, w: 4.2, h: 2.5, fill: { color: C.cardBg }, shadow: cardShadow(),
  });
  slide.addText("有些人担忧…", {
    x: 0.9, y: 1.5, w: 3.7, h: 0.4, fontSize: 15, fontFace: "Arial",
    color: C.red, bold: true, margin: 0,
  });
  slide.addText("觉得 AI 是对岗位的冲击\n觉得 AI 要抢自己饭碗了", {
    x: 0.9, y: 2.0, w: 3.7, h: 0.8, fontSize: 13, fontFace: "Calibri",
    color: C.textGray, margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.4, w: 4.2, h: 2.5, fill: { color: C.cardBg }, shadow: cardShadow(),
  });
  slide.addText("更多人觉得…", {
    x: 5.5, y: 1.5, w: 3.7, h: 0.4, fontSize: 15, fontFace: "Arial",
    color: C.green, bold: true, margin: 0,
  });
  slide.addText("AI 是真的方便\n效率拉满了很爽\n学会驾驭 AI 才是关键", {
    x: 5.5, y: 2.0, w: 3.7, h: 1.0, fontSize: 13, fontFace: "Calibri",
    color: C.textGray, margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.2, w: 8.8, h: 0.9, fill: { color: C.primary }, shadow: sectionShadow(),
  });
  slide.addText("不是 AI 淘汰人，是会用 AI 的人，正在淘汰不会用 AI 的人。", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.9, fontSize: 18, fontFace: "Arial",
    color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  addPageNumber(slide, 5, totalSlides);
})();

// =====================================================================
// SECTION 02: 实用干货分享 (Slides 6-13)
// =====================================================================
addSectionDivider(pres, 2, "实用干货分享", "模型选择 · 工具对比 · 实用技巧");

/* ---- SLIDE 7: 模型选择与对比 ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("模型选择与对比", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addText("🇨🇳 国内模型", {
    x: 0.6, y: 1.3, w: 4, h: 0.4, fontSize: 16, fontFace: "Arial",
    color: C.primary, bold: true, margin: 0,
  });
  const domestic = [
    { name: "DeepSeek V4 Pro / Flash", desc: "数学 & 推理能力很强，Agent 能力大幅提升。V4 Pro + Superpower ≈ 0.8× Opus 4.7，价格仅为 1/7" },
    { name: "MiniMax 2.7", desc: "最便宜的模型，适合处理自动化任务等固定流程，省钱首选" },
    { name: "GLM 5", desc: "Agent 和代码能力强，但目前请求较慢、不太稳定" },
  ];
  domestic.forEach((m, i) => {
    const y = 1.85 + i * 0.8;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 4.2, h: 0.65, fill: { color: C.lightBg } });
    slide.addText(m.name, {
      x: 0.75, y: y + 0.02, w: 3.9, h: 0.25, fontSize: 12, fontFace: "Arial",
      color: C.textDark, bold: true, margin: 0,
    });
    slide.addText(m.desc, {
      x: 0.75, y: y + 0.28, w: 3.9, h: 0.35, fontSize: 10, fontFace: "Calibri",
      color: C.textGray, margin: 0,
    });
  });

  slide.addText("🌍 国外模型", {
    x: 5.2, y: 1.3, w: 4, h: 0.4, fontSize: 16, fontFace: "Arial",
    color: C.primary, bold: true, margin: 0,
  });
  const intl = [
    { name: "Claude Opus 4.7", desc: "工程 & Agent 能力最强。缺点：价格较高，有封号风险" },
    { name: "GPT 5.5", desc: "自主规划与自我纠错强，思维发散。缺点：上下文较短（无 1M）" },
  ];
  intl.forEach((m, i) => {
    const y = 1.85 + i * 0.8;
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.2, y, w: 4.2, h: 0.65, fill: { color: C.lightBg } });
    slide.addText(m.name, {
      x: 5.35, y: y + 0.02, w: 3.9, h: 0.25, fontSize: 12, fontFace: "Arial",
      color: C.textDark, bold: true, margin: 0,
    });
    slide.addText(m.desc, {
      x: 5.35, y: y + 0.28, w: 3.9, h: 0.35, fontSize: 10, fontFace: "Calibri",
      color: C.textGray, margin: 0,
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 3.7, w: 8.8, h: 0.65, fill: { color: C.highlight },
  });
  slide.addText("推荐方案：不对比，都要 → QShot", {
    x: 0.8, y: 3.72, w: 8.4, h: 0.3, fontSize: 14, fontFace: "Arial",
    color: C.primary, bold: true, margin: 0,
  });
  slide.addText("集成多个 AI 对话框，同一个问题同步问给多个 AI，直接对比输出内容", {
    x: 0.8, y: 4.02, w: 8.4, h: 0.3, fontSize: 11, fontFace: "Calibri",
    color: C.textGray, margin: 0,
  });
  addPageNumber(slide, 7, totalSlides);
})();

/* ---- SLIDE 8: QShot ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("QShot — 多模型统一调度", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });
  slide.addText("作用：", {
    x: 0.6, y: 1.3, w: 2, h: 0.3, fontSize: 14, fontFace: "Arial",
    color: C.textDark, bold: true, margin: 0,
  });
  const uses = ["① 对比不同 AI 的生成质量", "② 批量搜索，多模型并行工作"];
  uses.forEach((t, i) => {
    slide.addText(t, {
      x: 1.8, y: 1.65 + i * 0.35, w: 3.5, h: 0.3, fontSize: 13, fontFace: "Calibri",
      color: C.textDark, margin: 0,
    });
  });
  try {
    slide.addImage({
      data: gifData("image2.gif"),
      x: 5.0, y: 1.2, w: 4.5, h: 3.5,
      sizing: { type: "contain", w: 4.5, h: 3.5 },
    });
  } catch (_) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.2, y: 1.5, w: 4.0, h: 2.8, fill: { color: C.lightBg }, shadow: cardShadow(),
    });
    slide.addText("[QShot 截图占位\n现场演示]", {
      x: 5.2, y: 1.5, w: 4.0, h: 2.8, fontSize: 14, fontFace: "Calibri",
      color: C.textLight, align: "center", valign: "middle", margin: 0,
    });
  }
  slide.addText("💡 QShot 的模式与 SubAgent 的模式非常类似", {
    x: 0.6, y: 4.6, w: 8, h: 0.4, fontSize: 13, fontFace: "Calibri",
    color: C.textGray, italic: true, margin: 0,
  });
  addPageNumber(slide, 8, totalSlides);
})();

/* ---- SLIDE 9: VibeCoding ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("VibeCoding", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });
  slide.addText('VibeCoding 是一种 AI 辅助编程的新范式，强调“随状态走、随感觉流”，\n用自然语言指导 AI 生成代码，极大降低编程门槛。', {
    x: 0.6, y: 1.3, w: 8.8, h: 0.8, fontSize: 14, fontFace: "Calibri",
    color: C.textDark, margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 2.4, w: 4.2, h: 2.5, fill: { color: C.highlight }, shadow: cardShadow(),
  });
  slide.addText("✅ 优点", {
    x: 0.8, y: 2.5, w: 3.8, h: 0.35, fontSize: 15, fontFace: "Arial",
    color: C.green, bold: true, margin: 0,
  });
  slide.addText([
    { text: "开发效率极高，快速迭代", options: { bullet: true, breakLine: true } },
    { text: "降低编程入门门槛", options: { bullet: true, breakLine: true } },
    { text: "适合原型验证和快速实现", options: { bullet: true } },
  ], { x: 0.8, y: 2.95, w: 3.8, h: 1.6, fontSize: 12, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 2.4, w: 4.2, h: 2.5, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addText("⚠️ 注意事项", {
    x: 5.4, y: 2.5, w: 3.8, h: 0.35, fontSize: 15, fontFace: "Arial",
    color: C.orange, bold: true, margin: 0,
  });
  slide.addText([
    { text: '需要明确约束，否则 AI 容易“放飞”', options: { bullet: true, breakLine: true } },
    { text: "生成的代码需人工 review", options: { bullet: true, breakLine: true } },
    { text: "配合 Harness/规范使用更佳", options: { bullet: true } },
  ], { x: 5.4, y: 2.95, w: 3.8, h: 1.6, fontSize: 12, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });
  addPageNumber(slide, 9, totalSlides);
})();

/* ---- SLIDES 10-13: 工具对比 (4 tools) ---- */
const tools = [
  { name: "Codex", tag: "综合性价比最高", tagColor: C.green,
    pros: ["量大管饱，生态齐全", "权限高，设置简单", "对新手友好，Cowork 场景最好用"],
    cons: ["绑定 OpenAI 模型", "响应速度较慢"] },
  { name: "Cursor", tag: "专门的 Coding IDE", tagColor: C.primary,
    pros: ["功能齐全（代码补齐/插件/外接模型）", "对国区友好，支付宝可支付", "Coding 体验最好"],
    cons: ["高级模型额度较少"] },
  { name: "Claude Code", tag: "灵活性强，Agent 深度", tagColor: C.accent,
    pros: ["支持外接多种模型", "所有操作显示在对话框", "自家模型强，适合学习 Agent"],
    cons: ["配置较复杂，需命令行", "模型对国区不友好"] },
  { name: "Copilot", tag: "微软生态集成", tagColor: C.orange,
    pros: ["VS Code / GitHub 深度集成", "GitHub Copilot 生态完善"],
    cons: ["需要付费订阅", "灵活性相对较低"] },
];

tools.forEach((tool, idx) => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("AI 编程工具对比", { ...slideTitleOpts(0.25), fontSize: 24 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 0.95, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.0, h: 3.8, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.0, h: 0.7, fill: { color: C.primary },
  });
  slide.addText(tool.name, {
    x: 0.8, y: 1.35, w: 3.6, h: 0.6, fontSize: 26, fontFace: "Arial",
    color: C.white, bold: true, valign: "middle", margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 2.15, w: 2.2, h: 0.3, fill: { color: tool.tagColor },
  });
  slide.addText(tool.tag, {
    x: 0.8, y: 2.15, w: 2.2, h: 0.3, fontSize: 11, fontFace: "Arial",
    color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  slide.addText("✅ 优点", {
    x: 0.8, y: 2.6, w: 3.6, h: 0.3, fontSize: 13, fontFace: "Arial",
    color: C.green, bold: true, margin: 0,
  });
  slide.addText(tool.pros.map((p, pi) => ({
    text: p, options: { bullet: true, breakLine: pi < tool.pros.length - 1 },
  })), { x: 0.8, y: 2.9, w: 3.6, h: 0.9, fontSize: 11, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addText("❌ 缺点", {
    x: 0.8, y: 3.7, w: 3.6, h: 0.3, fontSize: 13, fontFace: "Arial",
    color: C.red, bold: true, margin: 0,
  });
  slide.addText(tool.cons.map((c, ci) => ({
    text: c, options: { bullet: true, breakLine: ci < tool.cons.length - 1 },
  })), { x: 0.8, y: 4.0, w: 3.6, h: 0.8, fontSize: 11, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.0, y: 1.3, w: 4.5, h: 3.8, fill: { color: C.cardBg }, shadow: cardShadow(),
  });
  slide.addText("适合场景", {
    x: 5.2, y: 1.4, w: 4.1, h: 0.35, fontSize: 14, fontFace: "Arial",
    color: C.primary, bold: true, margin: 0,
  });
  slide.addText([
    { text: `${tool.name} 最适合的场景取决于团队的技术栈和工作流偏好。`, options: { breakLine: true } },
    { text: tool.name === "Codex" ? "适合团队协作、快速原型的场景" :
      tool.name === "Cursor" ? "适合深度编码、需要 IDE 全功能支持的用户" :
      tool.name === "Claude Code" ? "适合想要深入 Agent 开发、学习 AI 编程的用户" :
      "适合微软生态内的团队使用", options: { breakLine: true } },
  ], { x: 5.2, y: 1.9, w: 4.1, h: 0.8, fontSize: 11, fontFace: "Calibri", color: C.textDark, margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 3.8, w: 4.1, h: 0.04, fill: { color: C.cardBorder },
  });
  slide.addText("💡 使用建议", {
    x: 5.2, y: 4.0, w: 4.1, h: 0.3, fontSize: 13, fontFace: "Arial",
    color: C.accent, bold: true, margin: 0,
  });
  slide.addText(
    tool.name === "Codex" ? "推荐作为入门首选，生态完整、上手快" :
    tool.name === "Cursor" ? "VSCode 用户强烈推荐，体验无缝衔接" :
    tool.name === "Claude Code" ? "搭配 Claude.md + Agent 规范使用，效果最佳" :
    "与 GitHub 生态深度绑定，适合 Azure/Office 开发者",
    { x: 5.2, y: 4.3, w: 4.1, h: 0.5, fontSize: 11, fontFace: "Calibri", color: C.textGray, margin: 0 });

  if (tool.name === "Claude Code") {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.2, y: 4.7, w: 4.1, h: 0.04, fill: { color: C.cardBorder },
    });
    slide.addText("另：Claude Desktop 提供桌面端 AI 助手体验", {
      x: 5.2, y: 4.8, w: 4.1, h: 0.3, fontSize: 10, fontFace: "Calibri",
      color: C.textLight, italic: true, margin: 0,
    });
  }
  addPageNumber(slide, 10 + idx, totalSlides);
});

// =====================================================================
// SECTION 03: 实用命令 & 内置 Skills (NEW - Slides 14-17)
// =====================================================================
addSectionDivider(pres, 3, "实用命令 & 内置 Skills", "Claude Code 命令速查 · 办公/开发类 Skills");

/* ---- SLIDE 15: 实用命令速查 ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("Claude Code 实用命令速查", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  const groups = [
    {
      title: "对话控制",
      cmds: [
        { cmd: "/debug", desc: "进入调试模式，查看 AI 内部推理" },
        { cmd: "/compact", desc: "压缩上下文，节省 Token" },
        { cmd: "/clear", desc: "清空当前对话历史" },
      ],
    },
    {
      title: "输出控制",
      cmds: [
        { cmd: "双 Esc", desc: "立即中断 AI 输出" },
        { cmd: "/model", desc: "切换当前模型" },
        { cmd: "cc - switch", desc: "系统配置切换" },
      ],
    },
    {
      title: "实用技巧",
      cmds: [
        { cmd: "/init", desc: "初始化项目 CLAUDE.md" },
        { cmd: "/review", desc: "对当前分支进行代码审查" },
        { cmd: "/compact + /clear", desc: "组合使用清理上下文" },
      ],
    },
  ];

  groups.forEach((g, gi) => {
    const x = 0.4 + gi * 3.2;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.3, w: 3.0, h: 0.4, fill: { color: C.primary },
    });
    slide.addText(g.title, {
      x, y: 1.3, w: 3.0, h: 0.4, fontSize: 13, fontFace: "Arial",
      color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });

    g.cmds.forEach((c, i) => {
      const y = 1.8 + i * 0.95;
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w: 3.0, h: 0.8, fill: { color: C.lightBg },
      });
      slide.addText(c.cmd, {
        x: x + 0.15, y: y + 0.05, w: 2.7, h: 0.3, fontSize: 12, fontFace: "Consolas",
        color: C.accent, bold: true, margin: 0,
      });
      slide.addText(c.desc, {
        x: x + 0.15, y: y + 0.35, w: 2.7, h: 0.35, fontSize: 10, fontFace: "Calibri",
        color: C.textGray, margin: 0,
      });
    });
  });

  addPageNumber(slide, 15, totalSlides);
})();

/* ---- SLIDE 16: Claude Code 内置 Skills ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("Claude Code 内置 Skills 概览", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addText("Claude Code 通过 Skills 机制扩展 AI 能力，覆盖办公、开发、设计等场景。", {
    x: 0.6, y: 1.2, w: 8.8, h: 0.4, fontSize: 13, fontFace: "Calibri",
    color: C.textGray, margin: 0,
  });

  // 2x3 grid of skill cards
  const skills = [
    { icon: "📊", name: "xlsx", desc: "Excel / CSV 文件处理", category: "办公" },
    { icon: "📽", name: "pptx", desc: "PowerPoint 演示文稿制作", category: "办公" },
    { icon: "📝", name: "docx", desc: "Word 文档处理与编辑", category: "办公" },
    { icon: "📄", name: "pdf", desc: "PDF 文件读取与处理", category: "办公" },
    { icon: "🎨", name: "frontend-design", desc: "前端界面设计与交互", category: "开发" },
    { icon: "🔍", name: "code-review", desc: "代码审查与质量检查", category: "开发" },
  ];

  skills.forEach((s, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.6 + col * 3.1;
    const y = 1.7 + row * 1.6;

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.8, h: 1.35, fill: { color: C.lightBg }, shadow: cardShadow(),
    });
    // accent top bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.8, h: 0.06, fill: { color: row === 0 ? C.primary : C.accent },
    });
    // category tag
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.15, y: y + 0.15, w: 0.6, h: 0.22, fill: { color: row === 0 ? C.primary : C.accent },
    });
    slide.addText(s.category, {
      x: x + 0.15, y: y + 0.15, w: 0.6, h: 0.22, fontSize: 8, fontFace: "Arial",
      color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(s.icon, {
      x: x + 0.15, y: y + 0.5, w: 0.4, h: 0.4, fontSize: 18, fontFace: "Calibri",
      color: C.textDark, margin: 0,
    });
    slide.addText(s.name, {
      x: x + 0.6, y: y + 0.5, w: 2.0, h: 0.3, fontSize: 13, fontFace: "Arial",
      color: C.textDark, bold: true, margin: 0,
    });
    slide.addText(s.desc, {
      x: x + 0.15, y: y + 0.9, w: 2.5, h: 0.35, fontSize: 10, fontFace: "Calibri",
      color: C.textGray, margin: 0,
    });
  });

  addPageNumber(slide, 16, totalSlides);
})();

/* ---- SLIDE 17: 更多实用 Skills 推荐 ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("更多实用 Skills 推荐", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  const recs = [
    {
      cat: "🛠 开发效率",
      items: [
        { name: "playwright-cli", desc: "浏览器自动化测试，Web 交互验证" },
        { name: "security-review", desc: "安全审查，发现潜在漏洞" },
        { name: "brand-guidelines", desc: "品牌风格指南，保持设计一致" },
      ],
    },
    {
      cat: "🧠 AI 进阶",
      items: [
        { name: "claude-api", desc: "Claude API 集成开发，构建 AI 应用" },
        { name: "webapp-testing", desc: "Web 应用端到端测试" },
        { name: "mcp-builder", desc: "构建 MCP 服务器，扩展 AI 能力" },
      ],
    },
    {
      cat: "🎯 工作流",
      items: [
        { name: "test-driven-development", desc: "TDD 测试驱动开发流程" },
        { name: "subagent-driven-dev", desc: "子代理驱动开发，任务分治" },
        { name: "systematic-debugging", desc: "系统化调试流程" },
      ],
    },
  ];

  recs.forEach((group, gi) => {
    const x = 0.4 + gi * 3.2;
    slide.addText(group.cat, {
      x, y: 1.2, w: 3.0, h: 0.35, fontSize: 14, fontFace: "Arial",
      color: C.primary, bold: true, margin: 0,
    });

    group.items.forEach((item, ii) => {
      const y = 1.65 + ii * 1.1;
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w: 3.0, h: 0.95, fill: { color: C.lightBg }, shadow: cardShadow(),
      });
      slide.addText(item.name, {
        x: x + 0.12, y: y + 0.08, w: 2.76, h: 0.28, fontSize: 12, fontFace: "Consolas",
        color: C.accent, bold: true, margin: 0,
      });
      slide.addText(item.desc, {
        x: x + 0.12, y: y + 0.4, w: 2.76, h: 0.45, fontSize: 10, fontFace: "Calibri",
        color: C.textGray, margin: 0,
      });
    });
  });

  addPageNumber(slide, 17, totalSlides);
})();

// =====================================================================
// SECTION 04: Web 自动化 (Slides 18-19)
// =====================================================================
addSectionDivider(pres, 4, "Web 自动化", "BrowserUse · Playwright CLI");

/* ---- SLIDE 19: Web浏览器自动化 ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("Web 浏览器自动化工具", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.4, w: 4.2, h: 3.0, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.4, w: 4.2, h: 0.55, fill: { color: C.primary },
  });
  slide.addText("BrowserUse", {
    x: 0.8, y: 1.42, w: 3.8, h: 0.5, fontSize: 18, fontFace: "Arial",
    color: C.white, bold: true, valign: "middle", margin: 0,
  });
  slide.addText([
    { text: "适合复杂、需要灵活探索的场景", options: { bullet: true, breakLine: true } },
    { text: "每一步都会调用大模型，Token 消耗大", options: { bullet: true, breakLine: true } },
    { text: "建议探索完成后固化为 Skill 使用", options: { bullet: true } },
  ], { x: 0.8, y: 2.1, w: 3.8, h: 2.0, fontSize: 12, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.4, w: 4.2, h: 3.0, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.4, w: 4.2, h: 0.55, fill: { color: C.accent },
  });
  slide.addText("Playwright CLI", {
    x: 5.4, y: 1.42, w: 3.8, h: 0.5, fontSize: 18, fontFace: "Arial",
    color: C.white, bold: true, valign: "middle", margin: 0,
  });
  slide.addText([
    { text: "确定性的浏览器操作自动化", options: { bullet: true, breakLine: true } },
    { text: "适合固定流程的测试和操作", options: { bullet: true, breakLine: true } },
    { text: "效率高，Token 消耗可控", options: { bullet: true } },
  ], { x: 5.4, y: 2.1, w: 3.8, h: 2.0, fontSize: 12, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.6, w: 8.8, h: 0.5, fill: { color: C.highlight },
  });
  slide.addText("💡 建议：BrowserUse 探索 → 蒸馏成 Skill（固化流程）→ Playwright CLI 自动化执行", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.5, fontSize: 13, fontFace: "Calibri",
    color: C.primary, valign: "middle", margin: 0,
  });
  addPageNumber(slide, 19, totalSlides);
})();

// =====================================================================
// SECTION 05: Agent 约束 (Harness) (Slides 20-28)
// =====================================================================
addSectionDivider(pres, 5, "Agent 约束 (Harness)", "记忆文件体系 · SubAgent · Superpowers 工作流");

/* ---- SLIDE 21: SubAgent & Agent Teams ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("SubAgent & Agent Teams", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.2, h: 1.6, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addText("问题：AI 编程的随机性", {
    x: 0.8, y: 1.35, w: 3.8, h: 0.35, fontSize: 14, fontFace: "Arial",
    color: C.red, bold: true, margin: 0,
  });
  slide.addText([
    { text: "没有纪律的编程 = 给自己挖坑", options: { bullet: true, breakLine: true } },
    { text: "长对话导致上下文污染", options: { bullet: true, breakLine: true } },
    { text: 'AI 容易“遗忘”之前的约束', options: { bullet: true } },
  ], { x: 0.8, y: 1.7, w: 3.8, h: 1.0, fontSize: 11, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.3, w: 4.2, h: 1.6, fill: { color: C.highlight }, shadow: cardShadow(),
  });
  slide.addText("方案：SubAgent", {
    x: 5.4, y: 1.35, w: 3.8, h: 0.35, fontSize: 14, fontFace: "Arial",
    color: C.green, bold: true, margin: 0,
  });
  slide.addText([
    { text: "为每个小任务派遣全新 AI 子代理", options: { bullet: true, breakLine: true } },
    { text: '每个子任务“新手上路”，无历史包袱', options: { bullet: true, breakLine: true } },
    { text: "主Agent 统筹，子Agent 专注执行", options: { bullet: true } },
  ], { x: 5.4, y: 1.7, w: 3.8, h: 1.0, fontSize: 11, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  try {
    slide.addImage({
      data: gifData("image3.gif"),
      x: 1.5, y: 3.2, w: 7.0, h: 2.0,
      sizing: { type: "contain", w: 7.0, h: 2.0 },
    });
  } catch (_) {
    slide.addShape(pres.shapes.RECTANGLE, { x: 2, y: 3.3, w: 6, h: 1.8, fill: { color: C.lightBg } });
    slide.addText("[SubAgent 架构示意图]", {
      x: 2, y: 3.3, w: 6, h: 1.8, fontSize: 13, fontFace: "Calibri",
      color: C.textLight, align: "center", valign: "middle", margin: 0,
    });
  }
  addPageNumber(slide, 21, totalSlides);
})();

/* ---- SLIDE 22: 提示词与记忆文件体系 ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("提示词与记忆文件体系", { ...slideTitleOpts(0.25), fontSize: 24 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 0.95, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  const tiers = [
    { name: "全局提示词", file: "~/.claude/CLAUDE.md", desc: "Agent 运行规范和偏好", color: C.primary },
    { name: "项目提示词", file: "项目目录 / CLAUDE.md", desc: "项目简介、规范、约束", color: C.accent },
    { name: "文件夹提示词", file: "各目录 / AGENT.md", desc: "文件夹简介、约束、规范", color: C.accent2 },
  ];
  tiers.forEach((t, i) => {
    const y = 1.2 + i * 1.2;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 8.8, h: 1.0, fill: { color: C.lightBg } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.08, h: 1.0, fill: { color: t.color } });
    slide.addText(t.name, {
      x: 0.9, y: y + 0.05, w: 2.5, h: 0.4, fontSize: 15, fontFace: "Arial",
      color: C.textDark, bold: true, margin: 0,
    });
    slide.addText(t.file, {
      x: 0.9, y: y + 0.45, w: 3.5, h: 0.25, fontSize: 10, fontFace: "Consolas",
      color: C.textGray, margin: 0,
    });
    slide.addText(t.desc, {
      x: 4.5, y: y + 0.15, w: 4.5, h: 0.7, fontSize: 13, fontFace: "Calibri",
      color: C.textDark, valign: "middle", margin: 0,
    });
  });
  slide.addText("💡 项目提示词可以索引到文件夹提示词，压缩 CLAUDE.md 的内容", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4, fontSize: 12, fontFace: "Calibri",
    color: C.textGray, italic: true, margin: 0,
  });
  addPageNumber(slide, 22, totalSlides);
})();

/* ---- SLIDE 23: 实践结论 ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("实践结论", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.2, h: 2.5, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addText("📝 我们的发现", {
    x: 0.8, y: 1.4, w: 3.8, h: 0.35, fontSize: 14, fontFace: "Arial",
    color: C.primary, bold: true, margin: 0,
  });
  slide.addText([
    { text: "项目简介和规范有作用，但约束性不强", options: { bullet: true, breakLine: true } },
    { text: "对话框提示词优先级更高", options: { bullet: true, breakLine: true } },
    { text: "复杂约束可能被遗忘或不生效", options: { bullet: true, breakLine: true } },
    { text: "太宽泛的约束不一定能执行", options: { bullet: true } },
  ], { x: 0.8, y: 1.85, w: 3.8, h: 1.7, fontSize: 12, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.3, w: 4.2, h: 2.5, fill: { color: C.highlight }, shadow: cardShadow(),
  });
  slide.addText("💡 推荐方案 → Skills", {
    x: 5.4, y: 1.4, w: 3.8, h: 0.35, fontSize: 14, fontFace: "Arial",
    color: C.accent, bold: true, margin: 0,
  });
  slide.addText([
    { text: "Andrej-Karpathy-Skills", options: { bullet: true, breakLine: true } },
    { text: "规范代码编写流程（类似 Plan Mode）", options: { bullet: true, breakLine: true, indentLevel: 1 } },
    { text: "", options: { breakLine: true } },
    { text: "从一个插件入手，约束并驾驭 AI 编程", options: { bullet: true, breakLine: true } },
    { text: "→ Superpowers 强制工作流", options: { bullet: true, breakLine: true, indentLevel: 1 } },
    { text: "→ TDD 测试驱动开发", options: { bullet: true, breakLine: true, indentLevel: 1 } },
  ], { x: 5.4, y: 1.85, w: 3.8, h: 1.8, fontSize: 11, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.2, w: 8.8, h: 0.5, fill: { color: C.primary },
  });
  slide.addText("🔗 github.com/obra/superpowers", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.5, fontSize: 13, fontFace: "Consolas",
    color: C.white, align: "center", valign: "middle", margin: 0,
  });
  addPageNumber(slide, 23, totalSlides);
})();

/* ---- SLIDE 24: Superpowers 七大步骤 ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("Superpowers 七大步骤", { ...slideTitleOpts(0.3), fontSize: 26 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent },
  });
  slide.addText('核心目标：通过强制工作流，让 AI 编程从“随缘”变为“可靠”', {
    x: 0.6, y: 1.2, w: 8.8, h: 0.4, fontSize: 14, fontFace: "Calibri",
    color: C.textGray, margin: 0,
  });

  const steps = [
    { num: "01", title: "Brainstorming", desc: "把需求聊透，拒绝想当然" },
    { num: "02", title: "Git Worktrees", desc: "创建安全隔离沙箱" },
    { num: "03", title: "Writing Plans", desc: "拆解为 2-5 分钟原子任务" },
    { num: "04", title: "SubAgent Dev", desc: "子代理专注执行" },
    { num: "05", title: "TDD", desc: "先写测试，再写代码" },
    { num: "06", title: "Code Review", desc: "两阶段审查确保质量" },
  ];

  steps.forEach((s, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.6 + col * 4.7;
    const y = 1.7 + row * 1.25;

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.3, h: 1.0, fill: { color: C.lightBg }, shadow: cardShadow(),
    });
    slide.addText(s.num, {
      x: x + 0.15, y: y + 0.1, w: 0.6, h: 0.6, fontSize: 22, fontFace: "Arial",
      color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(s.title, {
      x: x + 0.9, y: y + 0.1, w: 3.0, h: 0.35, fontSize: 14, fontFace: "Arial",
      color: C.textDark, bold: true, valign: "middle", margin: 0,
    });
    slide.addText(s.desc, {
      x: x + 0.9, y: y + 0.5, w: 3.0, h: 0.35, fontSize: 11, fontFace: "Calibri",
      color: C.textGray, valign: "middle", margin: 0,
    });
  });
  addPageNumber(slide, 24, totalSlides);
})();

/* ---- SLIDE 25: Brainstorming ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("Step 1: Brainstorming（头脑风暴）", { ...slideTitleOpts(0.25), fontSize: 24 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 0.95, w: 1.2, h: 0.04, fill: { color: C.accent },
  });
  slide.addText("核心目标：在写代码前，通过苏格拉底式对话，将模糊需求转化为清晰的设计文档", {
    x: 0.6, y: 1.2, w: 8.8, h: 0.5, fontSize: 13, fontFace: "Calibri",
    color: C.textDark, margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.8, w: 5.0, h: 2.5, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.8, w: 5.0, h: 0.45, fill: { color: C.primary },
  });
  slide.addText("实盘案例：优惠券核销 API", {
    x: 0.8, y: 1.82, w: 4.6, h: 0.4, fontSize: 14, fontFace: "Arial",
    color: C.white, bold: true, valign: "middle", margin: 0,
  });
  slide.addText([
    { text: '需求："实现一个优惠券核销 API"', options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "AI 追问：优惠券有几种类型？", options: { bullet: true, breakLine: true } },
    { text: "AI 追问：一张券能用几次？", options: { bullet: true, breakLine: true } },
    { text: "AI 追问：并发场景怎么处理？", options: { bullet: true } },
  ], { x: 0.8, y: 2.35, w: 4.6, h: 1.8, fontSize: 11, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addText("方案对比", {
    x: 5.8, y: 1.8, w: 3.8, h: 0.35, fontSize: 14, fontFace: "Arial",
    color: C.primary, bold: true, margin: 0,
  });
  const approaches = [
    { name: "Redis 分布式锁", good: "✓ 强一致，性能好", bad: "✗ 引入新依赖" },
    { name: "数据库乐观锁", good: "✓ 实现简单", bad: "✗ 高并发重试多" },
    { name: "数据库悲观锁", good: "✓ 强一致", bad: "✗ 锁粒度粗" },
  ];
  approaches.forEach((a, i) => {
    const y = 2.25 + i * 0.6;
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.8, y, w: 3.8, h: 0.5, fill: { color: C.lightBg } });
    slide.addText(a.name, {
      x: 5.95, y: y + 0.02, w: 3.5, h: 0.22, fontSize: 11, fontFace: "Arial",
      color: C.textDark, bold: true, margin: 0,
    });
    slide.addText(`${a.good}  ${a.bad}`, {
      x: 5.95, y: y + 0.24, w: 3.5, h: 0.22, fontSize: 10, fontFace: "Calibri",
      color: C.textGray, margin: 0,
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.5, w: 8.8, h: 0.6, fill: { color: C.highlight },
  });
  slide.addText("📋 产出物：详细设计文档（数据模型 + API 端点 + 权限控制），经确认后才进入下一步", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.6, fontSize: 12, fontFace: "Calibri",
    color: C.primary, valign: "middle", margin: 0,
  });
  addPageNumber(slide, 25, totalSlides);
})();

/* ---- SLIDE 26: Worktrees & Plans ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("Step 2 & 3: Worktrees + Writing Plans", { ...slideTitleOpts(0.25), fontSize: 24 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 0.95, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.2, h: 3.5, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.2, h: 0.5, fill: { color: C.primary },
  });
  slide.addText("Git Worktrees（安全沙箱）", {
    x: 0.8, y: 1.32, w: 3.8, h: 0.45, fontSize: 14, fontFace: "Arial",
    color: C.white, bold: true, valign: "middle", margin: 0,
  });
  slide.addText([
    { text: "创建独立工作区，隔离开发环境", options: { bullet: true, breakLine: true } },
    { text: "不污染主分支代码库", options: { bullet: true, breakLine: true } },
    { text: "搞砸了直接删除，主分支无损", options: { bullet: true, breakLine: true } },
    { text: "自动初始化环境 & 测试基线", options: { bullet: true } },
  ], { x: 0.8, y: 1.95, w: 3.8, h: 2.5, fontSize: 12, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.3, w: 4.2, h: 3.5, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.3, w: 4.2, h: 0.5, fill: { color: C.accent },
  });
  slide.addText("Writing Plans（拆解任务）", {
    x: 5.4, y: 1.32, w: 3.8, h: 0.45, fontSize: 14, fontFace: "Arial",
    color: C.white, bold: true, valign: "middle", margin: 0,
  });
  slide.addText([
    { text: "将设计拆解为 2-5 分钟原子任务", options: { bullet: true, breakLine: true } },
    { text: "每个任务有明确输入/输出/验证", options: { bullet: true, breakLine: true } },
    { text: "精确到文件路径和每一步操作", options: { bullet: true, breakLine: true } },
    { text: "禁止使用 TBD 等占位符", options: { bullet: true } },
  ], { x: 5.4, y: 1.95, w: 3.8, h: 2.5, fontSize: 12, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.4, y: 3.4, w: 3.8, h: 1.2, fill: { color: C.cardBg }, shadow: cardShadow(),
  });
  slide.addText("Task 示例", {
    x: 5.55, y: 3.45, w: 3.5, h: 0.25, fontSize: 11, fontFace: "Arial",
    color: C.accent, bold: true, margin: 0,
  });
  slide.addText("定义 CouponRedeemRequest DTO\n→ 写测试 → 运行 → 实现 → 提交", {
    x: 5.55, y: 3.75, w: 3.5, h: 0.7, fontSize: 10, fontFace: "Consolas",
    color: C.textGray, margin: 0,
  });
  addPageNumber(slide, 26, totalSlides);
})();

/* ---- SLIDE 27: SubAgent + TDD ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("Step 4 & 5: SubAgent + TDD", { ...slideTitleOpts(0.25), fontSize: 24 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 0.95, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.2, h: 3.5, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.2, h: 0.5, fill: { color: C.primary },
  });
  slide.addText("SubAgent 驱动开发", {
    x: 0.8, y: 1.32, w: 3.8, h: 0.45, fontSize: 14, fontFace: "Arial",
    color: C.white, bold: true, valign: "middle", margin: 0,
  });
  slide.addText([
    { text: "主 Agent 读取计划，提取所有 Task", options: { bullet: true, breakLine: true } },
    { text: "为每个 Task 派遣全新子代理", options: { bullet: true, breakLine: true } },
    { text: "子代理只知当前任务，无历史包袱", options: { bullet: true, breakLine: true } },
    { text: "完成后报告状态，继续下一任务", options: { bullet: true, breakLine: true } },
    { text: "发现问题时记录下来并继续", options: { bullet: true } },
  ], { x: 0.8, y: 1.95, w: 3.8, h: 2.5, fontSize: 11, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.3, w: 4.2, h: 3.5, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.3, w: 4.2, h: 0.5, fill: { color: C.accent },
  });
  slide.addText("TDD：测试驱动开发（铁律）", {
    x: 5.4, y: 1.32, w: 3.8, h: 0.45, fontSize: 14, fontFace: "Arial",
    color: C.white, bold: true, valign: "middle", margin: 0,
  });
  slide.addText([
    { text: "铁律：先写测试，再写代码", options: { bullet: true, breakLine: true } },
    { text: "没有失败的测试不许写代码", options: { bullet: true, breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "红 → 绿 → 重构 循环", options: { bullet: true, breakLine: true } },
    { text: "每行生产代码都有测试覆盖", options: { bullet: true } },
  ], { x: 5.4, y: 1.95, w: 3.8, h: 2.5, fontSize: 12, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 3.8, w: 1.1, h: 0.5, fill: { color: C.red } });
  slide.addText("红", {
    x: 5.4, y: 3.8, w: 1.1, h: 0.5, fontSize: 16, fontFace: "Arial",
    color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.65, y: 3.8, w: 1.1, h: 0.5, fill: { color: C.green } });
  slide.addText("绿", {
    x: 6.65, y: 3.8, w: 1.1, h: 0.5, fontSize: 16, fontFace: "Arial",
    color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, { x: 7.9, y: 3.8, w: 1.1, h: 0.5, fill: { color: C.primary } });
  slide.addText("重构", {
    x: 7.9, y: 3.8, w: 1.1, h: 0.5, fontSize: 16, fontFace: "Arial",
    color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  addPageNumber(slide, 27, totalSlides);
})();

/* ---- SLIDE 28: Code Review & 实践效果 (实用命令已移除，移到 Section 03) ---- */
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("Step 6: Code Review & 实践效果", { ...slideTitleOpts(0.25), fontSize: 24 });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 0.95, w: 1.2, h: 0.04, fill: { color: C.accent },
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.2, h: 2.0, fill: { color: C.lightBg }, shadow: cardShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.2, h: 0.5, fill: { color: C.primary },
  });
  slide.addText("Code Review（代码审查）", {
    x: 0.8, y: 1.32, w: 3.8, h: 0.45, fontSize: 14, fontFace: "Arial",
    color: C.white, bold: true, valign: "middle", margin: 0,
  });
  slide.addText([
    { text: "任务完成后自动派遣审查子代理", options: { bullet: true, breakLine: true } },
    { text: "两阶段审查：", options: { bullet: true, breakLine: true } },
    { text: "① 规格合规性", options: { bullet: true, indentLevel: 1, breakLine: true } },
    { text: "② 代码质量", options: { bullet: true, indentLevel: 1 } },
  ], { x: 0.8, y: 1.95, w: 3.8, h: 1.2, fontSize: 12, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.3, w: 4.2, h: 2.0, fill: { color: C.highlight }, shadow: cardShadow(),
  });
  slide.addText("实践结论", {
    x: 5.4, y: 1.32, w: 3.8, h: 0.45, fontSize: 16, fontFace: "Arial",
    color: C.accent, bold: true, valign: "middle", margin: 0,
  });
  slide.addText([
    { text: "使用 Superpowers 后，", options: { breakLine: true } },
    { text: "对话平均“聪明程度”", options: { breakLine: true } },
    { text: "和完成度大幅提升", options: { breakLine: true } },
    { text: "（至少提升 10%，大工程更多）", options: { bold: true } },
  ], { x: 5.4, y: 1.8, w: 3.8, h: 1.3, fontSize: 13, fontFace: "Calibri", color: C.textDark, valign: "top", margin: 0 });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 3.8, w: 8.8, h: 0.7, fill: { color: C.primary },
  });
  slide.addText("流程优先，确保代码更稳而非更快 ———— Superpowers 核心哲学", {
    x: 0.8, y: 3.8, w: 8.4, h: 0.7, fontSize: 16, fontFace: "Arial",
    color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNumber(slide, 28, totalSlides);
})();

// =====================================================================
// SLIDE 29: Q&A
// =====================================================================
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent },
  });
  slide.addShape(pres.shapes.OVAL, {
    x: -1.5, y: -1.5, w: 4, h: 4, fill: { color: C.accent, transparency: 90 },
  });
  slide.addShape(pres.shapes.OVAL, {
    x: 7.5, y: 3.5, w: 4, h: 4, fill: { color: C.accent2, transparency: 85 },
  });
  slide.addText("Q & A", {
    x: 0.8, y: 1.5, w: 8.4, h: 1.2, fontSize: 54, fontFace: "Arial",
    color: C.white, bold: true, align: "center", margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, { x: 4.2, y: 2.7, w: 1.6, h: 0.04, fill: { color: C.accent } });
  slide.addText("感谢聆听 · 欢迎交流", {
    x: 0.8, y: 3.0, w: 8.4, h: 0.6, fontSize: 18, fontFace: "Calibri",
    color: C.textLight, align: "center", margin: 0,
  });
})();

// =====================================================================
// SLIDE 30: THANK YOU
// =====================================================================
(() => {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  slide.addText("THANK YOU", {
    x: 0.8, y: 1.8, w: 8.4, h: 1.0, fontSize: 44, fontFace: "Arial",
    color: C.primary, bold: true, align: "center", margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, { x: 4.0, y: 2.9, w: 2.0, h: 0.04, fill: { color: C.accent } });
  slide.addText("一起拥抱 AI，成为会用 AI 的人", {
    x: 0.8, y: 3.2, w: 8.4, h: 0.6, fontSize: 16, fontFace: "Calibri",
    color: C.textGray, align: "center", margin: 0,
  });
  slide.addText(`AI大矩阵分享 · ${new Date().toLocaleDateString("zh-CN")}`, {
    x: 0.8, y: 4.5, w: 8.4, h: 0.4, fontSize: 12, fontFace: "Calibri",
    color: C.textLight, align: "center", margin: 0,
  });
})();

// ========== WRITE ==========
const outPath = "C:/Users/20561/Desktop/AI大矩阵分享_v2.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log(`✅ PPT created: ${outPath}`);
}).catch(err => {
  console.error("❌ Error:", err);
});
