const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
const C = { darkBg: "1B2A4A", primary: "2B5A8C", accent: "00A3C4", accent2: "0BC5EA",
  lightBg: "F5F7FA", white: "FFFFFF", textDark: "1A202C", textGray: "4A5568",
  textLight: "A0AEC0", cardBg: "FFFFFF", cardBorder: "E2E8F0", highlight: "EBF8FF",
  green: "38A169", orange: "DD6B20", red: "E53E3E" };
const cardShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.10 });

/* SLIDE 1: Section divider */
const s1 = pres.addSlide();
s1.background = { color: C.darkBg };
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: C.accent } });
s1.addText("03", { x: 0.8, y: 1.2, w: 2, h: 1.2, fontSize: 72, fontFace: "Arial", color: C.accent, bold: true, margin: 0 });
s1.addText("实用命令 & 内置 Skills", { x: 0.8, y: 2.5, w: 8.4, h: 1.0, fontSize: 36, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
s1.addText("Claude Code 命令速查 · 办公/开发类 Skills", { x: 0.8, y: 3.5, w: 8.4, h: 0.6, fontSize: 16, fontFace: "Calibri", color: C.textLight, margin: 0 });
s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.8, w: 3, h: 0.04, fill: { color: C.accent } });

/* SLIDE 2: 实用命令速查 */
const s2 = pres.addSlide();
s2.background = { color: C.white };
s2.addText("Claude Code 实用命令速查", { x: 0.6, y: 0.3, w: 8.8, h: 0.7, fontSize: 26, fontFace: "Arial", color: C.textDark, bold: true, margin: 0 });
s2.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent } });
const groups = [
  { title: "对话控制", cmds: [{ cmd: "/debug", desc: "进入调试模式，查看 AI 内部推理" }, { cmd: "/compact", desc: "压缩上下文，节省 Token" }, { cmd: "/clear", desc: "清空当前对话历史" }] },
  { title: "输出控制", cmds: [{ cmd: "双 Esc", desc: "立即中断 AI 输出" }, { cmd: "/model", desc: "切换当前模型" }, { cmd: "cc - switch", desc: "系统配置切换" }] },
  { title: "实用技巧", cmds: [{ cmd: "/init", desc: "初始化项目 CLAUDE.md" }, { cmd: "/review", desc: "对当前分支进行代码审查" }, { cmd: "/compact + /clear", desc: "组合使用清理上下文" }] },
];
groups.forEach((g, gi) => {
  const x = 0.4 + gi * 3.2;
  s2.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 3.0, h: 0.4, fill: { color: C.primary } });
  s2.addText(g.title, { x, y: 1.3, w: 3.0, h: 0.4, fontSize: 13, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  g.cmds.forEach((c, i) => {
    const y = 1.8 + i * 0.95;
    s2.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.0, h: 0.8, fill: { color: C.lightBg } });
    s2.addText(c.cmd, { x: x + 0.15, y: y + 0.05, w: 2.7, h: 0.3, fontSize: 12, fontFace: "Consolas", color: C.accent, bold: true, margin: 0 });
    s2.addText(c.desc, { x: x + 0.15, y: y + 0.35, w: 2.7, h: 0.35, fontSize: 10, fontFace: "Calibri", color: C.textGray, margin: 0 });
  });
});

/* SLIDE 3: Skills 概览 */
const s3 = pres.addSlide();
s3.background = { color: C.white };
s3.addText("Claude Code 内置 Skills 概览", { x: 0.6, y: 0.3, w: 8.8, h: 0.7, fontSize: 26, fontFace: "Arial", color: C.textDark, bold: true, margin: 0 });
s3.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent } });
s3.addText("Claude Code 通过 Skills 机制扩展 AI 能力，覆盖办公、开发、设计等场景。", { x: 0.6, y: 1.2, w: 8.8, h: 0.4, fontSize: 13, fontFace: "Calibri", color: C.textGray, margin: 0 });
const skills = [
  { icon: "📊", name: "xlsx", desc: "Excel / CSV 文件处理", cat: "办公", catClr: C.primary },
  { icon: "📽", name: "pptx", desc: "PowerPoint 演示文稿制作", cat: "办公", catClr: C.primary },
  { icon: "📝", name: "docx", desc: "Word 文档处理与编辑", cat: "办公", catClr: C.primary },
  { icon: "📄", name: "pdf", desc: "PDF 文件读取与处理", cat: "办公", catClr: C.primary },
  { icon: "🎨", name: "frontend-design", desc: "前端界面设计与交互", cat: "开发", catClr: C.accent },
  { icon: "🔍", name: "code-review", desc: "代码审查与质量检查", cat: "开发", catClr: C.accent },
];
skills.forEach((sk, i) => {
  const col = i % 3, row = Math.floor(i / 3);
  const x = 0.6 + col * 3.1, y = 1.7 + row * 1.6;
  s3.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.8, h: 1.35, fill: { color: C.lightBg }, shadow: cardShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.8, h: 0.06, fill: { color: sk.catClr } });
  s3.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: y + 0.15, w: 0.6, h: 0.22, fill: { color: sk.catClr } });
  s3.addText(sk.cat, { x: x + 0.15, y: y + 0.15, w: 0.6, h: 0.22, fontSize: 8, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s3.addText(sk.icon, { x: x + 0.15, y: y + 0.5, w: 0.4, h: 0.4, fontSize: 18, margin: 0 });
  s3.addText(sk.name, { x: x + 0.6, y: y + 0.5, w: 2.0, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.textDark, bold: true, margin: 0 });
  s3.addText(sk.desc, { x: x + 0.15, y: y + 0.9, w: 2.5, h: 0.35, fontSize: 10, fontFace: "Calibri", color: C.textGray, margin: 0 });
});

/* SLIDE 4: 更多推荐 */
const s4 = pres.addSlide();
s4.background = { color: C.white };
s4.addText("更多实用 Skills 推荐", { x: 0.6, y: 0.3, w: 8.8, h: 0.7, fontSize: 26, fontFace: "Arial", color: C.textDark, bold: true, margin: 0 });
s4.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.0, w: 1.2, h: 0.04, fill: { color: C.accent } });
const recs = [
  { cat: "🛠 开发效率", items: [{ n: "playwright-cli", d: "浏览器自动化测试，Web 交互验证" }, { n: "security-review", d: "安全审查，发现潜在漏洞" }, { n: "brand-guidelines", d: "品牌风格指南，保持设计一致" }] },
  { cat: "🧠 AI 进阶", items: [{ n: "claude-api", d: "Claude API 集成开发，构建 AI 应用" }, { n: "webapp-testing", d: "Web 应用端到端测试" }, { n: "mcp-builder", d: "构建 MCP 服务器，扩展 AI 能力" }] },
  { cat: "🎯 工作流", items: [{ n: "test-driven-dev", d: "TDD 测试驱动开发流程" }, { n: "subagent-driven-dev", d: "子代理驱动开发，任务分治" }, { n: "systematic-debugging", d: "系统化调试流程" }] },
];
recs.forEach((g, gi) => {
  const x = 0.4 + gi * 3.2;
  s4.addText(g.cat, { x, y: 1.2, w: 3.0, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.primary, bold: true, margin: 0 });
  g.items.forEach((item, ii) => {
    const y = 1.65 + ii * 1.1;
    s4.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.0, h: 0.95, fill: { color: C.lightBg }, shadow: cardShadow() });
    s4.addText(item.n, { x: x + 0.12, y: y + 0.08, w: 2.76, h: 0.28, fontSize: 12, fontFace: "Consolas", color: C.accent, bold: true, margin: 0 });
    s4.addText(item.d, { x: x + 0.12, y: y + 0.4, w: 2.76, h: 0.45, fontSize: 10, fontFace: "Calibri", color: C.textGray, margin: 0 });
  });
});

pres.writeFile({ fileName: "C:/Users/20561/Desktop/new_slides.pptx" }).then(() => console.log("OK")).catch(e => console.error(e));
