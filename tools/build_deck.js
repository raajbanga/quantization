// Builds presentation.pptx.  node build_deck.js <repo_dir> <assets_dir>
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const SKILL = process.env.PPTX_SKILL;
const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));

const REPO = process.argv[2], ASSETS = process.argv[3];
const OUTFILE = path.join(REPO, "presentation.pptx");

// ---------- Design tokens ----------
const HEX = {
  paper: "F4F5F7", ink: "16161A", muted: "6E6E78", rule: "D9DAE0",
  accent: "3B3BF2", tint: "E4E4FD", alarm: "E5484D", white: "FFFFFF",
};
const THEME = {
  name: "Fewer Bits", headFontFace: "Helvetica Neue", bodyFontFace: "Helvetica Neue",
  colors: { dk1: HEX.ink, lt1: HEX.white, dk2: HEX.muted, lt2: HEX.paper,
    accent1: HEX.accent, accent2: HEX.alarm, accent3: HEX.tint, accent4: HEX.ink,
    accent5: HEX.muted, accent6: HEX.rule, hlink: HEX.accent, folHlink: HEX.muted },
};

// ---------- Results from the notebook ----------
function readCsv(p) {
  const [head, ...lines] = fs.readFileSync(p, "utf8").trim().split("\n");
  const keys = head.split(",");
  return lines.map(l => Object.fromEntries(l.split(",").map((v, i) => [keys[i], isNaN(+v) ? v : +v])));
}
const R = readCsv(path.join(REPO, "results/results.csv"));
const short = R.map(r => r.level.split(" ")[0]);

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Raaj Banga";
pres.title = "Fewer Bits, Same Brain";
const C = pres.SchemeColor;
const W = 13.333, M = 0.65;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "COVER", background: { color: HEX.paper },
  objects: [],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: HEX.paper },
  margin: [0.5, M, 0.6, M],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.45, w: W - 2 * M, h: 0.85,
      fontSize: 32, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: W - M - 0.6, y: 6.95, w: 0.6, h: 0.3, fontSize: 10, color: C.text2, align: "right" },
});

// ---------- Helpers ----------
const TOTAL = 18;
const notes = {};
let n = 0;
const declared = new Set(["Intro"]);
function slideOf(title, section) {
  n += 1;
  if (!declared.has(section)) { pres.addSection({ title: section }); declared.add(section); }
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  progress(s, n);
  return s;
}
// Progress motif: one small square per slide, filled up to the current one ("bits" filling in).
function progress(s, i) {
  for (let k = 1; k <= TOTAL; k++) {
    s.addShape(pres.shapes.RECTANGLE, { x: W - M - 0.75 - (TOTAL - k + 1) * 0.16, y: 7.04, w: 0.1, h: 0.1,
      fill: { color: k <= i ? C.accent1 : C.accent6 }, line: { type: "none" }, objectName: `progress-${k}` });
  }
}
function cite(s, text) {
  s.addText(text, { x: M, y: 6.92, w: 8.6, h: 0.32, fontSize: 10, color: C.text2, margin: 0, isTextBox: true, valign: "middle" });
}
function txt(s, text, o) { s.addText(text, Object.assign({ margin: 0, isTextBox: true, color: C.text1, valign: "top" }, o)); }
function box(s, o) { s.addShape(pres.shapes.RECTANGLE, Object.assign({ line: { type: "none" } }, o)); }
function bits(s, x, y, count, size, gap, color) {
  for (let k = 0; k < count; k++) box(s, { x: x + k * (size + gap), y, w: size, h: size, fill: { color } });
}
// Shape-drawn charts (render identically in PowerPoint and Keynote)
function hbars(s, x, y, labelW, barW, rowH, rows, max, fmt) {
  rows.forEach((r, i) => {
    const yy = y + i * rowH, bw = Math.max(0.03, barW * r.v / max);
    txt(s, r.label, { x, y: yy, w: labelW - 0.15, h: rowH * 0.8, fontSize: 15, bold: !!r.bold, color: r.labelColor || C.text1, valign: "middle" });
    box(s, { x: x + labelW, y: yy + rowH * 0.18, w: bw, h: rowH * 0.44, fill: { color: r.color } });
    txt(s, fmt(r.v), { x: x + labelW + bw + 0.1, y: yy, w: 1.6, h: rowH * 0.8, fontSize: 14, color: r.labelColor || C.text1, bold: !!r.bold, valign: "middle" });
  });
}
function seg(s, x1, y1, x2, y2, color, width) {
  s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001),
    flipV: (y2 - y1) * (x2 - x1) < 0, line: { color, width } });
}
const axisStyle = () => ({
  catAxisLabelColor: HEX.ink, valAxisLabelColor: HEX.muted, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 12,
  dataLabelColor: HEX.ink, valGridLine: { color: HEX.rule, size: 0.5 }, catGridLine: { style: "none" },
  catAxisLineShow: false, valAxisLineShow: false, legendFontFace: "+mn-lt", legendFontSize: 12, legendColor: HEX.ink,
});

// =====================================================================
// 1 · Cover
pres.addSection({ title: "Intro" });
n += 1;
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
  txt(s, "Fewer bits,\nsame brain", { x: M, y: 1.55, w: 7, h: 2.6, fontSize: 66, bold: true, fontFace: THEME.headFontFace, lineSpacingMultiple: 0.9 });
  txt(s, "How quantization lets big language models run on small machines", { x: M, y: 4.25, w: 6.4, h: 0.9, fontSize: 20, color: C.text2 });
  txt(s, "Raaj Banga   |   November 2026", { x: M, y: 6.55, w: 6, h: 0.35, fontSize: 13, color: C.text2 });
  // Motif: a number stored with 16 bits vs 4 bits
  const x0 = 8.0, sz = 0.26, g = 0.06;
  txt(s, "one weight, stored with 16 bits", { x: x0, y: 2.45, w: 5, h: 0.3, fontSize: 13, color: C.text2 });
  bits(s, x0, 2.85, 16, sz * 0.9, g * 0.6, HEX.ink);
  txt(s, "the same weight, with 4", { x: x0, y: 3.75, w: 5, h: 0.3, fontSize: 13, color: C.text2 });
  bits(s, x0, 4.15, 4, sz * 0.9, g * 0.6, HEX.accent);
  txt(s, "4× less memory", { x: x0 + 4 * (sz * 0.9 + g * 0.6) + 0.2, y: 4.1, w: 3, h: 0.35, fontSize: 16, bold: true, color: C.accent1 });
  progress(s, n);
  notes[n] = "Hook. Quantization = storing each number in a model with fewer bits. Goal for today: intuition, not math.";
}

// 2 · Motivation
{
  const s = slideOf("The best open models don't fit on your laptop", "Intro");
  txt(s, "Memory needed to hold each model at 16-bit (B in a name = billion weights)", { x: M, y: 1.55, w: 7.5, h: 0.35, fontSize: 14, color: C.text2 });
  hbars(s, M, 2.05, 2.25, 4.3, 0.78, [
    { label: "Qwen3-4B", v: 8, color: HEX.ink },
    { label: "This MacBook", v: 24, color: HEX.accent, labelColor: C.accent1, bold: true },
    { label: "Llama 3.1 70B", v: 141, color: HEX.ink },
    { label: "gpt-oss-120b", v: 234, color: HEX.ink },
    { label: "DeepSeek-V3", v: 1342, color: HEX.ink },
  ], 1342, v => `${v.toLocaleString("en-US")} GB`);
  txt(s, "1,342 GB", { x: 9.2, y: 1.7, w: 3.5, h: 0.9, fontSize: 48, bold: true });
  txt(s, "to hold DeepSeek-V3 (671 billion numbers) at 16 bits each", { x: 9.2, y: 2.6, w: 3.5, h: 0.8, fontSize: 15, color: C.text2 });
  txt(s, "24 GB", { x: 9.2, y: 3.85, w: 3.5, h: 0.9, fontSize: 48, bold: true, color: C.accent1 });
  txt(s, "of memory in the MacBook running today's demo", { x: 9.2, y: 4.75, w: 3.5, h: 0.8, fontSize: 15, color: C.text2 });
  cite(s, "Parameter counts: Qwen Team (2025); Meta AI (2024); OpenAI (2025); DeepSeek-AI (2024). Memory = parameters × 2 bytes.");
  notes[n] = "Size = number of parameters x 2 bytes. Frontier-class open models are 50x bigger than a laptop's memory.";
}

// 3 · Background
{
  const s = slideOf("A language model is a giant list of numbers", "Background");
  // flow: words -> numbers -> word
  txt(s, "“The cat sat on the”", { x: M, y: 1.95, w: 3.2, h: 0.6, fontSize: 20, valign: "middle" });
  box(s, { x: 4.15, y: 1.6, w: 4.4, h: 1.3, fill: { color: HEX.ink } });
  txt(s, "4,000,000,000 weights", { x: 4.15, y: 1.6, w: 4.4, h: 1.3, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle" });
  s.addShape(pres.shapes.LINE, { x: 3.7, y: 2.25, w: 0.35, h: 0, line: { color: HEX.muted, width: 1.5, endArrowType: "triangle" } });
  s.addShape(pres.shapes.LINE, { x: 8.65, y: 2.25, w: 0.35, h: 0, line: { color: HEX.muted, width: 1.5, endArrowType: "triangle" } });
  txt(s, "“mat”", { x: 9.15, y: 1.95, w: 2, h: 0.6, fontSize: 20, bold: true, color: C.accent1, valign: "middle" });
  txt(s, "...and repeat for every token (a word or piece of a word)", { x: 4.15, y: 3.0, w: 4.4, h: 0.4, fontSize: 13, color: C.text2, align: "center" });
  // three facts
  const facts = [["4 billion", "weights in Qwen3-4B, the model in today's demo"], ["× 2 bytes", "each: 16 bits (a bit is a 0 or 1; 8 bits = 1 byte)"], ["= 8 GB", "read from memory for every token it writes"]];
  facts.forEach(([big, small], i) => {
    const x = M + i * 4.0;
    txt(s, big, { x, y: 4.05, w: 3.7, h: 0.75, fontSize: 36, bold: true, color: i === 2 ? C.accent1 : C.text1 });
    txt(s, small, { x, y: 4.85, w: 3.5, h: 0.7, fontSize: 15, color: C.text2 });
  });
  txt(s, "The slow part isn't the math. It's moving all those numbers through memory, over and over.", { x: M, y: 5.95, w: 11.5, h: 0.5, fontSize: 18, italic: true });
  cite(s, "Qwen Team (2025).");
  notes[n] = "Model = weights. Every generated word reads every weight once. Memory size and memory speed are the bottleneck.";
}

// 4 · Core concept: buckets
{
  const s = slideOf("Quantization means rounding into fewer buckets", "Core idea");
  const rows = [["16-bit", 65536, "65,536 buckets"], ["8-bit", 256, "256"], ["4-bit", 16, "16"], ["2-bit", 4, "4"]];
  const rx = M + 1.15, rw = 5.3;
  rows.forEach(([label, k, cap], i) => {
    const y = 1.75 + i * 1.05;
    txt(s, label, { x: M, y: y - 0.05, w: 1.1, h: 0.4, fontSize: 16, bold: true, color: i === 2 ? C.accent1 : C.text1 });
    s.addShape(pres.shapes.LINE, { x: rx, y: y + 0.32, w: rw, h: 0, line: { color: HEX.muted, width: 1 } });
    const ticks = Math.min(k, 64);
    for (let t = 0; t < ticks; t++) {
      const x = rx + (rw * t) / (ticks - 1);
      s.addShape(pres.shapes.LINE, { x, y: y, w: 0, h: 0.32, line: { color: i === 2 ? HEX.accent : HEX.ink, width: k > 64 ? 2.2 : 1.25 } });
    }
    txt(s, cap, { x: rx, y: y + 0.4, w: rw, h: 0.3, fontSize: 11, color: C.text2 });
  });
  // recipe (a real sequence)
  const sx = 7.55;
  txt(s, "The recipe", { x: sx, y: 1.6, w: 5, h: 0.4, fontSize: 18, bold: true });
  const steps = [
    ["1", "Find the biggest number", "It sets how wide each bucket is (the scale)."],
    ["2", "Round every number to a bucket", "Store just the small bucket number."],
    ["3", "Multiply back when running", "bucket × scale ≈ the original weight"],
  ];
  steps.forEach(([k, h, d], i) => {
    const y = 2.2 + i * 1.0;
    s.addShape(pres.shapes.OVAL, { x: sx, y: y + 0.02, w: 0.42, h: 0.42, fill: { color: HEX.accent }, line: { type: "none" } });
    txt(s, k, { x: sx, y: y + 0.02, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, h, { x: sx + 0.6, y, w: 4.6, h: 0.35, fontSize: 16, bold: true });
    txt(s, d, { x: sx + 0.6, y: y + 0.38, w: 4.6, h: 0.35, fontSize: 14, color: C.text2 });
  });
  box(s, { x: sx, y: 5.3, w: 5.1, h: 1.15, fill: { color: HEX.white } });
  txt(s, [
    { text: "0.0374", options: { bold: true } }, { text: "  ÷ 0.02 = 1.87  →  rounds to " },
    { text: "2", options: { bold: true, color: HEX.accent, breakLine: true } }, { text: "2 × 0.02 = " },
    { text: "0.04", options: { bold: true } }, { text: "   (off by 0.0026)", options: { color: HEX.muted } },
  ], { x: sx + 0.25, y: 5.42, w: 4.7, h: 0.95, fontSize: 15, valign: "middle", lineSpacingMultiple: 1.2 });
  notes[n] = "Fewer bits = fewer buckets. The recipe is the same in every method; research is about choosing buckets well.";
}

// 5 · Picture analogy
{
  const s = slideOf("Pictures survive rounding surprisingly well", "Core idea");
  const pics = [[8, "256 levels per colour"], [4, "16 levels"], [2, "4 levels"], [1, "2 levels"]];
  const pw = 2.85, ph = 1.9, gap = 0.17;
  pics.forEach(([b, cap], i) => {
    const x = M + i * (pw + gap);
    s.addImage({ path: path.join(ASSETS, `pic_${b}bit.png`), x, y: 1.75, w: pw, h: ph, altText: `Landscape image rounded to ${b}-bit colour` });
    txt(s, `${b}-bit`, { x, y: 3.8, w: pw, h: 0.4, fontSize: 18, bold: true, color: b === 4 ? C.accent1 : C.text1 });
    txt(s, cap, { x, y: 4.2, w: pw, h: 0.3, fontSize: 13, color: C.text2 });
  });
  txt(s, "An image is a grid of numbers too. At 4 bits it still reads as the same picture, because a pixel that's a little off is invisible.",
    { x: M, y: 5.05, w: 9.5, h: 0.8, fontSize: 18 });
  txt(s, "Hold onto this. It's why image generators handle quantization so well.", { x: M, y: 5.9, w: 9.5, h: 0.45, fontSize: 18, bold: true, color: C.accent1 });
  cite(s, "Generated in quantization_demo.ipynb, Part 2.");
  notes[n] = "Visual intuition for bits. Small errors spread across many values are fine for images.";
}

// 6 · Outliers
{
  const s = slideOf("One giant number can ruin the ruler for everyone", "Core idea");
  // Panel A: one ruler
  const drawPanel = (x0, title, blocks) => {
    txt(s, title, { x: x0, y: 1.6, w: 5.6, h: 0.4, fontSize: 17, bold: true });
    const y = 2.55, w = 5.5;
    if (!blocks) {
      s.addShape(pres.shapes.LINE, { x: x0, y: y + 0.45, w, h: 0, line: { color: HEX.muted, width: 1 } });
      for (let t = 0; t < 7; t++) s.addShape(pres.shapes.LINE, { x: x0 + (w * t) / 6, y: y + 0.3, w: 0, h: 0.3, line: { color: HEX.ink, width: 1.25 } });
      for (let d = 0; d < 22; d++) s.addShape(pres.shapes.OVAL, { x: x0 + w / 2 - 0.35 + (d % 11) * 0.06, y: y - 0.05 - Math.floor(d / 11) * 0.12, w: 0.08, h: 0.08, fill: { color: HEX.ink }, line: { type: "none" } });
      s.addShape(pres.shapes.OVAL, { x: x0 + w - 0.1, y: y - 0.05, w: 0.14, h: 0.14, fill: { color: HEX.alarm }, line: { type: "none" } });
      txt(s, "outlier", { x: x0 + w - 0.9, y: y - 0.45, w: 1, h: 0.3, fontSize: 11, color: HEX.alarm, align: "right" });
      txt(s, "Buckets get so wide that every normal weight rounds to the middle one, which is zero.", { x: x0, y: 3.35, w: w, h: 0.75, fontSize: 14, color: C.text2 });
    } else {
      const bw = (w - 0.3) / 4;
      for (let b = 0; b < 4; b++) {
        const bx = x0 + b * (bw + 0.1);
        s.addShape(pres.shapes.LINE, { x: bx, y: y + 0.45, w: bw, h: 0, line: { color: HEX.muted, width: 1 } });
        for (let t = 0; t < 7; t++) s.addShape(pres.shapes.LINE, { x: bx + (bw * t) / 6, y: y + 0.3, w: 0, h: 0.3, line: { color: b === 3 ? HEX.alarm : HEX.accent, width: 1.25 } });
        for (let d = 0; d < 5; d++) s.addShape(pres.shapes.OVAL, { x: bx + 0.12 + d * (bw - 0.3) / 4, y: y - 0.05, w: 0.08, h: 0.08, fill: { color: HEX.ink }, line: { type: "none" } });
      }
      txt(s, "Split the weights into small blocks, each with its own ruler. The outlier only spoils its own block.", { x: x0, y: 3.35, w: w, h: 0.75, fontSize: 14, color: C.text2 });
    }
  };
  drawPanel(M, "One ruler for everything", false);
  drawPanel(7.1, "One ruler per block of 32", true);
  txt(s, "99.9%", { x: M, y: 4.55, w: 3.2, h: 0.9, fontSize: 48, bold: true, color: HEX.alarm });
  txt(s, "of 4,096 weights erased to zero at 4-bit, caused by one outlier (a value far bigger than the rest)", { x: M, y: 5.45, w: 5.2, h: 0.7, fontSize: 14, color: C.text2 });
  txt(s, "12.6%", { x: 7.1, y: 4.55, w: 3.2, h: 0.9, fontSize: 48, bold: true, color: C.accent1 });
  txt(s, "erased with block-wise rulers. Every modern 4-bit format uses this trick (next slides).", { x: 7.1, y: 5.45, w: 5.4, h: 0.7, fontSize: 14, color: C.text2 });
  cite(s, "Outliers in LLMs: Dettmers et al. (2022). Numbers: quantization_demo.ipynb, Part 3. Block formats: Rouhani et al. (2023).");
  notes[n] = "Biggest value sets bucket width. Outliers appear in big models (~6.7B+). Fix: per-block scales.";
}

// 7 · Research (two slides, two papers each)
function paperSlide(title, papers, citeText) {
  const s = slideOf(title, "Research");
  const cw = 5.9, gap = 0.33;
  papers.forEach(([name, venue, objective, method, big, bigLabel], i) => {
    const x = M + i * (cw + gap), y = 1.5;
    box(s, { x, y, w: cw, h: 5.15, fill: { color: HEX.white } });
    txt(s, name, { x: x + 0.4, y: y + 0.3, w: cw - 0.8, h: 0.55, fontSize: 26, bold: true });
    txt(s, venue, { x: x + 0.4, y: y + 0.85, w: cw - 0.8, h: 0.35, fontSize: 13, color: C.text2 });
    const row = (label, body, yy, h) => {
      txt(s, label, { x: x + 0.4, y: yy, w: 1.25, h: 0.35, fontSize: 13, bold: true, color: C.text2 });
      txt(s, body, { x: x + 1.7, y: yy, w: cw - 2.1, h, fontSize: 15 });
    };
    row("Objective", objective, y + 1.45, 0.8);
    row("Method", method, y + 2.3, 1.3);
    txt(s, "Result", { x: x + 0.4, y: y + 3.8, w: 1.25, h: 0.35, fontSize: 13, bold: true, color: C.text2 });
    txt(s, big, { x: x + 1.7, y: y + 3.62, w: cw - 2.1, h: 0.7, fontSize: 34, bold: true, color: C.accent1 });
    txt(s, bigLabel, { x: x + 1.7, y: y + 4.32, w: cw - 2.1, h: 0.65, fontSize: 14, color: C.text2 });
  });
  cite(s, citeText);
  return s;
}
paperSlide("Research 1: find the outliers, fix the errors", [
  ["LLM.int8()", "Dettmers et al., NeurIPS 2022",
   "Run huge models in 8-bit without losing quality.",
   "Find the rare outlier values. Keep those in 16-bit; round the other 99.9% to 8-bit.",
   "2×", "less memory, no quality loss up to 175B weights"],
  ["GPTQ", "Frantar et al., ICLR 2023",
   "Shrink models to 3–4 bits quickly, with no retraining.",
   "Round one column of weights at a time, then nudge the not-yet-rounded weights to cancel the error.",
   "4 hours", "to compress a 175B model on one GPU"],
], "NeurIPS, ICLR: top peer-reviewed AI conferences. Dettmers et al. (2022); Frantar et al. (2023).");
notes[n] = "LLM.int8 found outliers and used mixed precision. GPTQ rounds column by column and compensates the error.";
paperSlide("Research 2: smarter buckets, protected weights", [
  ["QLoRA", "Dettmers et al., NeurIPS 2023",
   "Let people fine-tune (keep training) big models on one GPU.",
   "Most weights sit near zero, so put more of the 16 buckets there (NF4). Freeze the 4-bit model; train small add-ons.",
   "48 GB", "one GPU was enough to fine-tune a 65B model"],
  ["AWQ", "Lin et al., MLSys 2024 (Best Paper)",
   "Get 4-bit quality without special hardware.",
   "~1% of weights matter most: the ones that get big inputs. Scale those up before rounding so they keep more detail.",
   "3×", "faster than 16-bit on desktop and mobile GPUs"],
], "MLSys: top peer-reviewed ML-systems conference. Dettmers et al. (2023); Lin et al. (2024).");
notes[n] = "QLoRA: NF4 buckets + adapters. AWQ: protect salient weights by scaling, no mixed precision needed.";

// 8 · Evaluation
{
  const s = slideOf("How do we know a shrunk model still works?", "Evaluation");
  const ms = [
    ["Memory", "GB", "How much space the weights take. Decides whether it fits at all."],
    ["Speed", "tokens / sec", "How fast it writes (a token is about ¾ of a word). Smaller files move faster."],
    ["Perplexity", "lower is better", "How surprised it is by Wikipedia text (WikiText-2). 10 ≈ \u201Cunsure between about 10 words\u201D each step."],
    ["Task accuracy", "% correct", "Easy and hard math, plus short and long JSON tool calls: the exact format AI agents use to run apps."],
  ];
  ms.forEach(([h, unit, d], i) => {
    const x = M + (i % 2) * 6.1, y = 1.65 + Math.floor(i / 2) * 2.15;
    box(s, { x, y, w: 5.8, h: 1.9, fill: { color: HEX.white } });
    txt(s, h, { x: x + 0.35, y: y + 0.3, w: 3.4, h: 0.5, fontSize: 22, bold: true });
    txt(s, unit, { x: x + 3.5, y: y + 0.36, w: 2, h: 0.4, fontSize: 13, color: C.accent1, align: "right" });
    txt(s, d, { x: x + 0.35, y: y + 0.88, w: 5.1, h: 0.9, fontSize: 17, color: C.text2 });
  });
  txt(s, "Watch the gap: perplexity notices small damage early, while tasks keep passing until it's too late.", { x: M, y: 6.1, w: 11.5, h: 0.45, fontSize: 17, bold: true, color: C.accent1 });
  cite(s, "Metrics follow Frantar et al. (2023); Lin et al. (2024). GSM8K: Cobbe et al. (2021). WikiText-2: Merity et al. (2017).");
  notes[n] = "Four things we measure in the demo. Perplexity = exp(average surprise).";
}

// 9 · Decoding a file name
{
  const s = slideOf("How to read a quantized model's name", "Demo");
  const seg = (t, c) => ({ text: t, options: { color: c, bold: c !== HEX.muted } });
  txt(s, [seg("Qwen3", HEX.ink), seg("-", HEX.muted), seg("4B", HEX.ink), seg("-", HEX.muted), seg("Instruct", HEX.ink), seg("-", HEX.muted),
    seg("2507", HEX.ink), seg("-", HEX.muted), seg("Q4", HEX.accent), seg("_", HEX.muted), seg("K", HEX.accent), seg("_", HEX.muted), seg("M", HEX.accent), seg(".gguf", HEX.ink)],
    { x: M, y: 1.45, w: W - 2 * M, h: 0.75, fontSize: 32, fontFace: "Menlo", valign: "middle" });
  const item = (x, y, term, desc, col, w) => {
    txt(s, term, { x, y, w: 1.55, h: 0.45, fontSize: 16, bold: true, fontFace: "Menlo", color: col });
    txt(s, desc, { x: x + 1.6, y: y + 0.03, w: w - 1.6, h: 0.7, fontSize: 15, color: C.text1 });
  };
  txt(s, "About the model", { x: M, y: 2.4, w: 5.5, h: 0.35, fontSize: 14, color: C.text2 });
  [["Qwen3", "Model family, made by Alibaba"], ["4B", "4 billion weights (B = billion)"], ["Instruct", "Tuned to chat and follow instructions"],
   ["2507", "Release date: July 2025"], [".gguf", "File type for llama.cpp, LM Studio, Ollama"]]
    .forEach(([t, d], i) => item(M, 2.85 + i * 0.5, t, d, HEX.ink, 5.9));
  const rx = 6.85;
  txt(s, "About the compression", { x: rx, y: 2.4, w: 5.5, h: 0.35, fontSize: 14, color: C.accent1 });
  [["Q4", "Quantized to about 4 bits per weight", 0.52], ["K", "“k-quant”: small blocks, each with its own ruler (slide 6), grouped into super-blocks of 256", 0.85],
   ["M", "Medium mix: a few sensitive layers keep more bits. S = small (fewer), L = large (more)", 0.85]]
    .reduce((y, [t, d, h]) => { item(rx, y, t, d, HEX.accent, 5.8); return y + h; }, 2.85);
  const others = [["F16", "The original: 16-bit decimal numbers (“floats”)"], ["Q8_0", "8-bit, simplest scheme: one ruler per block of 32"],
    ["IQ1_S", "“i-quant”, ~1.5 bits: snaps weights to a lookup table of patterns"]];
  others.forEach(([t, d], i) => {
    const x = M + i * 4.05;
    box(s, { x, y: 5.68, w: 3.85, h: 1.0, fill: { color: HEX.white } });
    txt(s, t, { x: x + 0.25, y: 5.76, w: 3.4, h: 0.35, fontSize: 15, bold: true, fontFace: "Menlo" });
    txt(s, d, { x: x + 0.25, y: 6.1, w: 3.45, h: 0.52, fontSize: 13, color: C.text2 });
  });
  cite(s, "ggml-org (2025), llama.cpp quantization docs. File names from unsloth/Qwen3-4B-Instruct-2507-GGUF.");
  notes[n] = "Decode the file name left to right. The quantization part is Q4_K_M: bits, scheme, mix.";
}

// 9 · Demo
{
  const s = slideOf("Demo: one model, six sizes", "Demo");
  txt(s, "Goal: shrink the same model step by step and find where it breaks.", { x: M, y: 1.5, w: 11.5, h: 0.45, fontSize: 18 });
  const files = R.map(r => [r.level.split(" ")[0], r.level.match(/\((.*)\)/)[1], r.file_gb]);
  const maxGb = Math.max(...files.map(f => f[2]));
  files.forEach(([name, b, gb], i) => {
    const y = 2.25 + i * 0.56, bw = 5.0 * gb / maxGb;
    txt(s, name, { x: M, y, w: 1.3, h: 0.42, fontSize: 15, bold: true, valign: "middle" });
    txt(s, b, { x: M + 1.3, y, w: 1.0, h: 0.42, fontSize: 12, color: C.text2, valign: "middle" });
    box(s, { x: M + 2.35, y: y + 0.07, w: bw, h: 0.28, fill: { color: name.startsWith("Q4") ? HEX.accent : HEX.ink } });
    txt(s, `${gb.toFixed(1)} GB`, { x: M + 2.45 + bw, y, w: 1.2, h: 0.42, fontSize: 13, valign: "middle" });
  });
  const tx = 9.55;
  txt(s, "Tools", { x: tx, y: 2.2, w: 3.2, h: 0.4, fontSize: 16, bold: true });
  txt(s, [
    { text: "Qwen3-4B-Instruct-2507", options: { breakLine: true } },
    { text: "GGUF files, pre-made by Unsloth", options: { breakLine: true } },
    { text: "llama-cpp-python: runs the files", options: { breakLine: true } },
    { text: "Metal: Apple's GPU interface", options: { breakLine: true } },
    { text: "Jupyter notebook" },
  ], { x: tx, y: 2.65, w: 3.2, h: 1.75, fontSize: 14, color: C.text2, paraSpaceAfter: 4 });
  txt(s, "Tests at every size", { x: tx, y: 4.6, w: 3.2, h: 0.4, fontSize: 16, bold: true });
  txt(s, [
    { text: "Math: 20 easy + 50 GSM8K", options: { breakLine: true } },
    { text: "JSON: 10 short + 8 long calls", options: { breakLine: true } },
    { text: "Perplexity (WikiText-2), speed", options: { breakLine: true } },
    { text: "Memory" },
  ], { x: tx, y: 5.05, w: 3.2, h: 1.1, fontSize: 14, color: C.text2, paraSpaceAfter: 4 });
  cite(s, "Notebook: quantization_demo.ipynb. Model: Qwen Team (2025). Runtime: ggml-org (2025).");
  notes[n] = "Switch to the notebook here.";
}

// 10 · Results
const f16 = R[0], q4 = R.find(r => r.level.startsWith("Q4")), worst = R[R.length - 1];
{
  const s = slideOf("Solid to 3-bit, slipping at 2-bit, broken below", "Results");
  {
    const gx = M + 0.6, gy = 1.95, gw = 7.0, gh = 3.6;
    txt(s, "% correct at each size", { x: M, y: 1.45, w: 4, h: 0.35, fontSize: 14, color: C.text2 });
    [0, 50, 100].forEach(v => {
      const yy = gy + gh * (1 - v / 100);
      s.addShape(pres.shapes.LINE, { x: gx, y: yy, w: gw, h: 0, line: { color: HEX.rule, width: 0.75 } });
      txt(s, `${v}%`, { x: M - 0.1, y: yy - 0.15, w: 0.6, h: 0.3, fontSize: 11, color: C.text2, align: "right" });
    });
    const px = i => gx + 0.3 + i * (gw - 0.6) / (R.length - 1);
    const py = v => gy + gh * (1 - v);
    // Easy tests (20 math + 10 short JSON) scored identically, so they share one faint line.
    const series = [["Easy tests", "math_acc", HEX.rule, 4], ["GSM8K math (50)", "gsm8k_acc", HEX.ink, 3], ["Long JSON (8)", "json_hard_correct", HEX.accent, 3]];
    series.forEach(([name, key, col, wd], si) => {
      R.forEach((r, i) => { if (i) seg(s, px(i - 1), py(R[i - 1][key]), px(i), py(r[key]), col, wd); });
      R.forEach((r, i) => s.addShape(pres.shapes.OVAL, { x: px(i) - 0.08, y: py(r[key]) - 0.08, w: 0.16, h: 0.16, fill: { color: col }, line: { color: HEX.paper, width: 1 } }));
      box(s, { x: M + 2.9 + si * 1.75, y: 1.55, w: 0.3, h: 0.08, fill: { color: col } });
      txt(s, name, { x: M + 3.25 + si * 1.75, y: 1.42, w: 1.5, h: 0.35, fontSize: 12, color: C.text1, valign: "middle" });
    });
    R.forEach((r, i) => {
      txt(s, short[i], { x: px(i) - 0.6, y: gy + gh + 0.15, w: 1.2, h: 0.3, fontSize: 13, bold: true, align: "center" });
      txt(s, `${r.file_gb.toFixed(1)} GB`, { x: px(i) - 0.6, y: gy + gh + 0.45, w: 1.2, h: 0.3, fontSize: 11, color: C.text2, align: "center" });
    });
  }
  const sx = 9.05;
  const pct = (a, b) => Math.round((1 - a / b) * 100);
  const q2 = R.find(r => r.level.startsWith("Q2"));
  txt(s, `${pct(q4.file_gb, f16.file_gb)}%`, { x: sx, y: 1.6, w: 3.6, h: 0.85, fontSize: 44, bold: true, color: C.accent1 });
  txt(s, `smaller and ${(q4.tokens_per_s / f16.tokens_per_s).toFixed(1)}× faster at 4-bit. Perplexity ${f16.wikitext_ppl.toFixed(1)} → ${q4.wikitext_ppl.toFixed(1)}.`, { x: sx, y: 2.45, w: 3.6, h: 0.75, fontSize: 14, color: C.text2 });
  txt(s, `${Math.round(q2.json_hard_correct * 100)}%`, { x: sx, y: 3.4, w: 3.6, h: 0.85, fontSize: 44, bold: true });
  txt(s, "long JSON at 2-bit. Valid JSON, wrong values: O'Hare became \u201CCHI\u201D.", { x: sx, y: 4.25, w: 3.6, h: 0.65, fontSize: 14, color: C.text2 });
  txt(s, `${Math.round(worst.gsm8k_acc * 100)}%`, { x: sx, y: 4.95, w: 3.6, h: 0.85, fontSize: 44, bold: true, color: HEX.alarm });
  txt(s, `on every test at ~1.5-bit. Perplexity ${worst.wikitext_ppl.toFixed(0)}: gibberish.`, { x: sx, y: 5.8, w: 3.6, h: 0.5, fontSize: 14, color: C.text2 });
  cite(s, "Measured on an M4 Pro MacBook (24 GB), quantization_demo.ipynb Part 5. GSM8K: Cobbe et al. (2021). Perplexity: WikiText-2.");
  notes[n] = "Our own results. Easy tests 100% to 2-bit. Long JSON drops to 75% at 2-bit (valid but wrong). GSM8K 88-92% to 2-bit (noise). 1.5-bit: 0%. See results/results.csv.";
}

// Applications today
{
  const s = slideOf("Where quantization is already working today", "Applications");
  const tiles = [
    ["On your phone", "Apple's on-device model runs at 2 bits per weight, trained with the rounding built in.", "Apple (2025)"],
    ["On your laptop", "Ollama and LM Studio offer most models as 4-bit GGUF files by default. That's what “run it locally” means.", "Ollama (2025); Element Labs (2025)"],
    ["On small devices", "AWQ ran Llama 2 70B on a Jetson Orin, a 64 GB computer built for robots and edge devices.", "Lin et al. (2024)"],
    ["In data centers", "8-bit weights and inputs (SmoothQuant): up to 1.56× faster and half the memory, so each answer costs less.", "Xiao et al. (2023)"],
  ];
  tiles.forEach(([h, d, src], i) => {
    const x = M + (i % 2) * 6.1, y = 1.55 + Math.floor(i / 2) * 2.55;
    box(s, { x, y, w: 5.85, h: 2.3, fill: { color: HEX.white } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.35, y: y + 0.42, w: 0.2, h: 0.2, fill: { color: HEX.accent }, line: { type: "none" } });
    txt(s, h, { x: x + 0.7, y: y + 0.3, w: 4.8, h: 0.45, fontSize: 21, bold: true, valign: "middle" });
    txt(s, d, { x: x + 0.35, y: y + 0.9, w: 5.15, h: 0.95, fontSize: 16 });
    txt(s, src, { x: x + 0.35, y: y + 1.85, w: 5.15, h: 0.3, fontSize: 11, color: C.text2 });
  });
  cite(s, "Sources under each example.");
  notes[n] = "Phone (Apple 2-bit QAT), laptop (Ollama/LM Studio 4-bit defaults), edge (AWQ on Jetson Orin), data center (SmoothQuant W8A8).";
}

// 11 · Frontier local
{
  const s = slideOf("Frontier models at home: closer, but not here yet", "Applications");
  {
    const x0 = M + 2.3, bw = 4.9, max = 1342, y0 = 1.95;
    txt(s, "GB of memory for the weights", { x: M, y: 1.5, w: 3.6, h: 0.35, fontSize: 14, color: C.text2 });
    const models = [["gpt-oss-120b", 234, 61], ["Llama 3.1 405B", 810, 203], ["DeepSeek-V3", 1342, 336]];
    // device reference lines
    [[128, "128 GB laptop"], [512, "512 GB Mac Studio"]].forEach(([gb, lab]) => {
      const xx = x0 + bw * gb / max;
      s.addShape(pres.shapes.LINE, { x: xx, y: y0 - 0.05, w: 0, h: 4.05, line: { color: HEX.muted, width: 1, dashType: "dash" } });
      txt(s, lab, { x: xx - 1, y: y0 + 4.05, w: 2, h: 0.3, fontSize: 11, color: C.text2, align: "center" });
    });
    models.forEach(([name, a, b], i) => {
      const y = y0 + i * 1.32;
      txt(s, name, { x: M, y: y + 0.02, w: 2.1, h: 0.4, fontSize: 15, bold: true, valign: "middle" });
      txt(s, [{ text: `${a.toLocaleString("en-US")} → ` , options: { color: HEX.muted } }, { text: `${b} GB`, options: { color: HEX.accent, bold: true } }],
        { x: M, y: y + 0.44, w: 2.1, h: 0.4, fontSize: 14, valign: "middle" });
      box(s, { x: x0, y: y + 0.08, w: bw * a / max, h: 0.36, fill: { color: HEX.rule } });
      box(s, { x: x0, y: y + 0.5, w: bw * b / max, h: 0.36, fill: { color: HEX.accent } });
    });
    box(s, { x: x0 + 1.6, y: 1.6, w: 0.25, h: 0.12, fill: { color: HEX.rule } });
    txt(s, "16-bit", { x: x0 + 1.95, y: 1.5, w: 0.8, h: 0.32, fontSize: 12, color: C.text2, valign: "middle" });
    box(s, { x: x0 + 2.8, y: 1.6, w: 0.25, h: 0.12, fill: { color: HEX.accent } });
    txt(s, "4-bit", { x: x0 + 3.15, y: 1.5, w: 0.8, h: 0.32, fontSize: 12, color: C.text2, valign: "middle" });
  }
  const sx = 8.75;
  const facts = [
    ["Already here", "OpenAI released gpt-oss-120b already in 4-bit (MXFP4, an open block format). Runs on one 80 GB GPU."],
    ["Within reach", "DeepSeek-V3 at 4-bit (~336 GB) fits in a 512 GB Mac Studio. Pricey and slow, but possible."],
    ["Not yet", "Laptops top out around 128 GB. Frontier-scale still needs ~3× more squeezing."],
  ];
  facts.forEach(([h, d], i) => {
    const y = 1.6 + i * 1.6;
    txt(s, h, { x: sx, y, w: 4, h: 0.4, fontSize: 17, bold: true, color: i === 0 ? C.accent1 : C.text1 });
    txt(s, d, { x: sx, y: y + 0.45, w: 3.95, h: 1.05, fontSize: 14, color: C.text2 });
  });
  cite(s, "OpenAI (2025); Meta AI (2024); DeepSeek-AI (2024). 4-bit ≈ params × 0.5 bytes (approx.).");
  notes[n] = "Same arithmetic as slide 2, divided by 4. Quantization turns 'impossible' into 'expensive'.";
}

// 12 · Image generators
{
  const s = slideOf("Why image generators quantize so easily", "Applications");
  s.addImage({ path: path.join(ASSETS, "pic_4bit.png"), x: M, y: 1.6, w: 4.2, h: 2.8, altText: "Landscape at 4-bit colour" });
  txt(s, "A pixel 3% off", { x: M, y: 4.55, w: 4.2, h: 0.4, fontSize: 18, bold: true });
  txt(s, "Nobody can tell. Still a great picture.", { x: M, y: 4.95, w: 4.2, h: 0.4, fontSize: 14, color: C.text2 });
  // code card
  const cx = 5.15;
  box(s, { x: cx, y: 1.6, w: 3.7, h: 2.8, fill: { color: HEX.ink } });
  txt(s, [
    { text: '{"city": "Paris"}', options: { color: HEX.white, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: '{"city": "Paris",', options: { color: "FF8A8D" } },
  ], { x: cx + 0.3, y: 1.9, w: 3.2, h: 2.2, fontSize: 18, fontFace: "Menlo", valign: "middle" });
  txt(s, "One token off", { x: cx, y: 4.55, w: 3.7, h: 0.4, fontSize: 18, bold: true });
  txt(s, "The program crashes. A text model gets no partial credit.", { x: cx, y: 4.95, w: 3.7, h: 0.7, fontSize: 14, color: C.text2 });
  // stat
  const sx = 9.4;
  txt(s, "3.5×", { x: sx, y: 1.55, w: 3.3, h: 0.9, fontSize: 48, bold: true, color: C.accent1 });
  txt(s, "less memory than 16-bit for FLUX.1 (a 12B image generator), with image quality close to the original", { x: sx, y: 2.45, w: 3.3, h: 0.8, fontSize: 14, color: C.text2 });
  txt(s, "16 GB", { x: sx, y: 3.45, w: 3.3, h: 0.9, fontSize: 48, bold: true });
  txt(s, "laptop GPU (RTX 4090) it runs on, 3.0× faster than 4-bit weight-only quantization", { x: sx, y: 4.35, w: 3.3, h: 0.8, fontSize: 14, color: C.text2 });
  txt(s, "Errors spread across millions of pixels average out. Errors in a single word don't.", { x: M, y: 6.0, w: 11.5, h: 0.45, fontSize: 17, italic: true });
  cite(s, "Li et al. (2025). Image: quantization_demo.ipynb Part 2.");
  notes[n] = "Images are forgiving, code/JSON is not. SVDQuant numbers: 3.5x memory, 3.0x speed vs W4A16, 16GB laptop 4090.";
}

// 13 · Strengths & limitations
{
  const s = slideOf("Strengths, limits, and risks", "Discussion");
  const cols = [
    ["Strengths", HEX.accent, ["Runs on your own device: private and offline", "About 4× less memory at 4-bit, and faster", "Less energy and hardware per answer"]],
    ["Limitations", HEX.ink, ["Easy tests miss small damage: ours stayed at 100% to 2-bit; long JSON fell to 75%", "Warning signs are subtle: 2-bit mistakes still looked like valid answers", "Rounding after training runs out of road below ~2 bits"]],
    ["Risks and ethics", HEX.alarm, ["Fails silently: confidently wrong, no error message", "Safety testing is done on the full model, not the shrunk copy", "Anyone can run strong models with no oversight: more access, harder to govern"]],
  ];
  const cw = 3.85, gap = 0.3;
  cols.forEach(([h, col, items], i) => {
    const x = M + i * (cw + gap);
    box(s, { x, y: 1.6, w: cw, h: 4.6, fill: { color: HEX.white } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.35, y: 1.95, w: 0.2, h: 0.2, fill: { color: col }, line: { type: "none" } });
    txt(s, h, { x: x + 0.7, y: 1.83, w: cw - 0.9, h: 0.45, fontSize: 20, bold: true, valign: "middle" });
    txt(s, items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1 } })),
      { x: x + 0.35, y: 2.55, w: cw - 0.6, h: 3.5, fontSize: 17, paraSpaceAfter: 16, color: C.text1 });
  });
  cite(s, "Limits: this demo (Part 5); Lin et al. (2024). Training for low bits: Apple (2025).");
  notes[n] = "Balance: huge practical win, but failure is silent and benchmarks can hide it.";
}

// 14 · Conclusion
{
  const s = slideOf("What to remember, and what's next", "Conclusion");
  const take = [
    ["Big models are limited by memory, not math.", "Fewer bits per weight = it fits, and it runs faster."],
    ["4-bit is the sweet spot.", "Ours: 69% smaller, 2.3× faster, no measurable loss."],
    ["Push too far and it slips, then falls off a cliff.", "Ours: long JSON 75% at 2-bit, everything 0% at 1.5-bit."],
  ];
  take.forEach(([h, d], i) => {
    const y = 1.65 + i * 1.3;
    txt(s, String(i + 1), { x: M, y, w: 0.6, h: 0.6, fontSize: 32, bold: true, color: C.accent1 });
    txt(s, h, { x: M + 0.7, y: y + 0.02, w: 6.2, h: 0.45, fontSize: 19, bold: true });
    txt(s, d, { x: M + 0.7, y: y + 0.5, w: 6.2, h: 0.45, fontSize: 15, color: C.text2 });
  });
  const fx = 7.9;
  box(s, { x: fx, y: 1.55, w: 4.8, h: 4.9, fill: { color: HEX.ink } });
  txt(s, "Next: train with the rounding built in (\u201Cquantization-aware training\u201D)", { x: fx + 0.4, y: 1.85, w: 4.0, h: 1.1, fontSize: 19, bold: true, color: C.background1 });
  const nxt = [
    ["Apple, 2025", "On-device model trained to work at 2 bits per weight"],
    ["BitNet b1.58", "Weights are only −1, 0 or +1 (3 options ≈ 1.58 bits)"],
    ["4-bit math in the chip", "NVIDIA's newest GPUs (Blackwell) run 4-bit blocks natively: NVFP4, and the open MXFP4"],
  ];
  nxt.forEach(([h, d], i) => {
    const y = 3.1 + i * 1.1;
    txt(s, h, { x: fx + 0.4, y, w: 4.0, h: 0.35, fontSize: 15, bold: true, color: "A9A9F8" });
    txt(s, d, { x: fx + 0.4, y: y + 0.38, w: 4.0, h: 0.6, fontSize: 14, color: C.background1 });
  });
  cite(s, "Apple (2025); Ma et al. (2024); Rouhani et al. (2023).");
  notes[n] = "Three takeaways, then future: quantization-aware training and native FP4 hardware.";
}

// 15 · References
{
  const s = slideOf("References", "References");
  const refs = fs.readFileSync(path.join(REPO, "tools/references.txt"), "utf8").trim().split("\n");
  const half = Math.ceil(refs.length / 2);
  [refs.slice(0, half), refs.slice(half)].forEach((col, i) => {
    txt(s, col.map((t, k) => ({ text: t, options: { breakLine: k < col.length - 1 } })),
      { x: M + i * 6.15, y: 1.4, w: 5.85, h: 5.4, fontSize: 10, color: C.text1, paraSpaceAfter: 5, lineSpacingMultiple: 1.0 });
  });
  notes[n] = "Full list also in README.md.";
}

// notes
pres.slides.forEach((s, i) => { if (notes[i + 1]) s.addNotes(notes[i + 1]); });

(async () => {
  await pres.writeFile({ fileName: OUTFILE });
  await applyTheme(OUTFILE, THEME);
  console.log("wrote", OUTFILE, "slides:", n);
})();
