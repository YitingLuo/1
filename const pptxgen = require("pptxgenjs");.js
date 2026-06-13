const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Colors
const C = {
  dark:    "1A1A2E",   // deep navy
  mid:     "16213E",   // navy
  accent:  "E94560",   // red accent
  accentB: "0F3460",   // blue accent
  teal:    "1B6CA8",   // teal
  gold:    "C9A84C",   // gold
  light:   "F5F5F5",   // off-white
  muted:   "A0A0B0",   // muted text
  white:   "FFFFFF",
  red:     "C0392B",
  green:   "1E8449",
  blue:    "1B6CA8",
  yellow:  "D4AC0D",
  gray:    "95A5A6",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "鄧如雯案：延伸討論與反思";

// ─────────────────────────────────────────────
// Helper: icon from react-icons → base64 PNG
// ─────────────────────────────────────────────
const { FaBalanceScale, FaBrain, FaFemale, FaExclamationTriangle, FaCheckCircle, FaSearch, FaGavel, FaHeart, FaShieldAlt, FaArrowRight } = require("react-icons/fa");
const { MdPsychology, MdOutlinePolicyRounded } = require("react-icons/md");

async function icon(Component, color = "FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Component, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ─────────────────────────────────────────────
// Slide builder helpers
// ─────────────────────────────────────────────
function darkBg(slide) {
  slide.background = { color: C.dark };
}
function lightBg(slide) {
  slide.background = { color: "F8F9FB" };
}
function addAccentBar(slide, color = C.accent) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color } });
}
function sectionTag(slide, text, x = 0.25, y = 0.18, color = C.accent) {
  slide.addText(text, { x, y, w: 2.5, h: 0.28, fontSize: 9, color: C.white, bold: true, align: "left",
    fill: { color }, rectRadius: 0.04 });
}
function slideTitle(slide, text, dark = true) {
  slide.addText(text, { x: 0.25, y: 0.55, w: 9.5, h: 0.65, fontSize: 28, bold: true,
    color: dark ? C.white : C.dark, fontFace: "Calibri", align: "left" });
}
function divLine(slide, y = 1.3, dark = true) {
  slide.addShape(pres.shapes.LINE, { x: 0.25, y, w: 9.5, h: 0,
    line: { color: dark ? "3A3A5C" : "DDDDEE", width: 0.8 } });
}

// ─────────────────────────────────────────────
// SLIDE 1 — Title slide
// ─────────────────────────────────────────────
async function slide1() {
  const s = pres.addSlide();
  darkBg(s);
  // Left accent stripe
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 3.5, h: 5.625, fill: { color: C.accentB } });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.2, y: 0, w: 0.35, h: 5.625, fill: { color: C.accent } });

  s.addText("1993", { x: 0.2, y: 0.5, w: 3.0, h: 0.8, fontSize: 60, bold: true, color: C.white, fontFace: "Georgia", align: "center", transparency: 20 });
  s.addText("鄧如雯\n殺夫案", { x: 0.15, y: 1.3, w: 3.2, h: 1.6, fontSize: 30, bold: true, color: C.white, fontFace: "Calibri", align: "center" });
  s.addText("台灣台北縣板橋市", { x: 0.2, y: 2.9, w: 3.0, h: 0.35, fontSize: 11, color: "AACCFF", align: "center" });

  s.addText("延伸討論與反思", { x: 3.7, y: 1.5, w: 6.1, h: 0.7, fontSize: 32, bold: true, color: C.white, fontFace: "Calibri" });
  s.addText("從犯罪心理學視角分析審判結果、\nBWS 理論與制度性困境", {
    x: 3.7, y: 2.35, w: 6.1, h: 0.9, fontSize: 14, color: "BBCCDD", fontFace: "Calibri", lineSpacingMultiple: 1.5 });

  s.addShape(pres.shapes.LINE, { x: 3.7, y: 2.2, w: 5.5, h: 0, line: { color: C.accent, width: 1.5 } });

  s.addText("· 銜接前段：犯罪風險因子與保護因子", { x: 3.7, y: 3.5, w: 6.1, h: 0.32, fontSize: 11, color: C.muted });
  s.addText("· 後段將延伸：家暴法的改善方向", { x: 3.7, y: 3.85, w: 6.1, h: 0.32, fontSize: 11, color: C.muted });

  s.addText("犯罪心理學期末報告", { x: 3.7, y: 4.9, w: 6.1, h: 0.3, fontSize: 10, color: "556677", align: "left" });
}

// ─────────────────────────────────────────────
// SLIDE 2 — 三審對照總表
// ─────────────────────────────────────────────
async function slide2() {
  const s = pres.addSlide();
  lightBg(s);
  addAccentBar(s, C.accentB);
  sectionTag(s, "三審判決比較", 0.25, 0.18, C.accentB);
  slideTitle(s, "三審判決：從定罪到定讞", false);
  divLine(s, 1.3, false);

  // Table
  const headerFill = { color: C.accentB };
  const rowY = { fill: { color: "FFF8E1" } };
  const rowG = { fill: { color: "E8F5E9" } };
  const rowB = { fill: { color: "E3F2FD" } };

  const th = (t) => ({ text: t, options: { bold: true, color: C.white, fill: headerFill, align: "center", valign: "middle", fontSize: 11 } });
  const tc = (t, fill, color = "1A1A2E") => ({ text: t, options: { color, fill, align: "center", valign: "middle", fontSize: 11 } });

  const rows = [
    [th("審級"), th("法院"), th("正當防衛"), th("義憤殺人"), th("精神耗弱"), th("判決")],
    [tc("一審", rowY), tc("板橋地方法院", rowY), tc("✗ 不成立", rowY, C.red), tc("✗ 不成立", rowY, C.red), tc("✗ 駁回鑑定", rowY, C.red), tc("5年6個月", rowY, C.red)],
    [tc("二審", rowG), tc("臺灣高等法院", rowG), tc("✗ 不成立", rowG, C.red), tc("✗ 不成立", rowG, C.red), tc("✓ 採信鑑定", rowG, C.green), tc("改判 3年", rowG, C.green)],
    [tc("三審", rowB), tc("最高法院", rowB), tc("—", rowB, C.gray), tc("—", rowB, C.gray), tc("維持二審", rowB, C.blue), tc("定讞 3年", rowB, C.blue)],
  ];

  s.addTable(rows, {
    x: 0.25, y: 1.4, w: 9.5, h: 2.8,
    colW: [0.8, 2.1, 1.7, 1.7, 1.7, 1.5],
    border: { pt: 0.5, color: "CCCCDD" },
    autoPage: false,
  });

  // Key insight box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 4.4, w: 9.5, h: 0.9,
    fill: { color: "EEF2FF" }, line: { color: "AABBEE", width: 0.5 } });
  s.addText("核心觀察：三審判決均未從「正當防衛」角度理解鄧如雯的行為，只以「精神耗弱」作為減刑依據——給的是寬恕，不是正名。", {
    x: 0.4, y: 4.48, w: 9.2, h: 0.72, fontSize: 12, color: C.accentB, italic: true, valign: "middle" });
}

// ─────────────────────────────────────────────
// SLIDE 3 — 一審深度分析
// ─────────────────────────────────────────────
async function slide3() {
  const s = pres.addSlide();
  darkBg(s);
  addAccentBar(s, C.yellow);
  sectionTag(s, "一審分析", 0.25, 0.18, C.yellow);
  slideTitle(s, "一審：法律邏輯與心理學盲點");
  divLine(s);

  // 3 cards
  const cards = [
    { title: "正當防衛 ✗", sub: "防衛必須即時存在\n熟睡中無現存威脅", color: C.red },
    { title: "義憤殺人 ✗", sub: "須「當場」行兇\n事隔2小時不符要件", color: C.red },
    { title: "精神耗弱 ✗", sub: "以「事後清醒\n能自首」駁回鑑定", color: C.red },
  ];
  for (let i = 0; i < cards.length; i++) {
    const x = 0.25 + i * 3.25;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.45, w: 3.0, h: 1.5, fill: { color: "242446" }, line: { color: cards[i].color, width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.45, w: 3.0, h: 0.35, fill: { color: cards[i].color } });
    s.addText(cards[i].title, { x, y: 1.45, w: 3.0, h: 0.35, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText(cards[i].sub, { x: x + 0.1, y: 1.85, w: 2.8, h: 1.0, fontSize: 11, color: "BBBBCC", valign: "top", lineSpacingMultiple: 1.4 });
  }

  // Verdict
  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 3.1, w: 9.5, h: 0.45, fill: { color: C.red } });
  s.addText("判決：有期徒刑 5年6個月（殺人罪，依自首減刑）　民事：子女監護權歸林榮祥", {
    x: 0.35, y: 3.1, w: 9.3, h: 0.45, fontSize: 12, bold: true, color: C.white, valign: "middle" });

  // Psychological critique
  s.addText("犯罪心理學評析", { x: 0.25, y: 3.7, w: 3.0, h: 0.3, fontSize: 11, bold: true, color: C.yellow });
  s.addText("「事後能清楚陳述」≠「行兇當時精神正常」", {
    x: 0.25, y: 4.05, w: 9.5, h: 0.3, fontSize: 12, bold: true, color: "FFDD88" });
  s.addText("PTSD 患者在高度應激後，可在解離狀態下完成行為，事後仍能陳述——這正是創傷反應的症狀，而非正常的標誌。一審以「自首」否定精神耗弱，犯了將事後理性等同當時正常的根本邏輯謬誤。", {
    x: 0.25, y: 4.4, w: 9.5, h: 0.9, fontSize: 11, color: "AAAACC", lineSpacingMultiple: 1.4 });
}

// ─────────────────────────────────────────────
// SLIDE 4 — 二審三審
// ─────────────────────────────────────────────
async function slide4() {
  const s = pres.addSlide();
  lightBg(s);
  addAccentBar(s, C.green);
  sectionTag(s, "二審・三審分析", 0.25, 0.18, C.green);
  slideTitle(s, "二審翻轉、三審定讞：進步與侷限", false);
  divLine(s, 1.3, false);

  // Two columns
  // Left: 二審
  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 1.45, w: 4.55, h: 2.7, fill: { color: "FFFFFF" }, line: { color: "CCEECC", width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 1.45, w: 4.55, h: 0.4, fill: { color: C.green } });
  s.addText("二審｜臺灣高等法院 → 改判 3年", { x: 0.3, y: 1.45, w: 4.45, h: 0.4, fontSize: 11, bold: true, color: C.white, valign: "middle" });
  s.addText([
    { text: "關鍵翻轉：", options: { bold: true, breakLine: true } },
    { text: "採信三軍總醫院精神鑑定\n「案發時精神極度耗弱」", options: { breakLine: true } },
    { text: "\n意涵：", options: { bold: true, breakLine: true } },
    { text: "二審承認案發當下與事後\n是兩種不同的心理狀態", options: { breakLine: true } },
    { text: "\n✓ 自首 + 精神耗弱 雙重減刑", options: { color: C.green } },
  ], { x: 0.35, y: 1.92, w: 4.35, h: 2.1, fontSize: 11, color: "333344", lineSpacingMultiple: 1.4 });

  // Right: 三審
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.45, w: 4.55, h: 2.7, fill: { color: "FFFFFF" }, line: { color: "BBCCEE", width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.45, w: 4.55, h: 0.4, fill: { color: C.accentB } });
  s.addText("三審｜最高法院 → 定讞 3年", { x: 5.25, y: 1.45, w: 4.45, h: 0.4, fontSize: 11, bold: true, color: C.white, valign: "middle" });
  s.addText([
    { text: "1995年3月23日定讞", options: { bold: true, breakLine: true } },
    { text: "實際服刑約一年半後假釋\n出獄，改名換姓重新生活", options: { breakLine: true } },
    { text: "\n定讞的雙重意義：", options: { bold: true, breakLine: true } },
    { text: "①  司法對受虐婦女仍定性為\n    「情有可原的殺人」", options: { breakLine: true } },
    { text: "②  「不滿意」的結果反而成為\n    《家暴法》立法的最大推力", options: {} },
  ], { x: 5.3, y: 1.92, w: 4.35, h: 2.1, fontSize: 11, color: "333344", lineSpacingMultiple: 1.4 });

  // Bottom insight
  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 4.35, w: 9.5, h: 0.95, fill: { color: "1A1A2E" } });
  s.addText("核心反思：二審給出了「量刑減輕」，而非「行為正當化」——這兩者有本質差異。前者是對罪行的寬恕，後者才是對受虐婦女處境的真正理解與法律正名。", {
    x: 0.4, y: 4.38, w: 9.2, h: 0.88, fontSize: 11, color: "AACCFF", italic: true, valign: "middle", lineSpacingMultiple: 1.35 });
}

// ─────────────────────────────────────────────
// SLIDE 5 — BWS 理論
// ─────────────────────────────────────────────
async function slide5() {
  const s = pres.addSlide();
  darkBg(s);
  addAccentBar(s, C.accent);
  sectionTag(s, "延伸討論 1／3", 0.25, 0.18, C.accent);
  slideTitle(s, "受虐婦女症候群（BWS）：為何她無法離開？");
  divLine(s);

  // 4 characteristics
  const chars = [
    { n: "01", title: "認為暴力是自己的錯", desc: "長期被教化「如果妳更好，他就不會打妳」" },
    { n: "02", title: "無法將責任歸咎他人", desc: "習得性無助使受虐者內化施暴者的邏輯" },
    { n: "03", title: "恐懼自己或孩子的生命", desc: "林阿棋曾將孩子倒插進洗衣機，威脅摔死" },
    { n: "04", title: "認為施暴者無所不在", desc: "多次逃家均被找回，求助警察也無效" },
  ];

  for (let i = 0; i < 4; i++) {
    const x = 0.25 + (i % 2) * 4.8;
    const y = 1.55 + Math.floor(i / 2) * 1.5;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 1.3, fill: { color: "1E1E3A" }, line: { color: C.accent, width: 0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.5, h: 1.3, fill: { color: C.accent } });
    s.addText(chars[i].n, { x, y, w: 0.5, h: 1.3, fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle" });
    s.addText(chars[i].title, { x: x + 0.6, y: y + 0.15, w: 3.7, h: 0.4, fontSize: 12, bold: true, color: "FFCCCC" });
    s.addText(chars[i].desc, { x: x + 0.6, y: y + 0.55, w: 3.7, h: 0.65, fontSize: 10.5, color: "AAAACC", lineSpacingMultiple: 1.35 });
  }

  // Walker citation
  s.addText("Walker (1984) · 習得性無助理論 (Learned Helplessness) · Lagunathan (2021) BJPsych Open", {
    x: 0.25, y: 5.2, w: 9.5, h: 0.25, fontSize: 9, color: "556677", italic: true });
}

// ─────────────────────────────────────────────
// SLIDE 6 — 正當防衛的性別盲點
// ─────────────────────────────────────────────
async function slide6() {
  const s = pres.addSlide();
  lightBg(s);
  addAccentBar(s, C.accentB);
  sectionTag(s, "延伸討論 2／3", 0.25, 0.18, C.accentB);
  slideTitle(s, "正當防衛框架的性別盲點", false);
  divLine(s, 1.3, false);

  // Two columns: Law says vs Reality
  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 1.45, w: 4.4, h: 3.5, fill: { color: "1A1A2E" } });
  s.addText("法律預設的框架", { x: 0.3, y: 1.5, w: 4.3, h: 0.38, fontSize: 13, bold: true, color: C.yellow, align: "center" });
  s.addText([
    { text: "「即時」威脅", options: { bold: true, color: "FFEEAA", breakLine: true } },
    { text: "威脅必須在當下發生中\n", options: { color: "AAAACC", breakLine: true } },
    { text: "「面對面」衝突", options: { bold: true, color: "FFEEAA", breakLine: true } },
    { text: "雙方正在對峙、肢體接觸中\n", options: { color: "AAAACC", breakLine: true } },
    { text: "「合理男性」標準", options: { bold: true, color: "FFEEAA", breakLine: true } },
    { text: "以男性對男性衝突為模型\n建構的合理反應基準", options: { color: "AAAACC" } },
  ], { x: 0.35, y: 1.93, w: 4.1, h: 2.85, fontSize: 11, lineSpacingMultiple: 1.4 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.35, y: 1.45, w: 4.4, h: 3.5, fill: { color: "FFFFFF" }, line: { color: "DDDDEE", width: 0.5 } });
  s.addText("受虐婦女的現實", { x: 5.4, y: 1.5, w: 4.3, h: 0.38, fontSize: 13, bold: true, color: C.accent, align: "center" });
  s.addText([
    { text: "「慢性」恐懼累積", options: { bold: true, color: C.accentB, breakLine: true } },
    { text: "威脅是長期存在的、下次\n暴力隨時可能發生\n", options: { color: "444455", breakLine: true } },
    { text: "「預防性」自保行為", options: { bold: true, color: C.accentB, breakLine: true } },
    { text: "趁熟睡時行動是唯一安全\n的自衛時機，非冷靜謀殺\n", options: { color: "444455", breakLine: true } },
    { text: "體力懸殊的現實", options: { bold: true, color: C.accentB, breakLine: true } },
    { text: "要求正面對抗等於要求她\n用最不利條件應戰", options: { color: "444455" } },
  ], { x: 5.45, y: 1.93, w: 4.1, h: 2.85, fontSize: 11, lineSpacingMultiple: 1.4 });

  // VS divider
  s.addShape(pres.shapes.OVAL, { x: 4.7, y: 2.6, w: 0.6, h: 0.6, fill: { color: C.accent } });
  s.addText("VS", { x: 4.7, y: 2.6, w: 0.6, h: 0.6, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });

  // Quote
  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 5.1, w: 9.5, h: 0.35, fill: { color: "EEEEFF" } });
  s.addText("「法律問的是合理的男人會怎麼做，而不是長期受虐的女性會怎麼做。」— Forell & Matthews (2000)", {
    x: 0.35, y: 5.12, w: 9.3, h: 0.3, fontSize: 10, italic: true, color: C.accentB, valign: "middle" });
}

// ─────────────────────────────────────────────
// SLIDE 7 — 制度性缺失
// ─────────────────────────────────────────────
async function slide7() {
  const s = pres.addSlide();
  darkBg(s);
  addAccentBar(s, C.gold);
  sectionTag(s, "延伸討論 3／3", 0.25, 0.18, C.gold);
  slideTitle(s, "制度性缺失：法律如何困住她？");
  divLine(s);

  // 3 blocks
  const blocks = [
    {
      icon: "⚖️", title: "無《家庭暴力防治法》",
      desc: "1993年台灣無家暴法。受虐婦女\n若無法舉證而離開，反被認定未\n履行同居義務，訴訟中居劣勢。",
      color: C.red
    },
    {
      icon: "👨‍⚖️", title: "子女監護權歸夫家",
      desc: "當時民法預設離婚後子女歸父方。\n鄧如雯要離婚就等於失去孩子。\n孩子是林阿棋控制她的籌碼。",
      color: C.yellow
    },
    {
      icon: "🚔", title: "警察系統無力保護",
      desc: "曾多次報案，無實質改善。丈夫\n更以失蹤人口名義通報，讓警察\n成為「把她送回去的工具」。",
      color: C.teal
    },
  ];

  for (let i = 0; i < 3; i++) {
    const x = 0.25 + i * 3.25;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.45, w: 3.0, h: 2.7, fill: { color: "1E1E3A" }, line: { color: blocks[i].color, width: 1 } });
    s.addText(blocks[i].icon, { x, y: 1.5, w: 3.0, h: 0.5, fontSize: 22, align: "center" });
    s.addText(blocks[i].title, { x: x + 0.1, y: 2.05, w: 2.8, h: 0.42, fontSize: 12, bold: true, color: blocks[i].color, align: "center" });
    s.addText(blocks[i].desc, { x: x + 0.1, y: 2.52, w: 2.8, h: 1.45, fontSize: 10.5, color: "AAAACC", lineSpacingMultiple: 1.4 });
  }

  // Key statement
  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 4.3, w: 9.5, h: 1.0, fill: { color: "14142E" }, line: { color: C.gold, width: 1 } });
  s.addText("制度不只是「未能保護她」，更是主動封鎖了她所有的合法出路。", {
    x: 0.4, y: 4.37, w: 9.2, h: 0.42, fontSize: 14, bold: true, color: C.gold, valign: "middle" });
  s.addText("鄧如雯的行為是個人悲劇，更是制度性失靈的必然結果。", {
    x: 0.4, y: 4.8, w: 9.2, h: 0.38, fontSize: 11, color: "AAAACC", valign: "middle" });
}

// ─────────────────────────────────────────────
// SLIDE 8 — RNR 模式反思
// ─────────────────────────────────────────────
async function slide8() {
  const s = pres.addSlide();
  lightBg(s);
  addAccentBar(s, C.teal);
  sectionTag(s, "反思", 0.25, 0.18, C.teal);
  slideTitle(s, "RNR 模式：這個框架適用嗎？", false);
  divLine(s, 1.3, false);

  // RNR table
  const thStyle = { bold: true, color: C.white, fill: { color: C.teal }, align: "center", valign: "middle", fontSize: 11 };
  const tdStyle = (c = "1A1A2E") => ({ color: c, align: "left", valign: "middle", fontSize: 11 });
  const rows = [
    [
      { text: "RNR 面向", options: thStyle },
      { text: "標準定義", options: thStyle },
      { text: "套用於鄧如雯案", options: thStyle },
      { text: "侷限", options: thStyle },
    ],
    [
      { text: "風險原則", options: { ...tdStyle(), fill: { color: "F0FAF0" }, bold: true } },
      { text: "高風險→密集介入\n低風險→輕度處遇", options: { ...tdStyle(), fill: { color: "F0FAF0" } } },
      { text: "無前科、情境型犯罪\n→ 再犯風險低", options: { ...tdStyle(C.green), fill: { color: "F0FAF0" } } },
      { text: "標準工具未考量\n受虐情境", options: { ...tdStyle(C.red), fill: { color: "F0FAF0" } } },
    ],
    [
      { text: "需求原則", options: { ...tdStyle(), fill: { color: "FAFFF0" }, bold: true } },
      { text: "針對「犯罪需求」\n（反社會因子）介入", options: { ...tdStyle(), fill: { color: "FAFFF0" } } },
      { text: "核心需求是創傷修復\n而非矯正反社會傾向", options: { ...tdStyle(C.blue), fill: { color: "FAFFF0" } } },
      { text: "BWS/PTSD 不在\n標準犯罪需求清單", options: { ...tdStyle(C.red), fill: { color: "FAFFF0" } } },
    ],
    [
      { text: "回應原則", options: { ...tdStyle(), fill: { color: "F0F0FF" }, bold: true } },
      { text: "以 CBT 為主\n依個案特質調整", options: { ...tdStyle(), fill: { color: "F0F0FF" } } },
      { text: "應採創傷知情照護\n（Trauma-Informed Care）", options: { ...tdStyle(C.blue), fill: { color: "F0F0FF" } } },
      { text: "一般 CBT 對長期\n創傷者效果有限", options: { ...tdStyle(C.red), fill: { color: "F0F0FF" } } },
    ],
  ];
  s.addTable(rows, {
    x: 0.25, y: 1.45, w: 9.5, h: 3.0,
    colW: [1.5, 2.3, 2.8, 2.9],
    border: { pt: 0.5, color: "CCDDCC" },
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 4.6, w: 9.5, h: 0.72, fill: { color: "EFF8FF" }, line: { color: "AACCEE", width: 0.5 } });
  s.addText("RNR 模式是「主動型犯罪者」的框架——鄧如雯屬「受害者反擊型犯罪」，套用此模型容易錯置處遇目標，進一步傷害當事人。", {
    x: 0.4, y: 4.64, w: 9.2, h: 0.62, fontSize: 11, color: C.accentB, italic: true, valign: "middle", lineSpacingMultiple: 1.35 });
}

// ─────────────────────────────────────────────
// SLIDE 9 — 銜接：前後段的連結
// ─────────────────────────────────────────────
async function slide9() {
  const s = pres.addSlide();
  darkBg(s);
  addAccentBar(s, C.accent);
  sectionTag(s, "前後銜接", 0.25, 0.18, C.accent);
  slideTitle(s, "本案在整體報告中的定位");
  divLine(s);

  // Arrow flow
  const boxes = [
    { label: "前段", sub: "犯罪風險因子\n加害者與受害者\n保護因子分析", color: C.accentB },
    { label: "本段", sub: "BWS 理論\n三審判決分析\n制度性缺失反思", color: C.accent },
    { label: "後段", sub: "改善方法\n家暴法立法歷程\n現代制度建議", color: C.green },
  ];

  for (let i = 0; i < 3; i++) {
    const x = 0.5 + i * 3.2;
    const isCenter = i === 1;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.6, w: 2.8, h: 2.0,
      fill: { color: isCenter ? boxes[i].color : "1E1E3A" },
      line: { color: boxes[i].color, width: isCenter ? 0 : 1.5 } });
    s.addText(boxes[i].label, { x, y: 1.6, w: 2.8, h: 0.45, fontSize: 14, bold: true,
      color: isCenter ? C.white : boxes[i].color, align: "center", valign: "middle" });
    s.addShape(pres.shapes.LINE, { x: x + 0.1, y: 2.08, w: 2.6, h: 0,
      line: { color: isCenter ? "FFFFFF" : boxes[i].color, width: 0.5, transparency: 40 } });
    s.addText(boxes[i].sub, { x: x + 0.1, y: 2.15, w: 2.6, h: 1.35, fontSize: 11,
      color: isCenter ? C.white : "AAAACC", align: "center", valign: "top", lineSpacingMultiple: 1.5 });
    if (i < 2) {
      s.addShape(pres.shapes.LINE, { x: x + 2.85, y: 2.6, w: 0.3, h: 0, line: { color: "555577", width: 1.5 } });
      s.addText("▶", { x: x + 2.82, y: 2.44, w: 0.35, h: 0.35, fontSize: 13, color: "555577", align: "center" });
    }
  }

  // Bridging text
  s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 3.85, w: 9.5, h: 1.5, fill: { color: "12122E" } });
  s.addText("本段的核心貢獻", { x: 0.4, y: 3.9, w: 9.2, h: 0.35, fontSize: 12, bold: true, color: C.gold });
  s.addText([
    { text: "• 用 BWS 理論解釋「風險因子」如何在心理層面產生效果，補充前段的犯罪學視角\n", options: { breakLine: false } },
    { text: "• 透過三審分析揭示制度性缺失，為後段「為什麼需要家暴法」提供具體的問題背景\n", options: { breakLine: false } },
    { text: "• RNR 框架的反思連結「處遇政策」，直接鋪陳後段改善方向的必要性", options: {} },
  ], { x: 0.4, y: 4.3, w: 9.2, h: 1.0, fontSize: 11, color: "AAAACC", lineSpacingMultiple: 1.45 });
}

// ─────────────────────────────────────────────
// SLIDE 10 — 結語
// ─────────────────────────────────────────────
async function slide10() {
  const s = pres.addSlide();
  darkBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark } });
  // Accent gradient left side
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.accent } });

  s.addText("反思", { x: 0.25, y: 0.35, w: 2, h: 0.35, fontSize: 10, color: C.muted, bold: true });
  s.addText("鄧如雯案教會我們什麼？", {
    x: 0.25, y: 0.72, w: 9.5, h: 0.65, fontSize: 28, bold: true, color: C.white, fontFace: "Calibri" });
  s.addShape(pres.shapes.LINE, { x: 0.25, y: 1.45, w: 9.5, h: 0, line: { color: "3A3A5C", width: 0.8 } });

  // 3 takeaways
  const pts = [
    { n: "01", head: "司法寬恕 ≠ 制度理解", body: "三審給的是「減刑」，不是對受虐婦女處境的法律正名。量刑的彈性是個人的，制度的改變才是系統性的。" },
    { n: "02", head: "犯罪不只是個人偏差", body: "BWS 與制度性缺失告訴我們：當所有合法出路都被封鎖，極端行為往往是結構性失靈的必然結果。" },
    { n: "03", head: "問題仍未解決", body: "2023 年新北季姓婦人案顯示，《家暴法》上路 25 年後，「長期受暴、無法掙脫」的困境仍在。制度的回應仍有盲點。" },
  ];

  for (let i = 0; i < 3; i++) {
    const y = 1.6 + i * 1.15;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y, w: 0.55, h: 0.95, fill: { color: C.accent } });
    s.addText(pts[i].n, { x: 0.25, y, w: 0.55, h: 0.95, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
    s.addText(pts[i].head, { x: 0.92, y: y + 0.05, w: 9.0, h: 0.36, fontSize: 13, bold: true, color: "FFDDDD" });
    s.addText(pts[i].body, { x: 0.92, y: y + 0.43, w: 9.0, h: 0.45, fontSize: 10.5, color: "AAAACC", lineSpacingMultiple: 1.3 });
  }

  s.addShape(pres.shapes.LINE, { x: 0.25, y: 5.18, w: 9.5, h: 0, line: { color: "3A3A5C", width: 0.5 } });
  s.addText("資料來源：國家文化記憶庫・臺灣法實證資料庫・Walker(1984)・Lagunathan(2021)・報導者(2023)", {
    x: 0.25, y: 5.22, w: 9.5, h: 0.25, fontSize: 8.5, color: "445566", italic: true });
}

// ─────────────────────────────────────────────
// Build all slides
// ─────────────────────────────────────────────
(async () => {
  await slide1();
  await slide2();
  await slide3();
  await slide4();
  await slide5();
  await slide6();
  await slide7();
  await slide8();
  await slide9();
  await slide10();

  await pres.writeFile({ fileName: "/mnt/user-data/outputs/鄧如雯案_延伸討論與反思.pptx" });
  console.log("Done");
})();