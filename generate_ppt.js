/**
 * generate_ppt.js — 2-Page Per Project Edition
 * 用法: node generate_ppt.js <data.json> <output.pptx>
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");

const dataPath   = process.argv[2];
const outputPath = process.argv[3];
const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));

// ── 色盤 ────────────────────────────────────────────────
const C = {
  navy:      "1E3A5F",
  blue:      "2E6DB4",
  lightBlue: "D6E4F0",
  paleBlue:  "EEF4FB",
  bg:        "F4F7FB",
  white:     "FFFFFF",
  text:      "1C2B3A",
  muted:     "64748B",
  divider:   "E2E8F0",
  accent:    "F0A500",
  red:       "DC2626",
  redSoft:   "FEE2E2",
  redDeep:   "991B1B",
  amber:     "F59E0B",
  amberSoft: "FEF3C7",
  amberDeep: "92400E",
  green:     "16A34A",
  greenSoft: "D1FAE5",
  greenDeep: "065F46",
};

const HEALTH_COLOR = { RED: C.red, AMBER: C.amber, GREEN: C.green };
const HEALTH_SOFT  = { RED: C.redSoft, AMBER: C.amberSoft, GREEN: C.greenSoft };
const HEALTH_LABEL = { RED: "警報", AMBER: "注意", GREEN: "健康" };

const STATUS_COLOR = {
  "完成":  C.green,
  "進行中": C.blue,
  "延遲":  C.red,
  "風險":  C.red,
  "待開始": C.muted,
};

// ── 共用工具 ────────────────────────────────────────────
function makeShadow() {
  return { type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.07 };
}

function slideHeader(slide, title, subtitle, pageTag) {
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.58, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText(title, {
    x: 0.35, y: 0, w: 7.2, h: 0.58,
    fontSize: 17, bold: true, color: C.white, fontFace: "Calibri", valign: "middle", margin: 0
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0, y: 0, w: 9.65, h: 0.58,
      fontSize: 9.5, color: "CADCFC", fontFace: "Calibri", align: "right", valign: "middle", margin: 0
    });
  }
  if (pageTag) {
    slide.addText(pageTag, {
      x: 0, y: 5.38, w: 9.85, h: 0.18,
      fontSize: 8, color: C.muted, fontFace: "Calibri", align: "right", valign: "middle"
    });
  }
  slide.addShape("rect", { x: 0, y: 0.58, w: 10, h: 0.03, fill: { color: C.accent }, line: { color: C.accent } });
}

function progressBar(slide, x, y, w, h, pct, fillColor) {
  slide.addShape("roundRect", { x, y, w, h, fill: { color: C.divider }, line: { color: C.divider }, rectRadius: 0.02 });
  const filled = Math.max(0, Math.min(100, pct)) / 100 * w;
  if (filled > 0.02) {
    slide.addShape("roundRect", { x, y, w: filled, h, fill: { color: fillColor }, line: { color: fillColor }, rectRadius: 0.02 });
  }
}

function card(slide, x, y, w, h, opts) {
  opts = opts || {};
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: opts.fill || C.white },
    line: { color: opts.border || C.divider, pt: opts.borderPt || 0.5 },
    shadow: opts.shadow ? makeShadow() : undefined
  });
  if (opts.accentLeft) {
    slide.addShape("rect", { x, y, w: 0.08, h, fill: { color: opts.accentLeft }, line: { color: opts.accentLeft } });
  }
}

// ══════════════════════════════════════════════════════
// PPT 開始
// ══════════════════════════════════════════════════════
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "專案管理報告";

// ══════════════════════════════════════════════════════
// Slide 1: 封面
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape("rect", { x: 7.3, y: 0, w: 2.7, h: 5.625, fill: { color: C.blue }, line: { color: C.blue } });
  s.addShape("rect", { x: 8.85, y: 0, w: 1.15, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });

  s.addText("專案管理報告", {
    x: 0.6, y: 1.3, w: 6.5, h: 0.95,
    fontSize: 40, bold: true, color: C.white, fontFace: "Calibri", charSpacing: 4
  });
  s.addText("Executive Status Report", {
    x: 0.6, y: 2.28, w: 6.5, h: 0.4,
    fontSize: 13, color: "CADCFC", fontFace: "Calibri", italic: true
  });
  s.addShape("rect", { x: 0.6, y: 2.82, w: 3.6, h: 0.04, fill: { color: C.accent }, line: { color: C.accent } });

  const hlColor = HEALTH_COLOR[data.overallHealth] || C.blue;
  s.addShape("roundRect", {
    x: 0.6, y: 3.05, w: 6.5, h: 0.72,
    fill: { color: hlColor }, line: { color: hlColor }, rectRadius: 0.08
  });
  s.addText(data.headline || "專案狀態報告", {
    x: 0.6, y: 3.05, w: 6.5, h: 0.72,
    fontSize: 17, bold: true, color: C.white, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });
  s.addText("報告日期　" + data.reportDate, {
    x: 0.6, y: 3.98, w: 6.5, h: 0.38, fontSize: 13, color: "CADCFC", fontFace: "Calibri"
  });
  s.addText("下次回報　" + data.nextReport, {
    x: 0.6, y: 4.36, w: 6.5, h: 0.38, fontSize: 13, color: "CADCFC", fontFace: "Calibri"
  });
}

// ══════════════════════════════════════════════════════
// Slide 2: 執行摘要
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "執行摘要", "一頁掌握所有專案 ｜ " + data.reportDate);

  var summary = [
    { label: "需立即關注", count: data.stats.redProjects,   color: C.red,   soft: C.redSoft,   level: "RED"   },
    { label: "需要觀察",   count: data.stats.amberProjects, color: C.amber, soft: C.amberSoft, level: "AMBER" },
    { label: "進度健康",   count: data.stats.greenProjects, color: C.green, soft: C.greenSoft, level: "GREEN" },
  ];
  summary.forEach(function(item, i) {
    var x = 0.35 + i * 3.12;
    card(s, x, 0.78, 2.95, 0.88, { shadow: true, border: item.soft, borderPt: 1, accentLeft: item.color });
    s.addText(String(item.count), {
      x: x + 0.22, y: 0.81, w: 0.88, h: 0.82,
      fontSize: 40, bold: true, color: item.color, fontFace: "Calibri", valign: "middle"
    });
    s.addText(item.label, {
      x: x + 1.18, y: 0.83, w: 1.65, h: 0.38,
      fontSize: 11, color: C.muted, fontFace: "Calibri", valign: "middle"
    });
    s.addText("個專案", {
      x: x + 1.18, y: 1.18, w: 1.65, h: 0.38,
      fontSize: 14, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
    });
  });

  var TX = 0.3, TY = 1.88, TW = 9.4;
  var COLS = [0.48, 1.72, 0.82, 1.52, 1.22, 0.92, 2.72];
  var HDRS = ["", "專案", "完成", "進度", "預估完工", "偏差", "主要關鍵"];
  var cx = TX;
  HDRS.forEach(function(h, i) {
    s.addShape("rect", { x: cx, y: TY, w: COLS[i], h: 0.3, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(h, {
      x: cx, y: TY, w: COLS[i], h: 0.3,
      fontSize: 9, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    cx += COLS[i];
  });

  var ROW_H = 0.54;
  data.projects.forEach(function(p, i) {
    var y = TY + 0.3 + i * ROW_H;
    var rowBg = i % 2 === 0 ? C.white : C.paleBlue;
    s.addShape("rect", { x: TX, y: y, w: TW, h: ROW_H, fill: { color: rowBg }, line: { color: C.divider, pt: 0.5 } });

    var rcx = TX;
    var h = p.health;
    var hColor = HEALTH_COLOR[h.level];

    s.addShape("ellipse", {
      x: rcx + (COLS[0] - 0.3) / 2, y: y + (ROW_H - 0.3) / 2, w: 0.3, h: 0.3,
      fill: { color: hColor }, line: { color: hColor }
    });
    rcx += COLS[0];

    s.addText(p.name, {
      x: rcx + 0.08, y: y + 0.05, w: COLS[1] - 0.1, h: 0.27,
      fontSize: 11, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
    });
    s.addText(p.owner + " · " + HEALTH_LABEL[h.level], {
      x: rcx + 0.08, y: y + 0.3, w: COLS[1] - 0.1, h: 0.2,
      fontSize: 8.5, color: hColor, fontFace: "Calibri"
    });
    rcx += COLS[1];

    s.addText(h.donePct + "%", {
      x: rcx, y: y, w: COLS[2], h: ROW_H,
      fontSize: 17, bold: true, color: C.text, fontFace: "Calibri", align: "center", valign: "middle"
    });
    rcx += COLS[2];

    var bX = rcx + 0.1, bW = COLS[3] - 0.2, bY = y + ROW_H / 2 - 0.07;
    progressBar(s, bX, bY, bW, 0.14, h.donePct, hColor);
    var tickX = bX + bW * Math.min(100, h.elapsedPct) / 100;
    s.addShape("rect", { x: tickX - 0.01, y: bY - 0.04, w: 0.02, h: 0.22, fill: { color: C.text }, line: { color: C.text } });
    s.addText("已過 " + h.elapsedPct + "%", {
      x: bX, y: y + ROW_H - 0.17, w: bW, h: 0.15,
      fontSize: 7, color: C.muted, fontFace: "Calibri", align: "center"
    });
    rcx += COLS[3];

    s.addText(h.forecastEnd || "—", {
      x: rcx, y: y + 0.05, w: COLS[4], h: 0.27,
      fontSize: 10, bold: true, color: C.text, fontFace: "Calibri", align: "center", valign: "middle"
    });
    s.addText("目標 " + p.targetDate, {
      x: rcx, y: y + 0.3, w: COLS[4], h: 0.2,
      fontSize: 7.5, color: C.muted, fontFace: "Calibri", align: "center"
    });
    rcx += COLS[4];

    var driftTxt, driftColor;
    if (h.forecastDrift > 0)      { driftTxt = "+" + h.forecastDrift + " 天"; driftColor = C.red; }
    else if (h.forecastDrift < 0) { driftTxt = h.forecastDrift + " 天";       driftColor = C.green; }
    else                          { driftTxt = "準時";                        driftColor = C.muted; }
    s.addText(driftTxt, {
      x: rcx, y: y, w: COLS[5], h: ROW_H,
      fontSize: 12, bold: true, color: driftColor, fontFace: "Calibri", align: "center", valign: "middle"
    });
    rcx += COLS[5];

    var projDec = (data.decisions || []).filter(function(d) { return d.project === p.name; });
    var keyText  = projDec.length > 0 ? projDec[0].issue : "進度正常";
    var keyColor = projDec.length > 0
      ? (projDec[0].severity === "高" ? C.red : C.amber)
      : C.green;
    s.addText(keyText, {
      x: rcx + 0.1, y: y, w: COLS[6] - 0.2, h: ROW_H,
      fontSize: 9, color: keyColor, fontFace: "Calibri",
      valign: "middle", italic: projDec.length === 0
    });
  });

  var tableEnd = TY + 0.3 + data.projects.length * ROW_H + 0.12;
  s.addText("需主管裁示事項：" + data.stats.decisionCount + " 件　·　詳見後續頁面", {
    x: 0.3, y: tableEnd, w: 9.4, h: 0.28,
    fontSize: 9.5, color: C.muted, fontFace: "Calibri", italic: true
  });
}

// ══════════════════════════════════════════════════════
// Slide 3: Portfolio Dashboard
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "Portfolio Dashboard", "任務統計 ｜ " + data.reportDate);

  var stats = data.stats;
  var kpiCards = [
    { label: "專案總數",    value: stats.totalProjects,   color: C.navy  },
    { label: "總任務數",    value: stats.totalTasks,      color: C.blue  },
    { label: "已完成",      value: stats.doneTasks,       color: C.green },
    { label: "進行中",      value: stats.inProgressTasks, color: C.blue  },
    { label: "風險 / 延遲", value: stats.riskTasks,       color: C.red   },
  ];
  kpiCards.forEach(function(kc, i) {
    var x = 0.35 + i * 1.86;
    card(s, x, 0.82, 1.7, 1.08, { shadow: true });
    s.addShape("rect", { x: x, y: 0.82, w: 1.7, h: 0.06, fill: { color: kc.color }, line: { color: kc.color } });
    s.addText(String(kc.value), {
      x: x, y: 0.93, w: 1.7, h: 0.6, fontSize: 34, bold: true, color: kc.color,
      fontFace: "Calibri", align: "center", valign: "middle"
    });
    s.addText(kc.label, {
      x: x, y: 1.58, w: 1.7, h: 0.28, fontSize: 10, color: C.muted, fontFace: "Calibri", align: "center"
    });
  });

  var doneRate = stats.totalTasks > 0 ? Math.round(stats.doneTasks / stats.totalTasks * 100) : 0;
  s.addChart(pres.charts.DOUGHNUT, [{
    name: "完成率", labels: ["已完成", "未完成"],
    values: [stats.doneTasks, stats.totalTasks - stats.doneTasks]
  }], {
    x: 0.3, y: 2.1, w: 3.2, h: 3.2,
    chartColors: [C.green, "E8EEF4"],
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: C.muted,
    holeSize: 60,
    chartArea: { fill: { color: C.bg } },
    showTitle: true, title: "整體完成率 " + doneRate + "%",
    titleFontSize: 13, titleColor: C.navy,
  });

  var labels   = data.projects.map(function(p) { return p.name; });
  var dataDone = data.projects.map(function(p) { return p.doneTasks; });
  var dataInP  = data.projects.map(function(p) { return p.inProgressTasks; });
  var dataRisk = data.projects.map(function(p) { return p.riskTasks; });
  var dataNS   = data.projects.map(function(p) { return p.notStartedTasks; });

  s.addChart(pres.charts.BAR, [
    { name: "已完成",    labels: labels, values: dataDone },
    { name: "進行中",    labels: labels, values: dataInP  },
    { name: "風險/延遲", labels: labels, values: dataRisk },
    { name: "待開始",    labels: labels, values: dataNS   },
  ], {
    x: 3.65, y: 2.0, w: 6.0, h: 3.3,
    barDir: "bar", barGrouping: "stacked",
    chartColors: [C.green, C.blue, C.red, "B0BEC5"],
    showLegend: true, legendPos: "b", legendFontSize: 10,
    catAxisLabelColor: C.text, catAxisLabelFontSize: 11, catAxisLabelBold: true,
    valAxisLabelColor: C.muted, valAxisLabelFontSize: 9,
    valGridLine: { color: C.divider, size: 0.5 }, catGridLine: { style: "none" },
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    showTitle: true, title: "各專案任務分佈",
    titleFontSize: 13, titleColor: C.navy, shadow: makeShadow()
  });
}

// ══════════════════════════════════════════════════════
// Slides 4~N: 每專案 2 頁
// ══════════════════════════════════════════════════════
data.projects.forEach(function(proj) {
  var h       = proj.health;
  var hColor  = HEALTH_COLOR[h.level];
  var hSoft   = HEALTH_SOFT[h.level];
  var projDec = (data.decisions || []).filter(function(d) { return d.project === proj.name; });

  // ─────────────────────────────────────────────────
  // Page A: 專案概況
  // ─────────────────────────────────────────────────
  {
    var s = pres.addSlide();
    s.background = { color: C.bg };
    slideHeader(s, proj.name, proj.owner + " ｜ " + proj.startDate + " ～ " + proj.targetDate, "概況  1 / 2");

    var LX = 0.25, LY = 0.75, LW = 3.1;

    // ── 健康度卡
    card(s, LX, LY, LW, 1.42, { shadow: true, accentLeft: hColor });
    s.addShape("ellipse", { x: LX + 0.22, y: LY + 0.22, w: 0.28, h: 0.28, fill: { color: hColor }, line: { color: hColor } });
    s.addText(HEALTH_LABEL[h.level], {
      x: LX + 0.6, y: LY + 0.18, w: 1.4, h: 0.36,
      fontSize: 18, bold: true, color: hColor, fontFace: "Calibri", valign: "middle"
    });
    // 完成度：取整數、字型縮小、%留在同一行
    s.addText(Math.round(h.donePct) + "%", {
      x: LX + 0.12, y: LY + 0.52, w: 1.3, h: 0.46,
      fontSize: 28, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
    });
    s.addText("任務完成", {
      x: LX + 0.12, y: LY + 0.98, w: 1.1, h: 0.2,
      fontSize: 8.5, color: C.muted, fontFace: "Calibri"
    });

    // 進度條（放右側，y 對齊中間）
    var pbarX = LX + 1.45, pbarY = LY + 0.60, pbarW = LW - 1.62, pbarH = 0.13;
    s.addText("完成率", { x: pbarX, y: pbarY - 0.19, w: pbarW, h: 0.16, fontSize: 7.5, color: C.muted, fontFace: "Calibri" });
    progressBar(s, pbarX, pbarY, pbarW, pbarH, h.donePct, hColor);
    var tickX = pbarX + pbarW * Math.min(100, h.elapsedPct) / 100;
    s.addShape("rect", { x: tickX - 0.01, y: pbarY - 0.04, w: 0.02, h: pbarH + 0.08, fill: { color: C.text }, line: { color: C.text } });
    s.addText(Math.round(h.donePct) + "%", { x: pbarX, y: pbarY + 0.16, w: pbarW * 0.5, h: 0.16, fontSize: 7.5, bold: true, color: hColor, fontFace: "Calibri" });
    s.addText("時間已過 " + h.elapsedPct + "%", { x: pbarX + pbarW * 0.5, y: pbarY + 0.16, w: pbarW * 0.5, h: 0.16, fontSize: 7.5, color: C.muted, fontFace: "Calibri", align: "right" });

    // ── 預估完工卡
    var fcY = LY + 1.55;
    card(s, LX, fcY, LW, 0.88, { shadow: true });
    s.addText("預估完工", { x: LX + 0.15, y: fcY + 0.08, w: 1.2, h: 0.22, fontSize: 9, color: C.muted, fontFace: "Calibri" });
    s.addText(h.forecastEnd || proj.targetDate, {
      x: LX + 0.15, y: fcY + 0.27, w: 1.7, h: 0.32,
      fontSize: 16, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
    });
    var driftTxt = "準時", driftColor = C.muted;
    if (h.forecastDrift > 0)      { driftTxt = "落後 " + h.forecastDrift + " 天"; driftColor = C.red; }
    else if (h.forecastDrift < 0) { driftTxt = "提前 " + (-h.forecastDrift) + " 天"; driftColor = C.green; }
    s.addText(driftTxt, { x: LX + 0.15, y: fcY + 0.6, w: 1.7, h: 0.22, fontSize: 11, bold: true, color: driftColor, fontFace: "Calibri" });
    s.addText("目標　" + proj.targetDate, {
      x: LX + 1.9, y: fcY + 0.08, w: 1.08, h: 0.22,
      fontSize: 8.5, color: C.muted, fontFace: "Calibri", align: "right"
    });
    var biasBarX = LX + 1.9, biasBarY = fcY + 0.36, biasBarW = 1.0, biasBarH = 0.12;
    s.addShape("roundRect", { x: biasBarX, y: biasBarY, w: biasBarW, h: biasBarH, fill: { color: C.divider }, line: { color: C.divider }, rectRadius: 0.02 });
    var biasRatio = Math.min(1.0, Math.abs(h.forecastDrift) / 90);
    if (biasRatio > 0.01) {
      s.addShape("roundRect", { x: biasBarX, y: biasBarY, w: biasBarW * biasRatio, h: biasBarH, fill: { color: driftColor }, line: { color: driftColor }, rectRadius: 0.02 });
    }

    // ── 任務統計
    var statsY = fcY + 1.02;
    card(s, LX, statsY, LW, 0.75, { shadow: true });
    var statItems = [
      { label: "總任務", val: proj.totalTasks,      color: C.navy  },
      { label: "完成",   val: proj.doneTasks,       color: C.green },
      { label: "進行中", val: proj.inProgressTasks, color: C.blue  },
      { label: "風險",   val: proj.riskTasks,       color: C.red   },
      { label: "待開始", val: proj.notStartedTasks, color: C.muted },
    ];
    var boxW = LW / statItems.length;
    statItems.forEach(function(st, i) {
      var bx = LX + i * boxW;
      if (i > 0) s.addShape("rect", { x: bx, y: statsY + 0.1, w: 0.01, h: 0.55, fill: { color: C.divider }, line: { color: C.divider } });
      s.addText(String(st.val), {
        x: bx, y: statsY + 0.08, w: boxW, h: 0.38,
        fontSize: 22, bold: true, color: st.color, fontFace: "Calibri", align: "center", valign: "middle"
      });
      s.addText(st.label, {
        x: bx, y: statsY + 0.48, w: boxW, h: 0.2,
        fontSize: 8, color: C.muted, fontFace: "Calibri", align: "center"
      });
    });

    // ── 階段目標
    var stageY = statsY + 0.89;
    card(s, LX, stageY, LW, 0.52, { fill: C.paleBlue, border: C.lightBlue });
    s.addText("階段目標", { x: LX + 0.12, y: stageY + 0.05, w: 0.7, h: 0.18, fontSize: 8, color: C.muted, fontFace: "Calibri" });
    s.addText(proj.stage || "—", {
      x: LX + 0.12, y: stageY + 0.22, w: LW - 0.22, h: 0.26,
      fontSize: 9.5, color: C.text, fontFace: "Calibri", valign: "middle"
    });

    // ── 里程碑甘特（右側）
    var GCX = LX + LW + 0.2, GCY = LY, GCW = 10 - GCX - 0.25, GCH = 4.52;
    card(s, GCX, GCY, GCW, GCH, { shadow: true });
    s.addShape("rect", { x: GCX, y: GCY, w: GCW, h: 0.34, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText("里程碑進度", {
      x: GCX + 0.15, y: GCY, w: 2.5, h: 0.34,
      fontSize: 11, bold: true, color: C.white, fontFace: "Calibri", valign: "middle"
    });

    // 圖例（固定位置，避免文字框太窄變直排）
    var legendDefs = [
      { label: "計畫範圍", color: C.divider, x: GCX + GCW - 2.38 },
      { label: "完成進度", color: hColor,    x: GCX + GCW - 1.62 },
      { label: "今日",     color: C.accent,  x: GCX + GCW - 0.82 },
    ];
    legendDefs.forEach(function(lg) {
      s.addShape("rect", { x: lg.x, y: GCY + 0.11, w: 0.14, h: 0.12, fill: { color: lg.color }, line: { color: lg.color } });
      s.addText(lg.label, {
        x: lg.x + 0.18, y: GCY + 0.08, w: 0.6, h: 0.18,
        fontSize: 7.5, color: C.white, fontFace: "Calibri", valign: "middle"
      });
    });

    if (proj.milestones && proj.milestones.length > 0) {
      var ms = proj.milestones;
      var allDates = [];
      ms.forEach(function(m) { allDates.push(new Date(m.startDate)); allDates.push(new Date(m.endDate)); });
      var minD   = new Date(Math.min.apply(null, allDates));
      var maxD   = new Date(Math.max.apply(null, allDates));
      var totalD = (maxD - minD) / 86400000 || 1;

      var LABEL_W  = 1.62;          // 名稱欄寬（含%）
      var PCT_W    = 0.38;           // %欄寬（固定在名稱欄右側）
      var GX       = GCX + LABEL_W;
      var GW       = GCW - LABEL_W - 0.12;
      var GY_HDR   = GCY + 0.34;
      var GY_START = GY_HDR + 0.24;
      var MAX_MS   = Math.min(ms.length, 10);   // 最多顯示 10 筆
      var ROW_AVAIL = GCH - 0.34 - 0.24 - 0.22;
      var ROW_H_MS  = Math.min(0.52, Math.max(0.29, ROW_AVAIL / MAX_MS));

      // 月份刻度
      var cur = new Date(minD);
      cur.setDate(1);
      while (cur <= maxD) {
        var lx = GX + (cur - minD) / 86400000 / totalD * GW;
        if (lx >= GX - 0.02 && lx <= GX + GW) {
          s.addShape("rect", { x: lx, y: GY_HDR, w: 0.01, h: GCH - 0.34 - 0.12, fill: { color: C.divider }, line: { color: C.divider } });
          s.addText((cur.getMonth() + 1) + "月", {
            x: lx + 0.03, y: GY_HDR + 0.02, w: 0.38, h: 0.2,
            fontSize: 7, color: C.muted, fontFace: "Calibri"
          });
        }
        cur.setMonth(cur.getMonth() + 1);
      }

      // 今日線
      var today = new Date();
      if (today >= minD && today <= maxD) {
        var todayX = GX + (today - minD) / 86400000 / totalD * GW;
        s.addShape("rect", {
          x: todayX - 0.01, y: GY_HDR, w: 0.02, h: GCH - 0.34 - 0.25,
          fill: { color: C.accent }, line: { color: C.accent }
        });
        s.addText("今日", {
          x: todayX - 0.22, y: GCY + GCH - 0.28, w: 0.44, h: 0.2,
          fontSize: 7, color: C.accent, bold: true, fontFace: "Calibri", align: "center"
        });
      }

      // 各里程碑列（最多 MAX_MS 筆）
      ms.slice(0, MAX_MS).forEach(function(m, i) {
        var rowY  = GY_START + i * ROW_H_MS;
        var barCY = rowY + ROW_H_MS / 2 - 0.065;
        var color = STATUS_COLOR[m.status] || C.blue;
        var rowBg = i % 2 === 0 ? C.white : C.paleBlue;

        s.addShape("rect", { x: GCX, y: rowY, w: GCW, h: ROW_H_MS, fill: { color: rowBg }, line: { color: C.divider, pt: 0.3 } });

        // 名稱（左側，保留%欄位置）
        var nameFontSize = ROW_H_MS >= 0.38 ? 8.5 : 7.5;
        s.addText(m.name, {
          x: GCX + 0.1, y: rowY, w: LABEL_W - PCT_W - 0.12, h: ROW_H_MS,
          fontSize: nameFontSize, bold: true, color: C.text, fontFace: "Calibri",
          valign: "middle", shrinkText: true
        });
        // %數字（固定欄位，不擠壓名稱）
        s.addText(m.progress + "%", {
          x: GCX + LABEL_W - PCT_W, y: rowY, w: PCT_W - 0.05, h: ROW_H_MS,
          fontSize: nameFontSize, bold: true, color: color, fontFace: "Calibri",
          valign: "middle", align: "right"
        });

        var x1 = GX + (new Date(m.startDate) - minD) / 86400000 / totalD * GW;
        var x2 = GX + (new Date(m.endDate)   - minD) / 86400000 / totalD * GW;
        var bw  = Math.max(0.12, x2 - x1);
        var barH = Math.min(0.15, ROW_H_MS * 0.48);

        s.addShape("roundRect", { x: GX, y: barCY, w: GW, h: barH, fill: { color: C.divider }, line: { color: C.divider }, rectRadius: 0.02 });
        var fillW = Math.max(0.05, bw * m.progress / 100);
        s.addShape("roundRect", { x: x1, y: barCY, w: fillW, h: barH, fill: { color: color }, line: { color: color }, rectRadius: 0.02 });
        s.addShape("roundRect", { x: x1, y: barCY, w: bw, h: barH, fill: { type: "none" }, line: { color: color, pt: 0.7 }, rectRadius: 0.02 });

        // 任務數（bar 右側，空間夠才顯示）
        if (x1 + bw + 0.08 < GX + GW - 0.3) {
          s.addText(m.doneTasks + "/" + m.totalTasks, {
            x: x1 + bw + 0.05, y: rowY, w: 0.45, h: ROW_H_MS,
            fontSize: 7, color: C.muted, fontFace: "Calibri", valign: "middle"
          });
        }
      });

      // 若超過 10 筆，顯示省略提示
      if (ms.length > MAX_MS) {
        var moreY = GY_START + MAX_MS * ROW_H_MS + 0.04;
        s.addText("…共 " + ms.length + " 個里程碑，僅顯示前 " + MAX_MS + " 筆", {
          x: GCX + 0.15, y: moreY, w: GCW - 0.3, h: 0.2,
          fontSize: 7.5, color: C.muted, italic: true, fontFace: "Calibri"
        });
      }
    } else {
      s.addText("（無里程碑資料）", {
        x: GCX + 0.3, y: GCY + GCH / 2 - 0.2, w: GCW - 0.6, h: 0.4,
        fontSize: 11, color: C.muted, italic: true, fontFace: "Calibri", align: "center"
      });
    }
  }

  // ─────────────────────────────────────────────────
  // Page B: 本期動態與風險
  // ─────────────────────────────────────────────────
  {
    var s = pres.addSlide();
    s.background = { color: C.bg };
    slideHeader(s, proj.name + "　—　本期動態", "近 14 天回顧 ｜ " + data.reportDate, "動態  2 / 2");

    // ── KPI 條
    var kpiY = 0.78, kpiH = 0.88;
    var kpiItems = [
      { label: "本期完成任務",      val: proj.activity.recentDone.length, color: C.green },
      { label: "進行中 / 即將到期", val: proj.activity.upcoming.length,   color: C.blue  },
      { label: "風險 / 需裁示",     val: projDec.length,                  color: C.red   },
      { label: "待開始任務",        val: proj.notStartedTasks,            color: C.muted },
    ];
    var kpiW = 9.5 / kpiItems.length;
    kpiItems.forEach(function(kc, i) {
      var kx = 0.25 + i * kpiW;
      card(s, kx, kpiY, kpiW - 0.15, kpiH, { shadow: true, accentLeft: kc.color });
      s.addText(String(kc.val), {
        x: kx + 0.2, y: kpiY + 0.08, w: 0.8, h: 0.55,
        fontSize: 30, bold: true, color: kc.color, fontFace: "Calibri", valign: "middle"
      });
      s.addText(kc.label, {
        x: kx + 1.05, y: kpiY + 0.28, w: kpiW - 1.3, h: 0.38,
        fontSize: 9.5, color: C.muted, fontFace: "Calibri", valign: "middle"
      });
    });

    // ── 三欄面板
    var PANELY = kpiY + kpiH + 0.12;
    var PANELH = 5.625 - PANELY - 0.22;
    var COL_X  = [0.25, 3.4, 6.55];
    var COL_W  = [3.0, 3.0, 3.2];

    function makePanel(idx, title, color) {
      var px = COL_X[idx], pw = COL_W[idx];
      card(s, px, PANELY, pw, PANELH, { shadow: true });
      s.addShape("rect", { x: px, y: PANELY, w: pw, h: 0.33, fill: { color: color }, line: { color: color } });
      s.addText(title, {
        x: px + 0.15, y: PANELY, w: pw - 0.3, h: 0.33,
        fontSize: 10.5, bold: true, color: C.white, fontFace: "Calibri", valign: "middle", margin: 0
      });
    }

    var ITEM_START_Y = PANELY + 0.43;
    var ITEM_H       = 0.48;
    var MAX_ITEMS    = Math.floor((PANELH - 0.43 - 0.15) / ITEM_H);

    // 左欄：近14天完成
    makePanel(0, "近 14 天完成（" + proj.activity.recentDone.length + "）", C.green);
    if (proj.activity.recentDone.length === 0) {
      s.addText("本期無完成任務", {
        x: COL_X[0] + 0.15, y: ITEM_START_Y + 0.5, w: COL_W[0] - 0.3, h: 0.35,
        fontSize: 10, italic: true, color: C.muted, fontFace: "Calibri", align: "center"
      });
    } else {
      proj.activity.recentDone.slice(0, MAX_ITEMS).forEach(function(item, i) {
        var iy = ITEM_START_Y + i * ITEM_H;
        s.addShape("ellipse", { x: COL_X[0] + 0.18, y: iy + 0.1, w: 0.1, h: 0.1, fill: { color: C.green }, line: { color: C.green } });
        s.addText(item.name, {
          x: COL_X[0] + 0.35, y: iy, w: COL_W[0] - 0.5, h: 0.26,
          fontSize: 9.5, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
        });
        s.addText(item.owner + " · " + item.date, {
          x: COL_X[0] + 0.35, y: iy + 0.26, w: COL_W[0] - 0.5, h: 0.18,
          fontSize: 8, color: C.muted, fontFace: "Calibri"
        });
      });
    }
    var doneThisPeriod = proj.activity.recentDone.length;
    s.addText("本期完成率　" + doneThisPeriod + "/" + proj.totalTasks + " = " + (proj.totalTasks > 0 ? Math.round(doneThisPeriod / proj.totalTasks * 100) : 0) + "%", {
      x: COL_X[0] + 0.15, y: PANELY + PANELH - 0.28, w: COL_W[0] - 0.3, h: 0.22,
      fontSize: 8, color: C.muted, fontFace: "Calibri", italic: true
    });

    // 中欄：進行中 / 即將到期
    makePanel(1, "進行中 / 即將到期（" + proj.activity.upcoming.length + "）", C.blue);
    if (proj.activity.upcoming.length === 0) {
      s.addText("目前無進行中任務", {
        x: COL_X[1] + 0.15, y: ITEM_START_Y + 0.5, w: COL_W[1] - 0.3, h: 0.35,
        fontSize: 10, italic: true, color: C.muted, fontFace: "Calibri", align: "center"
      });
    } else {
      proj.activity.upcoming.slice(0, MAX_ITEMS).forEach(function(item, i) {
        var iy = ITEM_START_Y + i * ITEM_H;
        var stColor = STATUS_COLOR[item.status] || C.blue;
        var isRisk  = item.status === "延遲" || item.status === "風險";
        if (isRisk) {
          s.addShape("rect", {
            x: COL_X[1] + 0.1, y: iy - 0.03, w: COL_W[1] - 0.2, h: ITEM_H - 0.02,
            fill: { color: C.redSoft }, line: { color: C.redSoft }
          });
        }
        s.addShape("ellipse", { x: COL_X[1] + 0.18, y: iy + 0.1, w: 0.1, h: 0.1, fill: { color: stColor }, line: { color: stColor } });
        s.addText(item.name, {
          x: COL_X[1] + 0.35, y: iy, w: COL_W[1] - 0.5, h: 0.26,
          fontSize: 9.5, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
        });
        s.addShape("roundRect", {
          x: COL_X[1] + 0.35, y: iy + 0.28, w: 0.38, h: 0.16,
          fill: { color: stColor }, line: { color: stColor }, rectRadius: 0.02
        });
        s.addText(item.status, {
          x: COL_X[1] + 0.35, y: iy + 0.28, w: 0.38, h: 0.16,
          fontSize: 7, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0
        });
        s.addText(item.owner + " · 截止 " + item.date, {
          x: COL_X[1] + 0.78, y: iy + 0.28, w: COL_W[1] - 0.93, h: 0.18,
          fontSize: 8, color: C.muted, fontFace: "Calibri"
        });
      });
    }

    // 右欄：風險與裁示
    makePanel(2, "風險議題（" + projDec.length + "）", C.red);
    if (projDec.length === 0) {
      s.addShape("roundRect", {
        x: COL_X[2] + 0.3, y: ITEM_START_Y + 0.6, w: COL_W[2] - 0.6, h: 0.7,
        fill: { color: C.greenSoft }, line: { color: C.greenSoft }, rectRadius: 0.06
      });
      s.addText("✓  目前無風險議題", {
        x: COL_X[2] + 0.3, y: ITEM_START_Y + 0.6, w: COL_W[2] - 0.6, h: 0.7,
        fontSize: 12, bold: true, color: C.green, fontFace: "Calibri", align: "center", valign: "middle"
      });
    } else {
      var RISK_ITEM_H = Math.min(0.9, (PANELH - 0.43 - 0.1) / Math.min(projDec.length, 4));
      projDec.slice(0, 4).forEach(function(d, i) {
        var iy       = ITEM_START_Y + i * RISK_ITEM_H;
        var sevColor = d.severity === "高" ? C.red   : (d.severity === "中" ? C.amber   : C.muted);
        var sevSoft  = d.severity === "高" ? C.redSoft : (d.severity === "中" ? C.amberSoft : C.divider);
        var sevDeep  = d.severity === "高" ? C.redDeep : (d.severity === "中" ? C.amberDeep : C.text);

        s.addShape("roundRect", {
          x: COL_X[2] + 0.12, y: iy, w: COL_W[2] - 0.24, h: RISK_ITEM_H - 0.1,
          fill: { color: sevSoft }, line: { color: sevColor, pt: 0.5 }, rectRadius: 0.05
        });
        s.addShape("rect", {
          x: COL_X[2] + 0.12, y: iy, w: 0.08, h: RISK_ITEM_H - 0.1,
          fill: { color: sevColor }, line: { color: sevColor }
        });
        s.addShape("roundRect", {
          x: COL_X[2] + 0.26, y: iy + 0.07, w: 0.32, h: 0.18,
          fill: { color: sevColor }, line: { color: sevColor }, rectRadius: 0.02
        });
        s.addText(d.severity, {
          x: COL_X[2] + 0.26, y: iy + 0.07, w: 0.32, h: 0.18,
          fontSize: 8, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0
        });
        s.addText(d.owner, {
          x: COL_X[2] + 0.62, y: iy + 0.07, w: COL_W[2] - 0.9, h: 0.18,
          fontSize: 8, color: sevDeep, fontFace: "Calibri", valign: "middle"
        });
        s.addText(d.issue, {
          x: COL_X[2] + 0.26, y: iy + 0.27, w: COL_W[2] - 0.5, h: 0.26,
          fontSize: 9, bold: true, color: sevDeep, fontFace: "Calibri", valign: "middle"
        });
        s.addText("→ " + d.ask, {
          x: COL_X[2] + 0.26, y: iy + 0.52, w: COL_W[2] - 0.5, h: 0.28,
          fontSize: 8.5, color: sevColor, bold: true, fontFace: "Calibri", valign: "middle"
        });
      });
    }
  }
});

// ══════════════════════════════════════════════════════
// Slide N+1: 需主管裁示事項
// ══════════════════════════════════════════════════════
{
  var s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "需主管裁示事項", "按嚴重度排序 ｜ " + data.decisions.length + " 件");

  if (!data.decisions || data.decisions.length === 0) {
    card(s, 2.5, 2.0, 5.0, 1.5, { shadow: true });
    s.addText("✓  目前無需裁示事項", {
      x: 2.5, y: 2.0, w: 5.0, h: 1.5,
      fontSize: 18, bold: true, color: C.green, fontFace: "Calibri", align: "center", valign: "middle"
    });
  } else {
    var HX = 0.3, HY = 0.82, HW = 9.4;
    var COLW = [0.55, 1.35, 2.65, 2.45, 2.4];
    var HDRS = ["", "專案", "議題", "影響評估", "建議行動"];
    var hcx = HX;
    HDRS.forEach(function(hd, i) {
      s.addShape("rect", { x: hcx, y: HY, w: COLW[i], h: 0.34, fill: { color: C.navy }, line: { color: C.navy } });
      s.addText(hd, {
        x: hcx, y: HY, w: COLW[i], h: 0.34,
        fontSize: 10, bold: true, color: C.white, fontFace: "Calibri",
        align: "center", valign: "middle", margin: 0
      });
      hcx += COLW[i];
    });

    var RH = 0.74;
    data.decisions.slice(0, 6).forEach(function(d, i) {
      var y = HY + 0.34 + i * RH;
      var sevColor = d.severity === "高" ? C.red : (d.severity === "中" ? C.amber : C.muted);
      var rowBg = i % 2 === 0 ? C.white : C.paleBlue;
      s.addShape("rect", { x: HX, y: y, w: HW, h: RH, fill: { color: rowBg }, line: { color: C.divider, pt: 0.5 } });

      var dcx = HX;
      s.addShape("rect", { x: dcx + 0.08, y: y + 0.14, w: COLW[0] - 0.16, h: RH - 0.28, fill: { color: sevColor }, line: { color: sevColor } });
      s.addText(d.severity, {
        x: dcx + 0.08, y: y + 0.14, w: COLW[0] - 0.16, h: RH - 0.28,
        fontSize: 13, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0
      });
      dcx += COLW[0];

      s.addText(d.project, { x: dcx + 0.08, y: y + 0.06, w: COLW[1] - 0.14, h: 0.3, fontSize: 9.5, bold: true, color: C.text, fontFace: "Calibri", valign: "middle" });
      s.addText("負責人 " + d.owner, { x: dcx + 0.08, y: y + 0.38, w: COLW[1] - 0.14, h: 0.28, fontSize: 8, color: C.muted, fontFace: "Calibri" });
      dcx += COLW[1];

      s.addText(d.issue,  { x: dcx + 0.1, y: y + 0.06, w: COLW[2] - 0.2, h: RH - 0.12, fontSize: 9.5, color: C.text, fontFace: "Calibri", valign: "middle" });
      dcx += COLW[2];
      s.addText(d.impact, { x: dcx + 0.1, y: y + 0.06, w: COLW[3] - 0.2, h: RH - 0.12, fontSize: 9, color: C.muted, fontFace: "Calibri", valign: "middle", italic: true });
      dcx += COLW[3];
      s.addText(d.ask,    { x: dcx + 0.1, y: y + 0.06, w: COLW[4] - 0.2, h: RH - 0.12, fontSize: 9.5, bold: true, color: sevColor, fontFace: "Calibri", valign: "middle" });
    });
  }
}

// ══════════════════════════════════════════════════════
// 結尾頁
// ══════════════════════════════════════════════════════
{
  var s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape("rect", { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });
  s.addShape("rect", { x: 0.18, y: 0, w: 0.08, h: 5.625, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText("謝謝聆聽", {
    x: 1, y: 1.7, w: 8, h: 1.0,
    fontSize: 40, bold: true, color: C.white, fontFace: "Calibri", align: "center", charSpacing: 8
  });
  s.addText("本期報告日期：" + data.reportDate, {
    x: 1, y: 2.95, w: 8, h: 0.38, fontSize: 13, color: "CADCFC", fontFace: "Calibri", align: "center"
  });
  s.addText("下次回報日期：" + data.nextReport, {
    x: 1, y: 3.35, w: 8, h: 0.38, fontSize: 13, color: "CADCFC", fontFace: "Calibri", align: "center"
  });
}

// ── 輸出
pres.writeFile({ fileName: outputPath }).then(function() {
  console.log("PPT 產生完成：" + outputPath);
}).catch(function(e) {
  console.error("失敗：", e);
  process.exit(1);
});