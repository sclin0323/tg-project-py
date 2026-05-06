/**
 * generate_ppt.js — 執行摘要導向版本
 * 用法: node generate_ppt.js <data.json> <output.pptx>
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");

const dataPath   = process.argv[2];
const outputPath = process.argv[3];
const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));

// ── 色盤 (企業藍 + RAG) ────────────────────────────────
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
  amber:     "F59E0B",
  green:     "16A34A",
  redSoft:   "FEE2E2",
  amberSoft: "FEF3C7",
  greenSoft: "D1FAE5",
};

const HEALTH_COLOR = { RED: C.red, AMBER: C.amber, GREEN: C.green };
const HEALTH_SOFT  = { RED: C.redSoft, AMBER: C.amberSoft, GREEN: C.greenSoft };
const HEALTH_LABEL = { RED: "警報", AMBER: "注意", GREEN: "健康" };
const HEALTH_ICON  = { RED: "🔴", AMBER: "🟡", GREEN: "🟢" };

const STATUS_COLOR = {
  "完成":  C.green,
  "進行中": C.blue,
  "延遲":  C.red,
  "風險":  C.red,
  "待開始": C.muted,
};

// ── 共用工具 ────────────────────────────────────────────
function makeShadow() {
  return { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 };
}

function slideHeader(slide, title, subtitle) {
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.6, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText(title, {
    x: 0.4, y: 0, w: 6.5, h: 0.6,
    fontSize: 18, bold: true, color: C.white, fontFace: "Calibri",
    valign: "middle", margin: 0
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0, y: 0, w: 9.6, h: 0.6,
      fontSize: 10, color: "CADCFC", fontFace: "Calibri",
      align: "right", valign: "middle", margin: 0
    });
  }
  slide.addShape("rect", { x: 0, y: 0.6, w: 10, h: 0.03, fill: { color: C.accent }, line: { color: C.accent } });
}

function progressBar(slide, x, y, w, h, pct, fillColor) {
  slide.addShape("rect", { x, y, w, h, fill: { color: C.divider }, line: { color: C.divider } });
  const filled = Math.max(0, Math.min(100, pct)) / 100 * w;
  if (filled > 0.02) {
    slide.addShape("rect", { x, y, w: filled, h, fill: { color: fillColor }, line: { color: fillColor } });
  }
}

// ══════════════════════════════════════════════════════
// PPT 開始
// ══════════════════════════════════════════════════════
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "專案管理報告";

// ── Slide 1: 封面（含一句話定調）────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // 右側裝飾
  s.addShape("rect", { x: 7.2, y: 0, w: 2.8, h: 5.625, fill: { color: C.blue }, line: { color: C.blue } });
  s.addShape("rect", { x: 8.8, y: 0, w: 1.2, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });

  s.addText("專案管理報告", {
    x: 0.6, y: 1.4, w: 6.4, h: 0.9,
    fontSize: 38, bold: true, color: C.white, fontFace: "Calibri", charSpacing: 4
  });
  s.addText("Executive Status Report", {
    x: 0.6, y: 2.3, w: 6.4, h: 0.4,
    fontSize: 13, color: "CADCFC", fontFace: "Calibri", italic: true
  });
  s.addShape("rect", { x: 0.6, y: 2.85, w: 3.5, h: 0.04, fill: { color: C.accent }, line: { color: C.accent } });

  // 一句話 headline
  const hlColor = HEALTH_COLOR[data.overallHealth];
  s.addShape("roundRect", {
    x: 0.6, y: 3.1, w: 6.4, h: 0.7,
    fill: { color: hlColor }, line: { color: hlColor }, rectRadius: 0.08
  });
  s.addText(data.headline, {
    x: 0.6, y: 3.1, w: 6.4, h: 0.7,
    fontSize: 17, bold: true, color: C.white, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  s.addText(`報告日期　${data.reportDate}`, {
    x: 0.6, y: 4.0, w: 6.4, h: 0.4,
    fontSize: 13, color: "CADCFC", fontFace: "Calibri"
  });
  s.addText(`下次回報　${data.nextReport}`, {
    x: 0.6, y: 4.4, w: 6.4, h: 0.4,
    fontSize: 13, color: "CADCFC", fontFace: "Calibri"
  });
}

// ── Slide 2: ★ 執行摘要（一頁掌握）─────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "執行摘要", `一頁掌握所有專案 ｜ ${data.reportDate}`);

  // 上方三個健康度燈號統計
  const summary = [
    { label: "需立即關注", count: data.stats.redProjects,   color: C.red,   level: "RED"   },
    { label: "需要觀察",   count: data.stats.amberProjects, color: C.amber, level: "AMBER" },
    { label: "進度健康",   count: data.stats.greenProjects, color: C.green, level: "GREEN" },
  ];
  summary.forEach((card, i) => {
    const x = 0.4 + i * 3.13;
    s.addShape("rect", { x, y: 0.8, w: 2.95, h: 0.85, fill: { color: C.white }, line: { color: HEALTH_SOFT[card.level], pt: 1 }, shadow: makeShadow() });
    s.addShape("rect", { x, y: 0.8, w: 0.08, h: 0.85, fill: { color: card.color }, line: { color: card.color } });
    // 數字
    s.addText(String(card.count), {
      x: x + 0.2, y: 0.83, w: 0.9, h: 0.78,
      fontSize: 38, bold: true, color: card.color, fontFace: "Calibri", valign: "middle"
    });
    s.addText(card.label, {
      x: x + 1.2, y: 0.85, w: 1.6, h: 0.4,
      fontSize: 11, color: C.muted, fontFace: "Calibri", valign: "middle"
    });
    s.addText("個專案", {
      x: x + 1.2, y: 1.2, w: 1.6, h: 0.4,
      fontSize: 14, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
    });
  });

  // 主表格：每專案一行（最關鍵）
  const TABLE_X = 0.3, TABLE_Y = 1.95, TABLE_W = 9.4;
  const COLS = [0.5, 1.7, 0.85, 1.55, 1.2, 0.95, 2.65];  // 燈號|專案|完成%|進度條|預估完工|偏差|主要關鍵

  // 表頭
  const headers = ["", "專案", "完成", "進度", "預估完工", "偏差", "主要關鍵"];
  let cx = TABLE_X;
  headers.forEach((h, i) => {
    s.addShape("rect", { x: cx, y: TABLE_Y, w: COLS[i], h: 0.32, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(h, {
      x: cx, y: TABLE_Y, w: COLS[i], h: 0.32,
      fontSize: 9.5, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    cx += COLS[i];
  });

  // 每專案一行
  const ROW_H = 0.55;
  data.projects.forEach((p, i) => {
    const y = TABLE_Y + 0.32 + i * ROW_H;
    const rowBg = i % 2 === 0 ? C.white : C.paleBlue;
    s.addShape("rect", { x: TABLE_X, y, w: TABLE_W, h: ROW_H, fill: { color: rowBg }, line: { color: C.divider, pt: 0.5 } });

    let cx = TABLE_X;
    const h = p.health;
    const hColor = HEALTH_COLOR[h.level];

    // 1. 燈號 (圓圈)
    s.addShape("ellipse", {
      x: cx + (COLS[0] - 0.32) / 2, y: y + (ROW_H - 0.32) / 2, w: 0.32, h: 0.32,
      fill: { color: hColor }, line: { color: hColor }
    });
    cx += COLS[0];

    // 2. 專案名 + 負責人
    s.addText(p.name, {
      x: cx + 0.1, y: y + 0.05, w: COLS[1] - 0.1, h: 0.28,
      fontSize: 11.5, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
    });
    s.addText(`${p.owner} · ${HEALTH_LABEL[h.level]}`, {
      x: cx + 0.1, y: y + 0.3, w: COLS[1] - 0.1, h: 0.22,
      fontSize: 9, color: hColor, fontFace: "Calibri", valign: "middle"
    });
    cx += COLS[1];

    // 3. 完成%
    s.addText(`${h.donePct}%`, {
      x: cx, y: y, w: COLS[2], h: ROW_H,
      fontSize: 18, bold: true, color: C.text, fontFace: "Calibri",
      align: "center", valign: "middle"
    });
    cx += COLS[2];

    // 4. 進度條 (含時間經過虛線標示)
    const barX = cx + 0.1, barW = COLS[3] - 0.2, barY = y + ROW_H/2 - 0.08;
    progressBar(s, barX, barY, barW, 0.16, h.donePct, hColor);
    // 時間經過刻度
    const elapsedX = barX + barW * h.elapsedPct / 100;
    s.addShape("rect", {
      x: elapsedX - 0.01, y: barY - 0.04, w: 0.02, h: 0.24,
      fill: { color: C.text }, line: { color: C.text }
    });
    s.addText(`已過 ${h.elapsedPct}%`, {
      x: barX, y: y + ROW_H - 0.18, w: barW, h: 0.16,
      fontSize: 7.5, color: C.muted, fontFace: "Calibri", align: "center"
    });
    cx += COLS[3];

    // 5. 預估完工
    s.addText(h.forecastEnd || "—", {
      x: cx, y: y + 0.05, w: COLS[4], h: 0.28,
      fontSize: 10.5, bold: true, color: C.text, fontFace: "Calibri",
      align: "center", valign: "middle"
    });
    s.addText(`目標 ${p.targetDate}`, {
      x: cx, y: y + 0.3, w: COLS[4], h: 0.22,
      fontSize: 8, color: C.muted, fontFace: "Calibri", align: "center"
    });
    cx += COLS[4];

    // 6. 偏差天數
    let driftTxt, driftColor;
    if (h.forecastDrift > 0) { driftTxt = `+${h.forecastDrift} 天`; driftColor = C.red; }
    else if (h.forecastDrift < 0) { driftTxt = `${h.forecastDrift} 天`; driftColor = C.green; }
    else { driftTxt = "準時"; driftColor = C.muted; }
    s.addText(driftTxt, {
      x: cx, y, w: COLS[5], h: ROW_H,
      fontSize: 13, bold: true, color: driftColor, fontFace: "Calibri",
      align: "center", valign: "middle"
    });
    cx += COLS[5];

    // 7. 主要關鍵 (取第一個高優先決策)
    const projDecisions = (data.decisions || []).filter(d => d.project === p.name);
    let keyText = "進度正常";
    let keyColor = C.green;
    if (projDecisions.length > 0) {
      keyText = projDecisions[0].issue;
      keyColor = projDecisions[0].severity === "高" ? C.red : C.amber;
    }
    s.addText(keyText, {
      x: cx + 0.1, y, w: COLS[6] - 0.2, h: ROW_H,
      fontSize: 9.5, color: keyColor, fontFace: "Calibri",
      valign: "middle", italic: keyText === "進度正常"
    });
  });

  // 底部提示
  const tableEnd = TABLE_Y + 0.32 + data.projects.length * ROW_H + 0.15;
  s.addText(`需主管裁示事項：${data.stats.decisionCount} 件　·　詳見後續頁面`, {
    x: 0.3, y: tableEnd, w: 9.4, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: "Calibri", italic: true
  });
}

// ── Slide 3: Portfolio Dashboard (任務狀態總覽) ────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "Portfolio Dashboard", `任務統計 ｜ ${data.reportDate}`);

  // KPI 卡片區 (5 張)
  const stats = data.stats;
  const cards = [
    { label: "專案總數",     value: stats.totalProjects, color: C.navy },
    { label: "總任務數",     value: stats.totalTasks,    color: C.blue },
    { label: "已完成",       value: stats.doneTasks,     color: C.green },
    { label: "進行中",       value: stats.inProgressTasks, color: C.blue },
    { label: "風險 / 延遲",  value: stats.riskTasks,     color: C.red },
  ];
  cards.forEach((card, i) => {
    const x = 0.4 + i * 1.86;
    s.addShape("rect", { x, y: 0.85, w: 1.7, h: 1.1, fill: { color: C.white }, line: { color: C.lightBlue }, shadow: makeShadow() });
    s.addShape("rect", { x, y: 0.85, w: 1.7, h: 0.06, fill: { color: card.color }, line: { color: card.color } });
    s.addText(String(card.value), {
      x, y: 0.95, w: 1.7, h: 0.6, fontSize: 32, bold: true, color: card.color,
      fontFace: "Calibri", align: "center", valign: "middle"
    });
    s.addText(card.label, {
      x, y: 1.55, w: 1.7, h: 0.32, fontSize: 10, color: C.muted, fontFace: "Calibri", align: "center"
    });
  });

  // 環形圖 - 整體完成率
  const doneRate = stats.totalTasks > 0 ? Math.round(stats.doneTasks / stats.totalTasks * 100) : 0;
  s.addChart(pres.charts.DOUGHNUT, [{
    name: "完成率", labels: ["已完成", "未完成"],
    values: [stats.doneTasks, stats.totalTasks - stats.doneTasks]
  }], {
    x: 0.3, y: 2.2, w: 3.2, h: 3.0,
    chartColors: [C.green, "E8EEF4"],
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: C.muted,
    holeSize: 60,
    chartArea: { fill: { color: C.bg } },
    showTitle: true, title: `整體完成率 ${doneRate}%`,
    titleFontSize: 13, titleColor: C.navy,
  });

  // 各專案任務狀態 - 堆疊橫條
  const labels = data.projects.map(p => p.name);
  const dataDone   = data.projects.map(p => p.doneTasks);
  const dataInProg = data.projects.map(p => p.inProgressTasks);
  const dataRisk   = data.projects.map(p => p.riskTasks);
  const dataNS     = data.projects.map(p => p.notStartedTasks);

  s.addChart(pres.charts.BAR, [
    { name: "已完成", labels, values: dataDone },
    { name: "進行中", labels, values: dataInProg },
    { name: "風險/延遲", labels, values: dataRisk },
    { name: "待開始", labels, values: dataNS },
  ], {
    x: 3.7, y: 2.1, w: 6.0, h: 3.2,
    barDir: "bar", barGrouping: "stacked",
    chartColors: [C.green, C.blue, C.red, "B0BEC5"],
    showLegend: true, legendPos: "b", legendFontSize: 10,
    catAxisLabelColor: C.text, catAxisLabelFontSize: 11, catAxisLabelBold: true,
    valAxisLabelColor: C.muted, valAxisLabelFontSize: 9,
    valGridLine: { color: C.divider, size: 0.5 }, catGridLine: { style: "none" },
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    showTitle: true, title: "各專案任務分佈",
    titleFontSize: 13, titleColor: C.navy,
    shadow: makeShadow()
  });
}

// ── Slide 4~N: 每專案一頁 (壓縮版) ─────────────────────
data.projects.forEach((proj) => {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  const h = proj.health;
  const hColor = HEALTH_COLOR[h.level];
  slideHeader(s, proj.name, `${proj.owner} ｜ ${proj.startDate} ～ ${proj.targetDate}`);

  // ─── 上半段: 健康度大卡 (左) + 里程碑時間軸 (右)

  // 左：健康度卡
  s.addShape("rect", {
    x: 0.3, y: 0.8, w: 3.0, h: 2.05,
    fill: { color: C.white }, line: { color: C.lightBlue }, shadow: makeShadow()
  });
  s.addShape("rect", { x: 0.3, y: 0.8, w: 0.1, h: 2.05, fill: { color: hColor }, line: { color: hColor } });

  // 燈號 + 標籤
  s.addShape("ellipse", { x: 0.55, y: 0.95, w: 0.32, h: 0.32, fill: { color: hColor }, line: { color: hColor } });
  s.addText(HEALTH_LABEL[h.level], {
    x: 0.95, y: 0.95, w: 1.5, h: 0.32,
    fontSize: 16, bold: true, color: hColor, fontFace: "Calibri", valign: "middle"
  });

  // 完成度大字
  s.addText(`${h.donePct}%`, {
    x: 0.5, y: 1.35, w: 1.5, h: 0.6,
    fontSize: 36, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
  });
  s.addText("已完成", {
    x: 0.5, y: 1.95, w: 1.5, h: 0.22,
    fontSize: 9.5, color: C.muted, fontFace: "Calibri"
  });

  // 預估完工
  s.addText("預估完工", {
    x: 1.95, y: 1.35, w: 1.3, h: 0.22,
    fontSize: 9.5, color: C.muted, fontFace: "Calibri"
  });
  s.addText(h.forecastEnd || "—", {
    x: 1.95, y: 1.55, w: 1.3, h: 0.32,
    fontSize: 13, bold: true, color: C.text, fontFace: "Calibri"
  });
  let driftTxt = "準時", driftColor = C.muted;
  if (h.forecastDrift > 0)      { driftTxt = `落後 ${h.forecastDrift} 天`; driftColor = C.red; }
  else if (h.forecastDrift < 0) { driftTxt = `提前 ${-h.forecastDrift} 天`; driftColor = C.green; }
  s.addText(driftTxt, {
    x: 1.95, y: 1.88, w: 1.3, h: 0.22,
    fontSize: 10, bold: true, color: driftColor, fontFace: "Calibri"
  });

  // 進度條
  const pbarX = 0.5, pbarY = 2.35, pbarW = 2.65, pbarH = 0.18;
  progressBar(s, pbarX, pbarY, pbarW, pbarH, h.donePct, hColor);
  const epX = pbarX + pbarW * h.elapsedPct / 100;
  s.addShape("rect", { x: epX - 0.01, y: pbarY - 0.04, w: 0.02, h: pbarH + 0.08, fill: { color: C.text }, line: { color: C.text } });
  s.addText(`時間已過 ${h.elapsedPct}%`, {
    x: pbarX, y: pbarY + 0.25, w: pbarW, h: 0.2,
    fontSize: 8, color: C.muted, fontFace: "Calibri", align: "center"
  });

  // 右：里程碑時間軸
  s.addShape("rect", {
    x: 3.5, y: 0.8, w: 6.2, h: 2.05,
    fill: { color: C.white }, line: { color: C.lightBlue }, shadow: makeShadow()
  });
  s.addText("里程碑", {
    x: 3.65, y: 0.85, w: 5.9, h: 0.3,
    fontSize: 12, bold: true, color: C.navy, fontFace: "Calibri"
  });

  if (proj.milestones.length > 0) {
    const ms = proj.milestones;
    const allDates = ms.flatMap(m => [new Date(m.startDate), new Date(m.endDate)]);
    const minD = new Date(Math.min(...allDates));
    const maxD = new Date(Math.max(...allDates));
    const totalD = (maxD - minD) / 86400000 || 1;
    const TX = 3.65, TW = 5.9, TY = 1.2;

    // 軸線
    s.addShape("rect", { x: TX, y: TY + 0.3, w: TW, h: 0.02, fill: { color: C.divider }, line: { color: C.divider } });

    // 今日標記
    const today = new Date();
    if (today >= minD && today <= maxD) {
      const todayX = TX + (today - minD) / 86400000 / totalD * TW;
      s.addShape("rect", { x: todayX - 0.01, y: TY + 0.18, w: 0.02, h: 1.4, fill: { color: C.accent }, line: { color: C.accent } });
      s.addText("今日", { x: todayX - 0.25, y: TY + 1.55, w: 0.5, h: 0.18, fontSize: 7.5, color: C.accent, bold: true, fontFace: "Calibri", align: "center" });
    }

    // 里程碑長條
    ms.forEach((m, i) => {
      const x1 = TX + (new Date(m.startDate) - minD) / 86400000 / totalD * TW;
      const x2 = TX + (new Date(m.endDate)   - minD) / 86400000 / totalD * TW;
      const w  = Math.max(0.25, x2 - x1);
      const color = STATUS_COLOR[m.status] || C.blue;

      // 條
      s.addShape("roundRect", {
        x: x1, y: TY + 0.22, w, h: 0.18,
        fill: { color }, line: { color }, rectRadius: 0.04
      });
      // 名稱 (上方交錯顯示，避免重疊)
      const labelY = i % 2 === 0 ? TY + 0.45 : TY + 0.65;
      s.addText(m.name, {
        x: x1 - 0.3, y: labelY, w: w + 0.6, h: 0.22,
        fontSize: 8, bold: true, color: C.text, fontFace: "Calibri", align: "center"
      });
      s.addText(`${m.progress}%`, {
        x: x1 - 0.3, y: labelY + 0.2, w: w + 0.6, h: 0.18,
        fontSize: 7.5, color, fontFace: "Calibri", align: "center", bold: true
      });
    });
  } else {
    s.addText("（無里程碑資料）", {
      x: 3.65, y: 1.5, w: 5.9, h: 0.4,
      fontSize: 11, color: C.muted, italic: true, fontFace: "Calibri", align: "center"
    });
  }

  // ─── 下半段: 本期完成 (左) + 下期重點 (中) + Top 風險 (右)
  const PANEL_Y = 2.95, PANEL_H = 2.55;

  function panel(x, w, title, color) {
    s.addShape("rect", { x, y: PANEL_Y, w, h: PANEL_H, fill: { color: C.white }, line: { color: C.lightBlue }, shadow: makeShadow() });
    s.addShape("rect", { x, y: PANEL_Y, w, h: 0.32, fill: { color }, line: { color } });
    s.addText(title, {
      x: x + 0.15, y: PANEL_Y, w: w - 0.3, h: 0.32,
      fontSize: 11, bold: true, color: C.white, fontFace: "Calibri", valign: "middle", margin: 0
    });
  }

  function listItems(x, w, items, formatter, emptyMsg) {
    const startY = PANEL_Y + 0.42;
    if (items.length === 0) {
      s.addText(emptyMsg, {
        x: x + 0.15, y: startY + 0.5, w: w - 0.3, h: 0.4,
        fontSize: 10, italic: true, color: C.muted, fontFace: "Calibri", align: "center"
      });
      return;
    }
    items.slice(0, 5).forEach((item, i) => {
      const y = startY + i * 0.4;
      formatter(s, x, y, w, item);
    });
  }

  // 左：本期完成
  panel(0.3, 3.13, `近 14 天完成 (${proj.activity.recentDone.length})`, C.green);
  listItems(0.3, 3.13, proj.activity.recentDone, (s, x, y, w, item) => {
    s.addShape("ellipse", { x: x + 0.15, y: y + 0.1, w: 0.1, h: 0.1, fill: { color: C.green }, line: { color: C.green } });
    s.addText(item.name, {
      x: x + 0.32, y, w: w - 0.4, h: 0.22,
      fontSize: 9.5, color: C.text, fontFace: "Calibri", valign: "middle"
    });
    s.addText(`${item.owner} · ${item.date}`, {
      x: x + 0.32, y: y + 0.2, w: w - 0.4, h: 0.16,
      fontSize: 8, color: C.muted, fontFace: "Calibri"
    });
  }, "本期無完成任務");

  // 中：進行中 / 即將到期
  panel(3.5, 3.13, `進行中 / 即將到期 (${proj.activity.upcoming.length})`, C.blue);
  listItems(3.5, 3.13, proj.activity.upcoming, (s, x, y, w, item) => {
    const stColor = STATUS_COLOR[item.status] || C.blue;
    s.addShape("ellipse", { x: x + 0.15, y: y + 0.1, w: 0.1, h: 0.1, fill: { color: stColor }, line: { color: stColor } });
    s.addText(item.name, {
      x: x + 0.32, y, w: w - 0.4, h: 0.22,
      fontSize: 9.5, color: C.text, fontFace: "Calibri", valign: "middle"
    });
    s.addText(`${item.owner} · ${item.status} · 截止 ${item.date}`, {
      x: x + 0.32, y: y + 0.2, w: w - 0.4, h: 0.16,
      fontSize: 8, color: C.muted, fontFace: "Calibri"
    });
  }, "目前無進行中任務");

  // 右：Top 風險
  const projDecisions = (data.decisions || []).filter(d => d.project === proj.name);
  panel(6.7, 3.0, `風險議題 (${projDecisions.length})`, C.red);
  listItems(6.7, 3.0, projDecisions, (s, x, y, w, item) => {
    const sevColor = item.severity === "高" ? C.red : C.amber;
    s.addShape("ellipse", { x: x + 0.15, y: y + 0.1, w: 0.1, h: 0.1, fill: { color: sevColor }, line: { color: sevColor } });
    s.addText(item.issue, {
      x: x + 0.32, y, w: w - 0.4, h: 0.22,
      fontSize: 9, color: C.text, fontFace: "Calibri", valign: "middle", bold: true
    });
    s.addText(item.ask, {
      x: x + 0.32, y: y + 0.2, w: w - 0.4, h: 0.16,
      fontSize: 7.5, color: sevColor, fontFace: "Calibri", italic: true
    });
  }, "目前無風險議題");
});

// ── Slide N+1: 議題與決策需求 ──────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  slideHeader(s, "需主管裁示事項", `按嚴重度排序 ｜ ${data.decisions.length} 件`);

  if (!data.decisions || data.decisions.length === 0) {
    s.addShape("rect", {
      x: 2, y: 1.8, w: 6, h: 1.4,
      fill: { color: C.white }, line: { color: C.lightBlue }, shadow: makeShadow()
    });
    s.addText("✓  目前無需裁示事項", {
      x: 2, y: 1.8, w: 6, h: 1.4,
      fontSize: 18, bold: true, color: C.green, fontFace: "Calibri",
      align: "center", valign: "middle"
    });
  } else {
    // 表頭
    const HX = 0.3, HY = 0.85, HW = 9.4;
    const COLW = [0.6, 1.4, 2.6, 2.4, 2.4];   // 嚴重度|專案|議題|影響|建議
    let cx = HX;
    ["", "專案", "議題", "影響評估", "建議行動"].forEach((h, i) => {
      s.addShape("rect", { x: cx, y: HY, w: COLW[i], h: 0.36, fill: { color: C.navy }, line: { color: C.navy } });
      s.addText(h, {
        x: cx, y: HY, w: COLW[i], h: 0.36,
        fontSize: 10.5, bold: true, color: C.white, fontFace: "Calibri",
        align: "center", valign: "middle", margin: 0
      });
      cx += COLW[i];
    });

    // 每筆決策
    const RH = 0.75;
    data.decisions.slice(0, 6).forEach((d, i) => {
      const y = HY + 0.36 + i * RH;
      const sevColor = d.severity === "高" ? C.red : (d.severity === "中" ? C.amber : C.muted);
      const rowBg = i % 2 === 0 ? C.white : C.paleBlue;
      s.addShape("rect", { x: HX, y, w: HW, h: RH, fill: { color: rowBg }, line: { color: C.divider, pt: 0.5 } });

      let cx = HX;
      // 嚴重度標
      s.addShape("rect", { x: cx + 0.1, y: y + 0.15, w: COLW[0] - 0.2, h: RH - 0.3, fill: { color: sevColor }, line: { color: sevColor } });
      s.addText(d.severity, {
        x: cx + 0.1, y: y + 0.15, w: COLW[0] - 0.2, h: RH - 0.3,
        fontSize: 14, bold: true, color: C.white, fontFace: "Calibri",
        align: "center", valign: "middle", margin: 0
      });
      cx += COLW[0];

      // 專案
      s.addText(d.project, {
        x: cx + 0.08, y: y + 0.05, w: COLW[1] - 0.16, h: 0.32,
        fontSize: 10, bold: true, color: C.text, fontFace: "Calibri", valign: "middle"
      });
      s.addText(`負責人 ${d.owner}`, {
        x: cx + 0.08, y: y + 0.4, w: COLW[1] - 0.16, h: 0.28,
        fontSize: 8.5, color: C.muted, fontFace: "Calibri"
      });
      cx += COLW[1];

      // 議題
      s.addText(d.issue, {
        x: cx + 0.1, y: y + 0.05, w: COLW[2] - 0.2, h: RH - 0.1,
        fontSize: 10, color: C.text, fontFace: "Calibri", valign: "middle"
      });
      cx += COLW[2];

      // 影響
      s.addText(d.impact, {
        x: cx + 0.1, y: y + 0.05, w: COLW[3] - 0.2, h: RH - 0.1,
        fontSize: 9.5, color: C.muted, fontFace: "Calibri", valign: "middle", italic: true
      });
      cx += COLW[3];

      // 建議
      s.addText(d.ask, {
        x: cx + 0.1, y: y + 0.05, w: COLW[4] - 0.2, h: RH - 0.1,
        fontSize: 10, bold: true, color: sevColor, fontFace: "Calibri", valign: "middle"
      });
    });
  }
}

// ── 結尾頁 ──────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape("rect", { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });
  s.addShape("rect", { x: 0.18, y: 0, w: 0.08, h: 5.625, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText("謝謝聆聽", {
    x: 1, y: 1.7, w: 8, h: 1.0,
    fontSize: 40, bold: true, color: C.white, fontFace: "Calibri", align: "center", charSpacing: 8
  });
  s.addText(`本期報告日期：${data.reportDate}`, {
    x: 1, y: 2.95, w: 8, h: 0.4,
    fontSize: 13, color: "CADCFC", fontFace: "Calibri", align: "center"
  });
  s.addText(`下次回報日期：${data.nextReport}`, {
    x: 1, y: 3.4, w: 8, h: 0.4,
    fontSize: 13, color: "CADCFC", fontFace: "Calibri", align: "center"
  });
}

// ── 輸出 ────────────────────────────────────────────────
pres.writeFile({ fileName: outputPath }).then(() => {
  console.log(`PPT 產生完成：${outputPath}`);
}).catch(e => {
  console.error("失敗：", e);
  process.exit(1);
});