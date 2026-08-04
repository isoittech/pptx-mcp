import { readFile } from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const pptxgen = require("pptxgenjs");

const [specificationPath, outputPath] = process.argv.slice(2);
if (!specificationPath || !outputPath) {
  throw new Error("Usage: node index.mjs <specification.json> <output.pptx>");
}

const spec = JSON.parse(await readFile(specificationPath, "utf8"));
if (!Array.isArray(spec.slides) || spec.slides.length < 1 || spec.slides.length > 50) {
  throw new Error("The visual deck must contain between 1 and 50 slides.");
}

const presets = {
  midnight: {
    primary: "14213D",
    secondary: "2D5BFF",
    accent: "2ED3C6",
    background: "F5F7FB",
    text: "17213A",
    muted: "667085",
  },
  aurora: {
    primary: "312E81",
    secondary: "7C3AED",
    accent: "14B8A6",
    background: "F7F5FF",
    text: "201A3D",
    muted: "706A86",
  },
  sunset: {
    primary: "3A1F3D",
    secondary: "F05A47",
    accent: "F5B642",
    background: "FFF8F3",
    text: "2B1B2E",
    muted: "7C687C",
  },
  forest: {
    primary: "153B32",
    secondary: "207A63",
    accent: "A7C957",
    background: "F5F8F4",
    text: "18332C",
    muted: "64746F",
  },
  minimal: {
    primary: "202124",
    secondary: "4F6BED",
    accent: "FFB000",
    background: "F7F7F5",
    text: "202124",
    muted: "6B6F76",
  },
  ocean: {
    primary: "063970",
    secondary: "087CA7",
    accent: "17C3B2",
    background: "F2F8FB",
    text: "102A43",
    muted: "627D98",
  },
  berry: {
    primary: "4A1942",
    secondary: "893168",
    accent: "E8C547",
    background: "FBF7FA",
    text: "2E172B",
    muted: "806579",
  },
  clay: {
    primary: "5E3023",
    secondary: "C08552",
    accent: "F3A712",
    background: "FCF8F2",
    text: "35251E",
    muted: "806E65",
  },
  cyber: {
    primary: "0B1020",
    secondary: "4F46E5",
    accent: "22D3EE",
    background: "F3F6FF",
    text: "111827",
    muted: "667085",
  },
};

const themeInput = spec.theme ?? {};
const presetName = String(themeInput.preset ?? "midnight").toLowerCase();
const baseTheme = presets[presetName];
if (!baseTheme) {
  throw new Error(`Unsupported theme preset: ${presetName}`);
}

function normalizeHexColor(value, fallback) {
  return String(value ?? fallback).replace(/^#/, "");
}

const theme = {
  ...baseTheme,
  primary: normalizeHexColor(themeInput.primaryColor, baseTheme.primary),
  secondary: normalizeHexColor(themeInput.secondaryColor, baseTheme.secondary),
  accent: normalizeHexColor(themeInput.accentColor, baseTheme.accent),
  background: normalizeHexColor(themeInput.backgroundColor, baseTheme.background),
  text: normalizeHexColor(themeInput.textColor, baseTheme.text),
  font: themeInput.fontFace ?? "Noto Sans CJK JP",
};
if (contrastRatio(theme.text, theme.background) < 4.5) {
  theme.text = luminance(theme.background) > 0.5 ? "17213A" : "FFFFFF";
}
if (contrastRatio(theme.muted, theme.background) < 3) {
  theme.muted = luminance(theme.background) > 0.5 ? "667085" : "CBD5E1";
}
theme.onPrimary = luminance(theme.primary) > 0.56 ? "17213A" : "FFFFFF";
theme.onSecondary = luminance(theme.secondary) > 0.56 ? "17213A" : "FFFFFF";

const designInput = spec.design ?? {};
const design = {
  style: String(designInput.style ?? "executive").toLowerCase(),
  density: String(designInput.density ?? "balanced").toLowerCase(),
  motif: String(designInput.motif ?? "geometric").toLowerCase(),
};
const densityScale = design.density === "airy" ? 1.05 : design.density === "detailed" ? 0.93 : 1;

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "pptx-mcp";
pptx.company = "MSI";
pptx.subject = spec.subject ?? "";
pptx.title = spec.title;
pptx.lang = spec.language ?? "ja-JP";
pptx.theme = {
  headFontFace: theme.font,
  bodyFontFace: theme.font,
  lang: spec.language ?? "ja-JP",
};
pptx.defineLayout({ name: "PPTX_MCP_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "PPTX_MCP_WIDE";

const W = 13.333;
const H = 7.5;
const totalSlides = spec.slides.length;

for (const [index, slideSpec] of spec.slides.entries()) {
  const slide = pptx.addSlide();
  const kind = String(slideSpec.kind).toLowerCase();
  switch (kind) {
    case "title":
      renderTitle(slide, slideSpec, index);
      break;
    case "agenda":
      renderAgenda(slide, slideSpec, index);
      break;
    case "section":
      renderSection(slide, slideSpec, index);
      break;
    case "bullets":
      renderBullets(slide, slideSpec, index);
      break;
    case "metrics":
      renderMetrics(slide, slideSpec, index);
      break;
    case "comparison":
      renderComparison(slide, slideSpec, index);
      break;
    case "process":
      renderProcess(slide, slideSpec, index);
      break;
    case "timeline":
      renderTimeline(slide, slideSpec, index);
      break;
    case "chart":
      renderChart(slide, slideSpec, index);
      break;
    case "statement":
      renderStatement(slide, slideSpec, index);
      break;
    case "cards":
      renderCards(slide, slideSpec, index);
      break;
    case "matrix":
      renderMatrix(slide, slideSpec, index);
      break;
    case "funnel":
      renderFunnel(slide, slideSpec, index);
      break;
    case "roadmap":
      renderRoadmap(slide, slideSpec, index);
      break;
    case "dashboard":
      renderDashboard(slide, slideSpec, index);
      break;
    case "quote":
      renderQuote(slide, slideSpec, index);
      break;
    case "closing":
      renderClosing(slide, slideSpec, index);
      break;
    default:
      throw new Error(`Unsupported slide kind: ${kind}`);
  }
}

await pptx.writeFile({ fileName: path.resolve(outputPath) });

function renderTitle(slide, data, index) {
  background(slide, theme.primary);
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 9.2, y: -1.15, w: 5.1, h: 5.1,
    fill: { color: theme.secondary, transparency: 8 },
    line: { color: theme.secondary, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 10.6, y: 2.15, w: 3.5, h: 3.5,
    fill: { color: theme.accent, transparency: 15 },
    line: { color: theme.accent, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.arc, {
    x: 8.55, y: 1.15, w: 3.7, h: 3.7,
    adjustPoint: 0.25,
    rotate: 18,
    fill: { color: theme.onPrimary, transparency: 100 },
    line: { color: theme.onPrimary, transparency: 72, width: 2.2 },
  });
  if (data.eyebrow) {
    slide.addText(data.eyebrow.toUpperCase(), {
      x: 0.75, y: 0.72, w: 7.7, h: 0.34,
      fontFace: theme.font, fontSize: 11, bold: true,
      charSpacing: 1.8, color: theme.accent, margin: 0,
    });
  }
  slide.addText(data.title, {
    x: 0.75, y: 1.34, w: 8.4, h: 2.2,
    fontFace: theme.font, fontSize: titleSize(data.title, 34, 25), bold: true,
    color: theme.onPrimary, margin: 0, breakLine: false,
    valign: "mid", fit: "shrink",
  });
  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: 0.75, y: 3.96, w: 7.7, h: 1.15,
      fontFace: theme.font, fontSize: 18, color: theme.onPrimary,
      transparency: 12, margin: 0, valign: "top", fit: "shrink",
    });
  }
  if (data.body) {
    pill(slide, data.body, 0.75, 6.18, Math.min(7.6, 1.7 + data.body.length * 0.14), theme.onPrimary, theme.primary);
  }
  footer(slide, index, true);
}

function renderAgenda(slide, data, index) {
  contentBase(slide, data, index);
  const items = data.bullets ?? [];
  const rows = Math.ceil(items.length / 2);
  const cardHeight = Math.min(1.02, 4.45 / rows - 0.16);
  items.forEach((item, itemIndex) => {
    const column = itemIndex % 2;
    const row = Math.floor(itemIndex / 2);
    const x = 0.7 + column * 6.12;
    const y = 2.05 + row * (cardHeight + 0.16);
    card(slide, x, y, 5.82, cardHeight, "FFFFFF");
    badge(slide, String(itemIndex + 1).padStart(2, "0"), x + 0.22, y + 0.2, 0.62, itemIndex % 3 === 2 ? theme.accent : theme.secondary);
    slide.addText(item, {
      x: x + 1.0, y: y + 0.15, w: 4.5, h: cardHeight - 0.3,
      fontFace: theme.font, fontSize: 16, bold: true, color: theme.text,
      margin: 0, valign: "mid", fit: "shrink",
    });
  });
}

function renderSection(slide, data, index) {
  background(slide, theme.background);
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 4.35, h: H,
    fill: { color: theme.primary }, line: { color: theme.primary },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 4.35, y: 0, w: 0.16, h: H,
    fill: { color: theme.accent }, line: { color: theme.accent },
  });
  slide.addText(String(index + 1).padStart(2, "0"), {
    x: 0.7, y: 0.92, w: 2.7, h: 1.35,
    fontFace: theme.font, fontSize: 60, bold: true,
    color: theme.onPrimary, transparency: 58, margin: 0,
  });
  if (data.eyebrow) {
    slide.addText(data.eyebrow.toUpperCase(), {
      x: 4.95, y: 1.34, w: 6.8, h: 0.34,
      fontFace: theme.font, fontSize: 11, bold: true,
      charSpacing: 1.6, color: theme.secondary, margin: 0,
    });
  }
  slide.addText(data.title, {
    x: 4.95, y: 1.9, w: 7.4, h: 1.6,
    fontFace: theme.font, fontSize: titleSize(data.title, 31, 24), bold: true,
    color: theme.text, margin: 0, valign: "mid", fit: "shrink",
  });
  if (data.subtitle || data.body) {
    slide.addText(data.subtitle ?? data.body, {
      x: 4.95, y: 3.82, w: 6.9, h: 1.25,
      fontFace: theme.font, fontSize: 17, color: theme.muted,
      margin: 0, valign: "top", fit: "shrink",
    });
  }
  slide.addShape(pptx.ShapeType.line, {
    x: 4.95, y: 5.5, w: 5.85, h: 0,
    line: { color: theme.secondary, transparency: 70, width: 1.3 },
  });
  footer(slide, index, false);
}

function renderBullets(slide, data, index) {
  contentBase(slide, data, index);
  const hasTakeaway = Boolean(data.takeaway);
  const listWidth = hasTakeaway ? 8.5 : 11.93;
  const items = data.bullets ?? [];
  if (String(data.variant ?? "auto").toLowerCase() === "split" && !hasTakeaway && items.length >= 4) {
    const midpoint = Math.ceil(items.length / 2);
    const columns = [items.slice(0, midpoint), items.slice(midpoint)];
    columns.forEach((columnItems, columnIndex) => {
      const x = 0.7 + columnIndex * 6.08;
      card(slide, x, 2.03, 5.85, 4.48, "FFFFFF");
      addNativeBulletList(slide, columnItems, x + 0.36, 2.37, 5.13, 3.84, 16.5);
    });
    return;
  }

  card(slide, 0.7, 2.03, listWidth, 4.48, "FFFFFF");
  slide.addText(items.map((item, itemIndex) => ({
    text: item,
    options: {
      bullet: true,
      breakLine: itemIndex < items.length - 1,
      paraSpaceAfter: 13,
      color: theme.text,
      fontSize: scaled(17),
    },
  })), {
    x: 1.05, y: 2.37, w: listWidth - 0.7, h: 3.84,
    fontFace: theme.font, color: theme.text,
    margin: 0.05, breakLine: false, valign: "mid", fit: "shrink",
  });
  if (data.takeaway) {
    takeawayCard(slide, data.takeaway, 9.55, 2.08, 3.08, 3.75);
  }
}

function renderMetrics(slide, data, index) {
  contentBase(slide, data, index);
  const metrics = data.metrics ?? [];
  if (String(data.variant ?? "auto").toLowerCase() === "spotlight" && metrics.length >= 3) {
    metricCard(slide, metrics[0], 0.7, 2.05, 5.18, 4.38, 0);
    const remaining = metrics.slice(1);
    const gap = 0.2;
    const height = (4.38 - gap * (remaining.length - 1)) / remaining.length;
    remaining.forEach((metric, metricIndex) => {
      smallMetricCard(slide, metric, 6.1, 2.05 + metricIndex * (height + gap), 6.53, height, metricIndex + 1);
    });
    return;
  }

  const columns = metrics.length === 3 ? 3 : 2;
  const rows = Math.ceil(metrics.length / columns);
  const gap = 0.24;
  const cardWidth = (11.93 - gap * (columns - 1)) / columns;
  const cardHeight = rows === 1 ? 3.8 : 2.05;
  metrics.forEach((metric, metricIndex) => {
    const column = metricIndex % columns;
    const row = Math.floor(metricIndex / columns);
    const x = 0.7 + column * (cardWidth + gap);
    const y = 2.05 + row * (cardHeight + 0.24);
    metricCard(slide, metric, x, y, cardWidth, cardHeight, metricIndex);
  });
  if (data.takeaway) {
    slide.addText(data.takeaway, {
      x: 0.78, y: 6.48, w: 11.7, h: 0.34,
      fontFace: theme.font, fontSize: 12.5, bold: true,
      color: theme.secondary, margin: 0, align: "center", fit: "shrink",
    });
  }
}

function renderComparison(slide, data, index) {
  contentBase(slide, data, index);
  const panels = data.panels ?? [];
  const gap = 0.22;
  const panelWidth = (11.93 - gap * (panels.length - 1)) / panels.length;
  panels.forEach((panel, panelIndex) => {
    const x = 0.7 + panelIndex * (panelWidth + gap);
    card(slide, x, 2.02, panelWidth, 4.55, "FFFFFF");
    const panelColor = panelIndex === 0 ? theme.secondary : panelIndex === 1 ? theme.accent : theme.primary;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + panelWidth / 2 - 0.27, y: 2.22, w: 0.54, h: 0.54,
      fill: { color: panelColor, transparency: 12 },
      line: { color: panelColor, transparency: 100 },
    });
    slide.addText(panel.title, {
      x: x + 0.26, y: 2.9, w: panelWidth - 0.52, h: 0.55,
      fontFace: theme.font, fontSize: 19, bold: true, color: theme.text,
      margin: 0, align: "center", valign: "mid", fit: "shrink",
    });
    if (panel.highlight) {
      pill(slide, panel.highlight, x + 0.35, 3.48, panelWidth - 0.7, panelColor, "FFFFFF");
    }
    const bullets = panel.bullets ?? [];
    const startY = panel.highlight ? 4.08 : 3.65;
    const itemHeight = Math.min(0.58, (6.25 - startY) / bullets.length);
    bullets.forEach((item, bulletIndex) => {
      slide.addShape(pptx.ShapeType.ellipse, {
        x: x + 0.35, y: startY + bulletIndex * itemHeight + 0.15, w: 0.13, h: 0.13,
        fill: { color: theme.accent }, line: { color: theme.accent },
      });
      slide.addText(item, {
        x: x + 0.62, y: startY + bulletIndex * itemHeight,
        w: panelWidth - 0.92, h: itemHeight,
        fontFace: theme.font, fontSize: 13.5, color: theme.text,
        margin: 0, valign: "mid", fit: "shrink",
      });
    });
  });
}

function renderProcess(slide, data, index) {
  contentBase(slide, data, index);
  const steps = data.steps ?? [];
  if (steps.length <= 4) {
    const width = (11.55 - 0.32 * (steps.length - 1)) / steps.length;
    steps.forEach((step, stepIndex) => {
      const x = 0.86 + stepIndex * (width + 0.32);
      processCard(slide, step, stepIndex, x, 2.43, width, 3.48);
      if (stepIndex < steps.length - 1) {
        slide.addShape(pptx.ShapeType.chevron, {
          x: x + width + 0.05, y: 3.83, w: 0.22, h: 0.55,
          fill: { color: theme.secondary, transparency: 40 },
          line: { color: theme.secondary, transparency: 100 },
        });
      }
    });
  } else {
    steps.forEach((step, stepIndex) => {
      const column = stepIndex % 3;
      const row = Math.floor(stepIndex / 3);
      const x = 0.72 + column * 4.06;
      const y = 2.0 + row * 2.28;
      processCard(slide, step, stepIndex, x, y, 3.72, 1.98);
    });
  }
  if (data.takeaway) {
    slide.addText(data.takeaway, {
      x: 1.1, y: 6.45, w: 11.1, h: 0.32,
      fontFace: theme.font, fontSize: 12, bold: true,
      color: theme.secondary, margin: 0, align: "center", fit: "shrink",
    });
  }
}

function renderTimeline(slide, data, index) {
  contentBase(slide, data, index);
  const steps = data.steps ?? [];
  const startX = 1.08;
  const endX = 12.2;
  const spacing = (endX - startX) / (steps.length - 1);
  slide.addShape(pptx.ShapeType.line, {
    x: startX, y: 4.05, w: endX - startX, h: 0,
    line: { color: theme.secondary, transparency: 38, width: 3.5 },
  });
  steps.forEach((step, stepIndex) => {
    const x = startX + stepIndex * spacing;
    const above = stepIndex % 2 === 0;
    const textWidth = Math.min(2.36, spacing * 0.88);
    const textX = Math.max(0.72, Math.min(x - textWidth / 2, W - 0.72 - textWidth));
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x - 0.19, y: 3.86, w: 0.38, h: 0.38,
      fill: { color: stepIndex === steps.length - 1 ? theme.accent : theme.secondary },
      line: { color: "FFFFFF", width: 2 },
    });
    slide.addText(step.label ?? String(stepIndex + 1).padStart(2, "0"), {
      x: x - 0.5, y: above ? 3.48 : 4.31, w: 1, h: 0.3,
      fontFace: theme.font, fontSize: 10.5, bold: true,
      color: theme.secondary, margin: 0, align: "center", fit: "shrink",
    });
    slide.addText(step.title, {
      x: textX,
      y: above ? 2.4 : 4.75,
      w: textWidth, h: 0.48,
      fontFace: theme.font, fontSize: 14, bold: true,
      color: theme.text, margin: 0, align: "center", valign: "mid", fit: "shrink",
    });
    if (step.description) {
      slide.addText(step.description, {
        x: textX,
        y: above ? 2.98 : 5.29,
        w: textWidth, h: 0.42,
        fontFace: theme.font, fontSize: 10.5, color: theme.muted,
        margin: 0, align: "center", valign: above ? "bottom" : "top", fit: "shrink",
      });
    }
  });
}

function renderChart(slide, data, index) {
  contentBase(slide, data, index);
  const chart = data.chart;
  const chartType = String(chart.kind).toLowerCase();
  const typeMap = {
    bar: pptx.ChartType.bar,
    line: pptx.ChartType.line,
    pie: pptx.ChartType.pie,
    doughnut: pptx.ChartType.doughnut,
  };
  const series = chart.series.map((item) => ({
    name: item.name,
    labels: chart.categories,
    values: item.values,
  }));
  const showSidebar = Boolean(data.takeaway || data.body);
  const chartWidth = showSidebar ? 8.65 : 11.9;
  card(slide, 0.7, 1.96, chartWidth, 4.78, "FFFFFF");
  const chartOptions = {
    x: 1.0, y: 2.23, w: chartWidth - 0.58, h: 4.18,
    showTitle: false,
    showLegend: chart.showLegend,
    legendPos: "b",
    legendFontFace: theme.font,
    legendFontSize: 10,
    chartColors: [theme.secondary, theme.accent, theme.primary, "8B9DC3"],
    showValue: chartType === "pie" || chartType === "doughnut",
    showPercent: chartType === "pie" || chartType === "doughnut",
    dataLabelPosition: "bestFit",
    catAxisLabelFontFace: theme.font,
    catAxisLabelFontSize: 10,
    catAxisLabelColor: theme.muted,
    valAxisLabelFontFace: theme.font,
    valAxisLabelFontSize: 10,
    valAxisLabelColor: theme.muted,
    valAxisLineColor: "D9DFE8",
    catAxisLineColor: "D9DFE8",
    valGridLine: { color: "D9DFE8", size: 1 },
    catGridLine: { style: "none" },
    showCatName: false,
    showSerName: false,
    showValue: chartType === "pie" || chartType === "doughnut",
    showPercent: chartType === "pie" || chartType === "doughnut",
    showLeaderLines: true,
    holeSize: chartType === "doughnut" ? 56 : undefined,
    dataLabelFormatCode: chart.valueSuffix ? `0${chart.valueSuffix}` : "0",
    showBorder: false,
  };
  slide.addChart(typeMap[chartType], series, chartOptions);
  if (showSidebar) {
    takeawayCard(slide, data.takeaway ?? data.body, 9.65, 2.03, 2.98, 4.55);
  }
}

function renderStatement(slide, data, index) {
  background(slide, theme.primary);
  addMotif(slide, index, true, false);
  if (data.eyebrow) {
    slide.addText(data.eyebrow.toUpperCase(), {
      x: 0.78, y: 0.65, w: 8.4, h: 0.32,
      fontFace: theme.font, fontSize: scaled(10.5), bold: true,
      charSpacing: 1.6, color: theme.accent, margin: 0,
    });
  }
  slide.addText(data.title, {
    x: 0.78, y: 1.12, w: 6.8, h: 0.68,
    fontFace: theme.font, fontSize: scaled(21), bold: true,
    color: theme.onPrimary, transparency: 18, margin: 0, fit: "shrink",
  });
  slide.addText(data.body, {
    x: 0.78, y: 2.0, w: 8.3, h: 3.3,
    fontFace: theme.font, fontSize: scaled(titleSize(data.body, 34, 23)), bold: true,
    color: theme.onPrimary, margin: 0, valign: "mid", fit: "shrink",
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 9.48, y: 1.36, w: 3.12, h: 4.82,
    rectRadius: 0.1,
    fill: { color: theme.onPrimary, transparency: 92 },
    line: { color: theme.onPrimary, transparency: 78, width: 1 },
  });
  renderIcon(slide, "insight", 10.45, 1.94, 1.16, theme.accent, true);
  slide.addText(data.takeaway ?? data.subtitle ?? "意思決定に必要な一文へ情報を圧縮する", {
    x: 9.86, y: 3.55, w: 2.36, h: 1.65,
    fontFace: theme.font, fontSize: scaled(16), bold: true,
    color: theme.onPrimary, margin: 0, align: "center", valign: "mid", fit: "shrink",
  });
  footer(slide, index, true);
}

function renderCards(slide, data, index) {
  contentBase(slide, data, index);
  const cards = data.cards ?? [];
  if (String(data.variant ?? "auto").toLowerCase() === "spotlight" && cards.length >= 3 && cards.length <= 4) {
    renderVisualCard(slide, cards[0], 0.7, 2.04, 4.58, 4.42, 0);
    const remaining = cards.slice(1);
    const gap = 0.2;
    const height = (4.42 - gap * (remaining.length - 1)) / remaining.length;
    remaining.forEach((item, itemIndex) => {
      renderVisualCard(slide, item, 5.5, 2.04 + itemIndex * (height + gap), 7.13, height, itemIndex + 1);
    });
    return;
  }

  const columns = cards.length === 4 ? 2 : 3;
  const rows = Math.ceil(cards.length / columns);
  const gap = 0.22;
  const width = (11.93 - gap * (columns - 1)) / columns;
  const height = rows === 1 ? 4.42 : 2.12;
  cards.forEach((item, itemIndex) => {
    const column = itemIndex % columns;
    const row = Math.floor(itemIndex / columns);
    const centeredLastRowOffset = cards.length === 5 && row === 1 ? (width + gap) / 2 : 0;
    const x = 0.7 + centeredLastRowOffset + column * (width + gap);
    const y = 2.04 + row * (height + 0.22);
    renderVisualCard(slide, item, x, y, width, height, itemIndex);
  });
}

function renderMatrix(slide, data, index) {
  contentBase(slide, data, index);
  const matrix = data.matrix;
  const quadrants = matrix.quadrants;
  const x = 1.36;
  const y = 2.16;
  const width = 10.74;
  const height = 4.18;
  const gap = 0.12;
  const cellWidth = (width - gap) / 2;
  const cellHeight = (height - gap) / 2;
  const colors = [theme.secondary, theme.accent, theme.primary, toneColor("positive", 0)];
  quadrants.forEach((quadrant, quadrantIndex) => {
    const column = quadrantIndex % 2;
    const row = Math.floor(quadrantIndex / 2);
    const cellX = x + column * (cellWidth + gap);
    const cellY = y + row * (cellHeight + gap);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: cellX, y: cellY, w: cellWidth, h: cellHeight,
      rectRadius: 0.06,
      fill: { color: colors[quadrantIndex], transparency: 87 },
      line: { color: colors[quadrantIndex], transparency: 58, width: 1 },
    });
    slide.addText(quadrant.title, {
      x: cellX + 0.24, y: cellY + 0.2, w: cellWidth - 0.48, h: 0.38,
      fontFace: theme.font, fontSize: scaled(15), bold: true,
      color: theme.text, margin: 0, fit: "shrink",
    });
    if (quadrant.highlight) {
      slide.addText(quadrant.highlight, {
        x: cellX + cellWidth - 1.25, y: cellY + 0.22, w: 0.98, h: 0.28,
        fontFace: theme.font, fontSize: scaled(9), bold: true,
        color: colors[quadrantIndex], margin: 0, align: "right", fit: "shrink",
      });
    }
    addNativeBulletList(slide, quadrant.bullets, cellX + 0.25, cellY + 0.72, cellWidth - 0.5, cellHeight - 0.9, 11.5);
  });
  slide.addText(matrix.horizontalAxis, {
    x: 4.1, y: 6.48, w: 5.25, h: 0.26,
    fontFace: theme.font, fontSize: scaled(9.5), bold: true,
    color: theme.muted, margin: 0, align: "center", fit: "shrink",
  });
  slide.addText(matrix.verticalAxis, {
    x: -0.42, y: 4.0, w: 2.4, h: 0.32,
    fontFace: theme.font, fontSize: scaled(9.5), bold: true,
    color: theme.muted, margin: 0, align: "center", valign: "mid", rotate: 270, fit: "shrink",
  });
  if (data.takeaway) {
    slide.addText(data.takeaway, {
      x: 1.7, y: 6.78, w: 9.95, h: 0.2,
      fontFace: theme.font, fontSize: scaled(9.5), bold: true,
      color: theme.secondary, margin: 0, align: "center", fit: "shrink",
    });
  }
}

function renderFunnel(slide, data, index) {
  contentBase(slide, data, index);
  const steps = data.steps ?? [];
  const maxWidth = 8.85;
  const minWidth = 4.2;
  const rowHeight = Math.min(0.82, 4.24 / steps.length);
  steps.forEach((step, stepIndex) => {
    const ratio = steps.length === 1 ? 0 : stepIndex / (steps.length - 1);
    const width = maxWidth - (maxWidth - minWidth) * ratio;
    const x = 0.78 + (maxWidth - width) / 2;
    const y = 2.08 + stepIndex * rowHeight;
    const color = stepIndex % 2 === 0 ? theme.secondary : theme.primary;
    slide.addShape(pptx.ShapeType.trapezoid, {
      x, y, w: width, h: rowHeight - 0.08,
      fill: { color, transparency: Math.min(12 + stepIndex * 8, 48) },
      line: { color, transparency: 100 },
    });
    slide.addText(`${step.label ?? String(stepIndex + 1).padStart(2, "0")}  ${step.title}`, {
      x: x + 0.32, y: y + 0.06, w: width - 0.64, h: rowHeight - 0.2,
      fontFace: theme.font, fontSize: scaled(13.5), bold: true,
      color: luminance(color) > 0.56 ? theme.text : "FFFFFF",
      margin: 0, align: "center", valign: "mid", fit: "shrink",
    });
  });
  const detail = steps.map((step) => step.description).filter(Boolean).join("\n");
  takeawayCard(slide, data.takeaway || detail || "段階ごとの転換率と阻害要因を確認", 9.92, 2.1, 2.72, 4.35);
}

function renderRoadmap(slide, data, index) {
  contentBase(slide, data, index);
  const steps = data.steps ?? [];
  const gap = 0.18;
  const width = (11.86 - gap * (steps.length - 1)) / steps.length;
  slide.addShape(pptx.ShapeType.rightArrow, {
    x: 0.86, y: 3.75, w: 11.54, h: 0.48,
    fill: { color: theme.secondary, transparency: 72 },
    line: { color: theme.secondary, transparency: 100 },
  });
  steps.forEach((step, stepIndex) => {
    const x = 0.72 + stepIndex * (width + gap);
    const above = stepIndex % 2 === 0;
    const y = above ? 2.12 : 4.52;
    card(slide, x, y, width, 1.42, "FFFFFF");
    badge(slide, step.label ?? String(stepIndex + 1), x + 0.18, y + 0.18, 0.5, stepIndex === steps.length - 1 ? theme.accent : theme.secondary);
    slide.addText(step.title, {
      x: x + 0.82, y: y + 0.18, w: width - 1.02, h: 0.38,
      fontFace: theme.font, fontSize: scaled(12), bold: true,
      color: theme.text, margin: 0, fit: "shrink",
    });
    if (step.description) {
      slide.addText(step.description, {
        x: x + 0.2, y: y + 0.74, w: width - 0.4, h: 0.44,
        fontFace: theme.font, fontSize: scaled(9.5),
        color: theme.muted, margin: 0, align: "center", fit: "shrink",
      });
    }
    slide.addShape(pptx.ShapeType.line, {
      x: x + width / 2, y: above ? 3.54 : 4.22, w: 0, h: above ? 0.21 : 0.3,
      line: { color: theme.secondary, transparency: 35, width: 1.5 },
    });
  });
  if (data.takeaway) {
    slide.addText(data.takeaway, {
      x: 1.1, y: 6.42, w: 11.1, h: 0.34,
      fontFace: theme.font, fontSize: scaled(11.5), bold: true,
      color: theme.secondary, margin: 0, align: "center", fit: "shrink",
    });
  }
}

function renderDashboard(slide, data, index) {
  contentBase(slide, data, index);
  const metrics = data.metrics ?? [];
  const metricGap = 0.18;
  const metricWidth = (11.93 - metricGap * (metrics.length - 1)) / metrics.length;
  metrics.forEach((metric, metricIndex) => {
    smallMetricCard(slide, metric, 0.7 + metricIndex * (metricWidth + metricGap), 2.0, metricWidth, 1.34, metricIndex);
  });
  const chart = data.chart;
  const chartType = String(chart.kind).toLowerCase();
  const typeMap = {
    bar: pptx.ChartType.bar,
    line: pptx.ChartType.line,
    pie: pptx.ChartType.pie,
    doughnut: pptx.ChartType.doughnut,
  };
  const series = chart.series.map((item) => ({
    name: item.name,
    labels: chart.categories,
    values: item.values,
  }));
  const hasTakeaway = Boolean(data.takeaway);
  const chartWidth = hasTakeaway ? 8.72 : 11.93;
  card(slide, 0.7, 3.58, chartWidth, 3.12, "FFFFFF");
  slide.addChart(typeMap[chartType], series, {
    x: 0.98, y: 3.82, w: chartWidth - 0.56, h: 2.58,
    showLegend: chart.showLegend,
    legendPos: "b",
    legendFontFace: theme.font,
    legendFontSize: 9,
    chartColors: [theme.secondary, theme.accent, theme.primary, "8B9DC3"],
    showValue: chartType === "pie" || chartType === "doughnut",
    showPercent: chartType === "pie" || chartType === "doughnut",
    dataLabelPosition: "bestFit",
    catAxisLabelFontFace: theme.font,
    catAxisLabelFontSize: 9,
    catAxisLabelColor: theme.muted,
    valAxisLabelFontFace: theme.font,
    valAxisLabelFontSize: 9,
    valAxisLabelColor: theme.muted,
    valAxisLineColor: "D9DFE8",
    catAxisLineColor: "D9DFE8",
    valGridLine: { color: "D9DFE8", size: 1 },
    catGridLine: { style: "none" },
    showBorder: false,
    holeSize: chartType === "doughnut" ? 56 : undefined,
  });
  if (data.takeaway) {
    takeawayCard(slide, data.takeaway, 9.65, 3.6, 2.98, 3.08);
  }
}

function renderQuote(slide, data, index) {
  background(slide, theme.background);
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.22, h: H,
    fill: { color: theme.accent }, line: { color: theme.accent },
  });
  slide.addText("“", {
    x: 0.78, y: 0.55, w: 2.1, h: 1.85,
    fontFace: "Georgia", fontSize: 104, bold: true,
    color: theme.secondary, transparency: 20, margin: 0,
  });
  if (data.eyebrow) {
    slide.addText(data.eyebrow.toUpperCase(), {
      x: 2.0, y: 0.88, w: 8.7, h: 0.32,
      fontFace: theme.font, fontSize: 11, bold: true,
      charSpacing: 1.6, color: theme.secondary, margin: 0,
    });
  }
  slide.addText(data.body, {
    x: 1.72, y: 1.63, w: 9.92, h: 3.52,
    fontFace: theme.font, fontSize: titleSize(data.body, 28, 20), bold: true,
    color: theme.text, margin: 0, valign: "mid", align: "center", fit: "shrink",
  });
  if (data.attribution) {
    slide.addText(`— ${data.attribution}`, {
      x: 3.1, y: 5.53, w: 7.12, h: 0.5,
      fontFace: theme.font, fontSize: 14, color: theme.muted,
      margin: 0, align: "center", fit: "shrink",
    });
  }
  footer(slide, index, false);
}

function renderClosing(slide, data, index) {
  const bullets = data.bullets ?? [];
  const hasBullets = bullets.length > 0;
  background(slide, theme.primary);
  slide.addShape(pptx.ShapeType.arc, {
    x: -1.35, y: 3.8, w: 4.7, h: 4.7,
    rotate: 12,
    fill: { color: theme.primary, transparency: 100 },
    line: { color: theme.accent, transparency: 24, width: 3 },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 10.7, y: -0.85, w: 3.45, h: 3.45,
    fill: { color: theme.secondary, transparency: 8 },
    line: { color: theme.secondary, transparency: 100 },
  });
  slide.addText(data.title, {
    x: 1.25, y: hasBullets ? 0.88 : 1.73, w: 10.83, h: hasBullets ? 1.08 : 1.45,
    fontFace: theme.font, fontSize: titleSize(data.title, hasBullets ? 31 : 35, 25), bold: true,
    color: theme.onPrimary, margin: 0, align: "center", valign: "mid", fit: "shrink",
  });
  if (data.subtitle || data.body) {
    slide.addText(data.subtitle ?? data.body, {
      x: 2.0, y: hasBullets ? 2.05 : 3.48, w: 9.33, h: hasBullets ? 0.72 : 1.2,
      fontFace: theme.font, fontSize: hasBullets ? 15 : 18, color: theme.onPrimary,
      transparency: 10, margin: 0, align: "center", valign: "mid", fit: "shrink",
    });
  }
  if (hasBullets) {
    slide.addText(bullets.slice(0, 4).map((item, itemIndex) => ({
      text: item,
      options: {
        bullet: true,
        breakLine: itemIndex < Math.min(bullets.length, 4) - 1,
        paraSpaceAfter: 10,
        color: theme.onPrimary,
        fontSize: scaled(15),
      },
    })), {
      x: 2.15, y: 3.05, w: 9.03, h: 2.12,
      fontFace: theme.font, color: theme.onPrimary,
      margin: 0.04, valign: "mid", fit: "shrink",
    });
  }
  if (data.takeaway) {
    pill(slide, data.takeaway, 3.55, hasBullets ? 5.62 : 5.34, 6.23, theme.accent, theme.primary);
  }
  footer(slide, index, true);
}

function contentBase(slide, data, index) {
  background(slide, theme.background);
  addMotif(slide, index, false, true);
  if (data.eyebrow) {
    pill(slide, data.eyebrow.toUpperCase(), 0.7, 0.28, Math.min(3.4, 0.9 + data.eyebrow.length * 0.115), theme.secondary, "FFFFFF");
  }
  slide.addText(data.title, {
    x: 0.7, y: data.eyebrow ? 0.84 : 0.46, w: 10.9, h: 0.7,
    fontFace: theme.font, fontSize: scaled(titleSize(data.title, 25, 20)), bold: true,
    color: theme.text, margin: 0, valign: "mid", fit: "shrink",
  });
  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: 0.7, y: 1.48, w: 11.55, h: 0.34,
      fontFace: theme.font, fontSize: scaled(12), color: theme.muted,
      margin: 0, fit: "shrink",
    });
  }
  slide.addText(String(index + 1).padStart(2, "0"), {
    x: 11.72, y: 0.42, w: 0.8, h: 0.42,
    fontFace: theme.font, fontSize: scaled(13), bold: true,
    color: theme.secondary, transparency: 20, margin: 0, align: "right",
  });
  footer(slide, index, false);
}

function background(slide, color) {
  slide.background = { color };
}

function footer(slide, index, inverse) {
  const color = inverse ? theme.onPrimary : theme.muted;
  slide.addText(spec.title, {
    x: 0.7, y: 7.12, w: 8.6, h: 0.18,
    fontFace: theme.font, fontSize: 7.5, color, transparency: 36, margin: 0,
  });
  slide.addText(`${index + 1} / ${totalSlides}`, {
    x: 11.6, y: 7.08, w: 1.03, h: 0.22,
    fontFace: theme.font, fontSize: 8.5, bold: true,
    color, transparency: 22, margin: 0, align: "right",
  });
}

function card(slide, x, y, w, h, color) {
  const shape = design.style === "technical" ? pptx.ShapeType.rect : pptx.ShapeType.roundRect;
  slide.addShape(shape, {
    x, y, w, h,
    rectRadius: 0.08,
    fill: { color },
    line: { color: "E1E6EE", width: 0.7 },
    shadow: design.style === "editorial"
      ? undefined
      : { type: "outer", color: "182230", opacity: design.style === "bold" ? 0.16 : 0.09, blur: 2, angle: 45, distance: 1 },
  });
}

function badge(slide, text, x, y, diameter, color) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x, y, w: diameter, h: diameter,
    fill: { color }, line: { color, transparency: 100 },
  });
  slide.addText(text, {
    x, y: y + 0.01, w: diameter, h: diameter - 0.02,
    fontFace: theme.font, fontSize: 10.5, bold: true,
    color: luminance(color) > 0.56 ? theme.text : "FFFFFF",
    margin: 0, align: "center", valign: "mid", fit: "shrink",
  });
}

function pill(slide, text, x, y, w, fillColor, textColor) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h: 0.52,
    rectRadius: 0.12,
    fill: { color: fillColor, transparency: 4 },
    line: { color: fillColor, transparency: 100 },
  });
  slide.addText(text, {
    x: x + 0.18, y: y + 0.03, w: w - 0.36, h: 0.44,
    fontFace: theme.font, fontSize: 11, bold: true, color: textColor,
    margin: 0, align: "center", valign: "mid", fit: "shrink",
  });
}

function takeawayCard(slide, text, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.08,
    fill: { color: theme.primary },
    line: { color: theme.primary },
    shadow: { type: "outer", color: "182230", opacity: 0.16, blur: 2.5, angle: 45, distance: 1.2 },
  });
  slide.addText("KEY TAKEAWAY", {
    x: x + 0.3, y: y + 0.38, w: w - 0.6, h: 0.28,
    fontFace: theme.font, fontSize: 9, bold: true,
    charSpacing: 1.2, color: theme.accent, margin: 0,
  });
  slide.addText(text, {
    x: x + 0.3, y: y + 0.93, w: w - 0.6, h: h - 1.35,
    fontFace: theme.font, fontSize: 18, bold: true,
    color: theme.onPrimary, margin: 0, valign: "mid", fit: "shrink",
  });
}

function metricCard(slide, metric, x, y, w, h, index) {
  const tone = toneColor(metric.tone, index);
  card(slide, x, y, w, h, "FFFFFF");
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + w - 0.78, y: y + 0.26, w: 0.42, h: 0.42,
    fill: { color: tone, transparency: 75 }, line: { color: tone, transparency: 100 },
  });
  slide.addText(metric.value, {
    x: x + 0.38, y: y + 0.34, w: w - 0.82, h: h > 2.5 ? 1.25 : 0.68,
    fontFace: theme.font, fontSize: h > 2.5 ? 36 : 27, bold: true,
    color: tone, margin: 0, valign: "mid", fit: "shrink",
  });
  slide.addText(metric.label, {
    x: x + 0.4, y: y + (h > 2.5 ? 1.72 : 1.07), w: w - 0.8, h: 0.48,
    fontFace: theme.font, fontSize: 15.5, bold: true,
    color: theme.text, margin: 0, fit: "shrink",
  });
  if (metric.detail) {
    slide.addText(metric.detail, {
      x: x + 0.4, y: y + (h > 2.5 ? 2.42 : 1.52), w: w - 0.8, h: h > 2.5 ? 0.72 : 0.32,
      fontFace: theme.font, fontSize: 11.5, color: theme.muted,
      margin: 0, fit: "shrink",
    });
  }
  slide.addText(String(index + 1).padStart(2, "0"), {
    x: x + w - 0.72, y: y + h - 0.44, w: 0.36, h: 0.22,
    fontFace: theme.font, fontSize: 8, bold: true,
    color: theme.muted, transparency: 25, margin: 0, align: "right",
  });
}

function processCard(slide, step, index, x, y, w, h) {
  card(slide, x, y, w, h, "FFFFFF");
  const label = step.label ?? String(index + 1).padStart(2, "0");
  const stepColor = index % 2 === 0 ? theme.secondary : theme.accent;
  if ([...label].length > 4) {
    pill(slide, label, x + 0.28, y + 0.28, Math.min(1.3, w - 0.56), stepColor, "FFFFFF");
  } else {
    badge(slide, label, x + 0.28, y + 0.28, 0.68, stepColor);
  }
  slide.addText(step.title, {
    x: x + 0.28, y: y + 1.15, w: w - 0.56, h: h > 2.5 ? 0.62 : 0.42,
    fontFace: theme.font, fontSize: h > 2.5 ? ([...step.title].length > 8 ? 14.5 : 17) : 14.5, bold: true,
    color: theme.text, margin: 0, fit: "shrink",
  });
  if (step.description) {
    slide.addText(step.description, {
      x: x + 0.28, y: y + (h > 2.5 ? 1.95 : 1.5), w: w - 0.56, h: h > 2.5 ? 1.08 : 0.3,
      fontFace: theme.font, fontSize: h > 2.5 ? 12.5 : 10.5,
      color: theme.muted, margin: 0, valign: "top", fit: "shrink",
    });
  }
}

function smallMetricCard(slide, metric, x, y, w, h, index) {
  const tone = toneColor(metric.tone, index);
  card(slide, x, y, w, h, "FFFFFF");
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + 0.22, y: y + 0.24, w: 0.42, h: 0.42,
    fill: { color: tone, transparency: 72 },
    line: { color: tone, transparency: 100 },
  });
  slide.addText(metric.value, {
    x: x + 0.78, y: y + 0.18, w: w - 1.02, h: 0.5,
    fontFace: theme.font, fontSize: scaled(21), bold: true,
    color: tone, margin: 0, fit: "shrink",
  });
  slide.addText(metric.label, {
    x: x + 0.24, y: y + 0.78, w: w - 0.48, h: 0.3,
    fontFace: theme.font, fontSize: scaled(10.5), bold: true,
    color: theme.text, margin: 0, fit: "shrink",
  });
  if (metric.detail) {
    slide.addText(metric.detail, {
      x: x + 0.24, y: y + 1.08, w: w - 0.48, h: 0.18,
      fontFace: theme.font, fontSize: scaled(7.5), color: theme.muted,
      margin: 0, fit: "shrink",
    });
  }
}

function renderVisualCard(slide, item, x, y, w, h, index) {
  const tone = toneColor(item.tone, index);
  const spacious = h > 2.5;
  const compact = h <= 1.6;
  card(slide, x, y, w, h, "FFFFFF");
  renderIcon(slide, item.icon, x + 0.28, y + (compact ? 0.34 : 0.28), compact ? 0.62 : spacious ? 0.88 : 0.66, tone, false);
  if (item.value) {
    slide.addText(item.value, {
      x: x + w - Math.min(2.1, w * 0.42), y: y + 0.29, w: Math.min(1.72, w * 0.36), h: 0.57,
      fontFace: theme.font, fontSize: scaled(spacious ? 22 : 17), bold: true,
      color: tone, margin: 0, align: "right", valign: "mid", fit: "shrink",
    });
  }
  if (compact) {
    slide.addText(item.title, {
      x: x + 1.12, y: y + 0.2, w: w - (item.value ? 3.0 : 1.42), h: 0.38,
      fontFace: theme.font, fontSize: scaled(14), bold: true,
      color: theme.text, margin: 0, fit: "shrink",
    });
    if (item.description) {
      slide.addText(item.description, {
        x: x + 1.12, y: y + 0.68, w: w - 1.42, h: h - 0.84,
        fontFace: theme.font, fontSize: scaled(9.5), color: theme.muted,
        margin: 0, fit: "shrink",
      });
    }
    return;
  }

  slide.addText(item.title, {
    x: x + 0.3, y: y + (spacious ? 1.47 : 1.04), w: w - 0.6, h: spacious ? 0.62 : 0.42,
    fontFace: theme.font, fontSize: scaled(spacious ? 19 : 15), bold: true,
    color: theme.text, margin: 0, fit: "shrink",
  });
  if (item.description) {
    slide.addText(item.description, {
      x: x + 0.3, y: y + (spacious ? 2.28 : 1.52), w: w - 0.6, h: spacious ? h - 2.65 : Math.max(0.3, h - 1.7),
      fontFace: theme.font, fontSize: scaled(spacious ? 13 : 10.5),
      color: theme.muted, margin: 0, valign: "top", fit: "shrink",
    });
  }
}

function addNativeBulletList(slide, items, x, y, w, h, fontSize) {
  slide.addText((items ?? []).map((item, itemIndex) => ({
    text: item,
    options: {
      bullet: true,
      breakLine: itemIndex < items.length - 1,
      paraSpaceAfter: 7,
      fontSize: scaled(fontSize),
      color: theme.text,
    },
  })), {
    x, y, w, h,
    fontFace: theme.font, color: theme.text,
    margin: 0.03, valign: "mid", fit: "shrink",
  });
}

function renderIcon(slide, iconName, x, y, size, color, inverse) {
  const normalizedIcon = String(iconName ?? "insight").toLowerCase();
  const shapeMap = {
    insight: pptx.ShapeType.sun,
    target: pptx.ShapeType.donut,
    growth: pptx.ShapeType.chartPlus,
    people: pptx.ShapeType.hexagon,
    shield: pptx.ShapeType.pentagon,
    clock: pptx.ShapeType.pie,
    cloud: pptx.ShapeType.cloud,
    settings: pptx.ShapeType.gear6,
    data: pptx.ShapeType.can,
    warning: pptx.ShapeType.triangle,
    check: pptx.ShapeType.star4,
    idea: pptx.ShapeType.lightningBolt,
    compliance: pptx.ShapeType.flowChartDocument,
    decision: pptx.ShapeType.flowChartDecision,
    lock: pptx.ShapeType.pentagon,
    network: pptx.ShapeType.flowChartConnector,
    document: pptx.ShapeType.actionButtonDocument,
    communication: pptx.ShapeType.wedgeRoundRectCallout,
    recovery: pptx.ShapeType.actionButtonReturn,
    backup: pptx.ShapeType.flowChartOfflineStorage,
    legal: pptx.ShapeType.flowChartDocument,
    monitor: pptx.ShapeType.flowChartDisplay,
    automation: pptx.ShapeType.gear6,
  };
  const shape = shapeMap[normalizedIcon] ?? pptx.ShapeType.sun;
  const foreground = inverse ? theme.onPrimary : color;
  slide.addShape(pptx.ShapeType.ellipse, {
    x, y, w: size, h: size,
    fill: { color, transparency: inverse ? 68 : 84 },
    line: { color, transparency: 100 },
  });
  if (normalizedIcon === "search") {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + size * 0.24, y: y + size * 0.2, w: size * 0.4, h: size * 0.4,
      fill: { color: foreground, transparency: 100 },
      line: { color: foreground, transparency: inverse ? 2 : 5, width: 1.8 },
    });
    slide.addShape(pptx.ShapeType.line, {
      x: x + size * 0.58, y: y + size * 0.57, w: size * 0.2, h: size * 0.2,
      line: { color: foreground, transparency: inverse ? 2 : 5, width: 2.2 },
    });
    return;
  }
  slide.addShape(shape, {
    x: x + size * 0.23, y: y + size * 0.23, w: size * 0.54, h: size * 0.54,
    fill: { color: foreground, transparency: inverse ? 2 : 5 },
    line: { color: foreground, transparency: 100 },
  });
}

function addMotif(slide, index, inverse, subtle) {
  if (design.motif === "none") return;
  const foreground = inverse ? theme.onPrimary : theme.secondary;
  const transparency = subtle ? 90 : 72;
  const offset = (index % 3) * 0.16;
  if (design.motif === "orbit") {
    slide.addShape(pptx.ShapeType.arc, {
      x: 10.52 - offset, y: -0.78, w: 3.18, h: 3.18,
      rotate: 18 + index * 7,
      fill: { color: foreground, transparency: 100 },
      line: { color: foreground, transparency, width: 1.4 },
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 11.83, y: 0.42 + offset, w: 0.24, h: 0.24,
      fill: { color: theme.accent, transparency: subtle ? 52 : 18 },
      line: { color: theme.accent, transparency: 100 },
    });
    return;
  }

  if (design.motif === "nodes") {
    const points = [[10.88, 0.34], [11.72, 0.72], [12.34, 0.22]];
    slide.addShape(pptx.ShapeType.line, {
      x: points[0][0], y: points[0][1], w: points[1][0] - points[0][0], h: points[1][1] - points[0][1],
      line: { color: foreground, transparency, width: 1 },
    });
    slide.addShape(pptx.ShapeType.line, {
      x: points[1][0], y: points[1][1], w: points[2][0] - points[1][0], h: points[2][1] - points[1][1],
      line: { color: foreground, transparency, width: 1 },
    });
    points.forEach(([px, py], pointIndex) => slide.addShape(pptx.ShapeType.ellipse, {
      x: px, y: py, w: 0.16 + pointIndex * 0.03, h: 0.16 + pointIndex * 0.03,
      fill: { color: pointIndex === 1 ? theme.accent : foreground, transparency: subtle ? 52 : 18 },
      line: { color: foreground, transparency: 100 },
    }));
    return;
  }

  if (design.motif === "ribbon") {
    slide.addShape(pptx.ShapeType.wave, {
      x: 9.98, y: -0.18, w: 3.8, h: 1.18,
      rotate: -8,
      fill: { color: theme.accent, transparency },
      line: { color: theme.accent, transparency: 100 },
    });
    return;
  }

  slide.addShape(pptx.ShapeType.ellipse, {
    x: 11.48 - offset, y: -0.55, w: 2.38, h: 2.38,
    fill: { color: foreground, transparency },
    line: { color: foreground, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.triangle, {
    x: 10.9, y: 0.22 + offset, w: 0.72, h: 0.72,
    rotate: 18 + index * 6,
    fill: { color: theme.accent, transparency: subtle ? 76 : 40 },
    line: { color: theme.accent, transparency: 100 },
  });
}

function toneColor(tone, index) {
  const normalizedTone = String(tone ?? "accent").replace(/^#/, "").toLowerCase();
  if (/^[0-9a-f]{6}$/.test(normalizedTone)) return normalizedTone.toUpperCase();
  const colors = {
    accent: theme.secondary,
    primary: theme.primary,
    secondary: theme.accent,
    info: theme.secondary,
    positive: "159D73",
    success: "159D73",
    warning: "E89222",
    critical: "D92D20",
    danger: "D92D20",
    negative: "D92D20",
    risk: "D92D20",
    neutral: theme.muted,
    muted: theme.muted,
  };
  const requested = colors[normalizedTone];
  if (requested) return requested;
  return index % 2 === 0 ? theme.secondary : theme.accent;
}

function scaled(value) {
  return Number((value * densityScale).toFixed(2));
}

function titleSize(text, maximum, minimum) {
  const length = [...String(text)].length;
  if (length <= 24) return maximum;
  if (length >= 70) return minimum;
  return maximum - ((length - 24) / 46) * (maximum - minimum);
}

function luminance(hex) {
  const components = [0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16) / 255)
    .map((value) => value <= 0.03928 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4);
  return 0.2126 * components[0] + 0.7152 * components[1] + 0.0722 * components[2];
}

function contrastRatio(first, second) {
  const lighter = Math.max(luminance(first), luminance(second));
  const darker = Math.min(luminance(first), luminance(second));
  return (lighter + 0.05) / (darker + 0.05);
}
