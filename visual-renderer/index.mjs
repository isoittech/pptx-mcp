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
theme.onPrimary = luminance(theme.primary) > 0.56 ? "17213A" : "FFFFFF";
theme.onSecondary = luminance(theme.secondary) > 0.56 ? "17213A" : "FFFFFF";

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
  slide.addShape(pptx.ShapeType.line, {
    x: 0.75, y: 3.92, w: 1.2, h: 0,
    line: { color: theme.accent, width: 5, beginArrowType: "none", endArrowType: "none" },
  });
  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: 0.75, y: 4.18, w: 7.7, h: 1.15,
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
  const rowHeight = Math.min(0.75, 4.62 / items.length);
  items.forEach((item, itemIndex) => {
    const y = 1.98 + itemIndex * rowHeight;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 0.74, y: y + 0.12, w: 0.34, h: 0.34,
      fill: { color: itemIndex % 2 === 0 ? theme.secondary : theme.accent },
      line: { color: itemIndex % 2 === 0 ? theme.secondary : theme.accent },
    });
    slide.addText(item, {
      x: 1.28, y, w: listWidth - 1.1, h: rowHeight - 0.04,
      fontFace: theme.font, fontSize: 16.5, color: theme.text,
      margin: 0, valign: "mid", fit: "shrink",
    });
    if (itemIndex < items.length - 1) {
      slide.addShape(pptx.ShapeType.line, {
        x: 1.28, y: y + rowHeight - 0.03, w: listWidth - 1.18, h: 0,
        line: { color: "DCE2EC", width: 0.6 },
      });
    }
  });
  if (data.takeaway) {
    takeawayCard(slide, data.takeaway, 9.55, 2.08, 3.08, 3.75);
  }
}

function renderMetrics(slide, data, index) {
  contentBase(slide, data, index);
  const metrics = data.metrics ?? [];
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
    slide.addShape(pptx.ShapeType.rect, {
      x, y: 2.02, w: panelWidth, h: 0.13,
      fill: { color: panelIndex === 0 ? theme.secondary : panelIndex === 1 ? theme.accent : theme.primary },
      line: { color: panelIndex === 0 ? theme.secondary : panelIndex === 1 ? theme.accent : theme.primary },
    });
    slide.addText(panel.title, {
      x: x + 0.26, y: 2.34, w: panelWidth - 0.52, h: 0.55,
      fontFace: theme.font, fontSize: 19, bold: true, color: theme.text,
      margin: 0, align: "center", valign: "mid", fit: "shrink",
    });
    if (panel.highlight) {
      pill(slide, panel.highlight, x + 0.35, 3.02, panelWidth - 0.7, theme.secondary, "FFFFFF");
    }
    const bullets = panel.bullets ?? [];
    const startY = panel.highlight ? 3.67 : 3.18;
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
    x: 1.25, y: 1.73, w: 10.83, h: 1.45,
    fontFace: theme.font, fontSize: titleSize(data.title, 35, 27), bold: true,
    color: theme.onPrimary, margin: 0, align: "center", valign: "mid", fit: "shrink",
  });
  if (data.subtitle || data.body) {
    slide.addText(data.subtitle ?? data.body, {
      x: 2.0, y: 3.48, w: 9.33, h: 1.2,
      fontFace: theme.font, fontSize: 18, color: theme.onPrimary,
      transparency: 10, margin: 0, align: "center", valign: "mid", fit: "shrink",
    });
  }
  if (data.takeaway) {
    pill(slide, data.takeaway, 3.55, 5.34, 6.23, theme.accent, theme.primary);
  }
  footer(slide, index, true);
}

function contentBase(slide, data, index) {
  background(slide, theme.background);
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.16, h: 1.76,
    fill: { color: theme.accent }, line: { color: theme.accent },
  });
  if (data.eyebrow) {
    slide.addText(data.eyebrow.toUpperCase(), {
      x: 0.7, y: 0.35, w: 8.5, h: 0.26,
      fontFace: theme.font, fontSize: 9.5, bold: true,
      charSpacing: 1.4, color: theme.secondary, margin: 0,
    });
  }
  slide.addText(data.title, {
    x: 0.7, y: data.eyebrow ? 0.7 : 0.46, w: 10.9, h: 0.7,
    fontFace: theme.font, fontSize: titleSize(data.title, 25, 20), bold: true,
    color: theme.text, margin: 0, valign: "mid", fit: "shrink",
  });
  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: 0.7, y: 1.43, w: 11.55, h: 0.34,
      fontFace: theme.font, fontSize: 12, color: theme.muted,
      margin: 0, fit: "shrink",
    });
  }
  slide.addShape(pptx.ShapeType.line, {
    x: 0.7, y: 1.84, w: 11.93, h: 0,
    line: { color: "D9DFE8", width: 0.7 },
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
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.08,
    fill: { color },
    line: { color: "E1E6EE", width: 0.7 },
    shadow: { type: "outer", color: "182230", opacity: 0.11, blur: 2, angle: 45, distance: 1 },
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
  const toneColors = {
    accent: theme.secondary,
    positive: "159D73",
    warning: "E89222",
    neutral: theme.muted,
  };
  const tone = toneColors[String(metric.tone ?? "accent").toLowerCase()] ?? theme.secondary;
  card(slide, x, y, w, h, "FFFFFF");
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w: 0.12, h,
    fill: { color: tone }, line: { color: tone },
  });
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
  badge(slide, step.label ?? String(index + 1).padStart(2, "0"), x + 0.28, y + 0.28, 0.68, index % 2 === 0 ? theme.secondary : theme.accent);
  slide.addText(step.title, {
    x: x + 0.28, y: y + 1.15, w: w - 0.56, h: h > 2.5 ? 0.62 : 0.42,
    fontFace: theme.font, fontSize: h > 2.5 ? 17 : 14.5, bold: true,
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
