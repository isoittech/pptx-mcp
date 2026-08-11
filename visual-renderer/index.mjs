import { readFile } from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import {
  addBravuraGlyph,
  loadBravuraFont,
  MusicGlyph,
  timeSignatureGlyph,
} from "./music-glyphs.mjs";

const require = createRequire(import.meta.url);
const pptxgen = require("pptxgenjs");

const [specificationPath, outputPath] = process.argv.slice(2);
if (!specificationPath || !outputPath) {
  throw new Error("Usage: node index.mjs <specification.json> <output.pptx>");
}

const spec = JSON.parse(await readFile(specificationPath, "utf8"));
const rendererContract = String(spec.rendererContract ?? "visual-v4").toLowerCase();
if (!["visual-v4", "visual-v5"].includes(rendererContract)) {
  throw new Error(`Unsupported renderer contract: ${rendererContract}`);
}
const usesModernRendererContract = rendererContract === "visual-v5";
const bravuraFont = await loadBravuraFont();
const templateChrome = spec.templateChrome === true;
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

const legacyFontFace = themeInput.fontFace ?? "Noto Sans CJK JP";
const resolvedBackground = normalizeHexColor(themeInput.backgroundColor, baseTheme.background);
const defaultSurface = luminance(resolvedBackground) > 0.5
  ? "FFFFFF"
  : mixHex(resolvedBackground, "FFFFFF", 0.08);
const theme = {
  ...baseTheme,
  primary: normalizeHexColor(themeInput.primaryColor, baseTheme.primary),
  secondary: normalizeHexColor(themeInput.secondaryColor, baseTheme.secondary),
  accent: normalizeHexColor(themeInput.accentColor, baseTheme.accent),
  background: resolvedBackground,
  surface: normalizeHexColor(themeInput.surfaceColor, defaultSurface),
  text: normalizeHexColor(themeInput.textColor, baseTheme.text),
  muted: normalizeHexColor(themeInput.mutedTextColor, baseTheme.muted),
  positive: normalizeHexColor(themeInput.positiveColor, "159D73"),
  warning: normalizeHexColor(themeInput.warningColor, "E89222"),
  critical: normalizeHexColor(themeInput.criticalColor, "D92D20"),
  headingFont: themeInput.headingFontFace ?? legacyFontFace,
  bodyFont: themeInput.bodyFontFace ?? legacyFontFace,
};
if (contrastRatio(theme.text, theme.background) < 4.5) {
  theme.text = luminance(theme.background) > 0.5 ? "17213A" : "FFFFFF";
}
if (contrastRatio(theme.muted, theme.background) < 3) {
  theme.muted = luminance(theme.background) > 0.5 ? "667085" : "CBD5E1";
}
theme.onSurface = !usesModernRendererContract || contrastRatio(theme.text, theme.surface) >= 4.5
  ? theme.text
  : luminance(theme.surface) > 0.5 ? "17213A" : "FFFFFF";
theme.mutedOnSurface = !usesModernRendererContract || contrastRatio(theme.muted, theme.surface) >= 3
  ? theme.muted
  : luminance(theme.surface) > 0.5 ? "667085" : "CBD5E1";
theme.dataSeries = usesModernRendererContract
  && Array.isArray(themeInput.dataSeriesColors)
  && themeInput.dataSeriesColors.length > 0
  ? themeInput.dataSeriesColors.map((color) => normalizeHexColor(color, theme.secondary))
  : [theme.secondary, theme.accent, theme.primary, "8B9DC3"];
theme.onPrimary = usesModernRendererContract
  ? readableForeground(theme.primary)
  : luminance(theme.primary) > 0.56 ? "17213A" : "FFFFFF";
theme.onSecondary = usesModernRendererContract
  ? readableForeground(theme.secondary)
  : luminance(theme.secondary) > 0.56 ? "17213A" : "FFFFFF";
theme.onAccent = readableForeground(theme.accent);
theme.secondaryTextOnBackground = readableToneOnBackground(theme.secondary, theme.background, theme.text);
theme.accentTextOnPrimary = readableToneOnBackground(theme.accent, theme.primary, theme.onPrimary);

const designInput = spec.design ?? {};
const design = {
  style: String(designInput.style ?? "executive").toLowerCase(),
  density: String(designInput.density ?? "balanced").toLowerCase(),
  motif: String(designInput.motif ?? "geometric").toLowerCase(),
};
const styleProfiles = {
  executive: {
    titleScale: 1.08,
    surface: theme.surface,
    cardShape: "round",
    cardShadow: false,
    borderColor: mixHex(theme.surface, theme.onSurface, 0.22),
    borderWidthScale: 1.15,
  },
  editorial: {
    titleScale: 1,
    surface: theme.surface,
    cardShape: "rect",
    cardShadow: false,
    borderColor: theme.accent,
    borderWidthScale: 0.7,
  },
  bold: {
    titleScale: 1.08,
    surface: theme.surface,
    cardShape: "round",
    cardShadow: true,
    borderColor: theme.secondary,
    borderWidthScale: 1.25,
  },
  technical: {
    titleScale: 0.96,
    surface: mixHex(theme.surface, theme.background, 0.18),
    cardShape: "rect",
    cardShadow: false,
    borderColor: mixHex(theme.surface, theme.onSurface, 0.26),
    borderWidthScale: 0.75,
  },
  playful: {
    titleScale: 1.03,
    surface: theme.surface,
    cardShape: "round",
    cardShadow: true,
    borderColor: theme.accent,
    borderWidthScale: 1,
  },
};
const selectedStyleProfile = styleProfiles[design.style];
if (!selectedStyleProfile) {
  throw new Error(`Unsupported design style: ${design.style}`);
}
const styleProfile = usesModernRendererContract
  ? selectedStyleProfile
  : {
      ...selectedStyleProfile,
      titleScale: 1,
      surface: "FFFFFF",
      borderColor: "D9E0E8",
      borderWidthScale: 1,
    };
const cardTextColor = !usesModernRendererContract
  ? theme.text
  : contrastRatio(theme.text, styleProfile.surface) >= 4.5
    ? theme.text
    : readableForeground(styleProfile.surface);
const preferredCardMutedTextColor = luminance(styleProfile.surface) > 0.5 ? "667085" : "CBD5E1";
const cardMutedTextColor = !usesModernRendererContract
  ? theme.muted
  : contrastRatio(theme.muted, styleProfile.surface) >= 3
    ? theme.muted
    : contrastRatio(preferredCardMutedTextColor, styleProfile.surface) >= 3
      ? preferredCardMutedTextColor
      : readableForeground(styleProfile.surface);
const densityProfiles = {
  airy: {
    fontScale: 1.05,
    outerX: 0.82,
    contentTop: 2.12,
    contentBottom: 6.5,
    gap: 0.28,
    panelPadding: 0.38,
    titleY: 0.52,
    titleHeight: 0.74,
    subtitleY: 1.54,
    cardBorderWidth: 0.65,
    cardShadow: true,
  },
  balanced: {
    fontScale: 1,
    outerX: 0.7,
    contentTop: 2.02,
    contentBottom: 6.62,
    gap: 0.22,
    panelPadding: 0.3,
    titleY: 0.46,
    titleHeight: 0.7,
    subtitleY: 1.48,
    cardBorderWidth: 0.7,
    cardShadow: true,
  },
  detailed: {
    // Detailed is a distinct information-design system, not a global font shrink.
    // It uses more of the canvas, tighter rhythm, thin rules, and no card shadows.
    fontScale: 0.98,
    outerX: 0.52,
    contentTop: 1.62,
    contentBottom: 6.9,
    gap: 0.12,
    panelPadding: 0.22,
    titleY: 0.31,
    titleHeight: 0.6,
    subtitleY: 1.18,
    cardBorderWidth: 0.5,
    cardShadow: false,
  },
};
const deckDensityName = design.density;
let currentDensityName = deckDensityName;
let density = densityProfiles[currentDensityName] ?? densityProfiles.balanced;
let densityScale = density.fontScale;

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "pptx-mcp";
pptx.company = "pptx-mcp";
pptx.subject = spec.subject ?? "";
pptx.title = spec.title;
pptx.lang = spec.language ?? "ja-JP";
pptx.theme = {
  headFontFace: theme.headingFont,
  bodyFontFace: theme.bodyFont,
  lang: spec.language ?? "ja-JP",
};
pptx.defineLayout({ name: "PPTX_MCP_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "PPTX_MCP_WIDE";

const W = 13.333;
const H = 7.5;
const totalSlides = spec.slides.length;

for (const [index, slideSpec] of spec.slides.entries()) {
  currentDensityName = String(slideSpec.density ?? deckDensityName).toLowerCase();
  density = densityProfiles[currentDensityName];
  if (!density) {
    throw new Error(`Unsupported slide density: ${currentDensityName}`);
  }
  densityScale = density.fontScale;
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
    case "structuredbrief":
    case "structured_brief":
      renderStructuredBrief(slide, slideSpec, index);
      break;
    case "scorecard":
      renderScorecard(slide, slideSpec, index);
      break;
    case "datatable":
    case "data_table":
      renderDataTable(slide, slideSpec, index);
      break;
    case "musicscore":
    case "music_score":
      renderMusicScore(slide, slideSpec, index);
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
  addDarkSlideStyle(slide, index);
  if (data.eyebrow) {
    slide.addText(data.eyebrow.toUpperCase(), {
      x: 0.75, y: 0.72, w: 7.7, h: 0.34,
      fontFace: theme.bodyFont, fontSize: 11, bold: true,
      charSpacing: 1.8, color: theme.accentTextOnPrimary, margin: 0,
    });
  }
  slide.addText(data.title, {
    x: 0.75, y: 1.34, w: 8.4, h: 2.2,
    fontFace: theme.headingFont,
    fontSize: titleSize(data.title, 34, 25) * styleProfile.titleScale,
    bold: true,
    color: theme.onPrimary, margin: 0, breakLine: false,
    valign: "mid", fit: "shrink",
  });
  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: 0.75, y: 3.96, w: 7.7, h: 1.15,
      fontFace: theme.bodyFont, fontSize: 18, color: theme.onPrimary,
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
      fontFace: theme.headingFont, fontSize: 16, bold: true, color: cardTextColor,
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
  addSectionStyle(slide, index);
  slide.addText(String(index + 1).padStart(2, "0"), {
    x: 0.7, y: 0.92, w: 2.7, h: 1.35,
    fontFace: theme.bodyFont, fontSize: 60, bold: true,
    color: theme.onPrimary, transparency: 58, margin: 0,
  });
  if (data.eyebrow) {
    slide.addText(data.eyebrow.toUpperCase(), {
      x: 4.95, y: 1.34, w: 6.8, h: 0.34,
      fontFace: theme.bodyFont, fontSize: 11, bold: true,
      charSpacing: 1.6, color: theme.secondaryTextOnBackground, margin: 0,
    });
  }
  slide.addText(data.title, {
    x: 4.95, y: 1.9, w: 7.4, h: 1.6,
    fontFace: theme.headingFont,
    fontSize: titleSize(data.title, 31, 24) * styleProfile.titleScale,
    bold: true,
    color: theme.text, margin: 0, valign: "mid", fit: "shrink",
  });
  if (data.subtitle || data.body) {
    slide.addText(data.subtitle ?? data.body, {
      x: 4.95, y: 3.82, w: 6.9, h: 1.25,
      fontFace: theme.bodyFont, fontSize: 17, color: theme.muted,
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
      addNativeBulletList(slide, columnItems, x + 0.36, 2.37, 5.13, 3.84, 16.5, "mid", cardTextColor);
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
      color: cardTextColor,
      fontSize: scaled(17),
    },
  })), {
    x: 1.05, y: 2.37, w: listWidth - 0.7, h: 3.84,
    fontFace: theme.bodyFont, color: cardTextColor,
    margin: 0.05, breakLine: false, valign: "mid", fit: "shrink",
  });
  if (data.takeaway) {
    takeawayCard(slide, data.takeaway, 9.55, 2.08, 3.08, 3.75);
  }
}

function renderMetrics(slide, data, index) {
  contentBase(slide, data, index);
  const metrics = data.metrics ?? [];
  if (String(data.variant ?? "auto").toLowerCase() === "spotlight" && metrics.length === 3) {
    metricCard(slide, metrics[0], 0.7, 2.05, 5.18, 4.38, 0);
    const remaining = metrics.slice(1);
    const gap = 0.2;
    const height = (4.38 - gap * (remaining.length - 1)) / remaining.length;
    remaining.forEach((metric, metricIndex) => {
      smallMetricCard(slide, metric, 6.1, 2.05 + metricIndex * (height + gap), 6.53, height, metricIndex + 1);
    });
    return;
  }

  // Five or six metrics use a compact 3x2 grid. This is a common model output
  // for risk-signal dashboards and remains readable without discarding items.
  const columns = metrics.length === 3 || metrics.length >= 5 ? 3 : 2;
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
      fontFace: theme.bodyFont, fontSize: 12.5, bold: true,
      color: theme.secondaryTextOnBackground, margin: 0, align: "center", fit: "shrink",
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
      fontFace: theme.headingFont, fontSize: 19, bold: true, color: cardTextColor,
      margin: 0, align: "center", valign: "mid", fit: "shrink",
    });
    if (panel.highlight) {
      pill(
        slide,
        panel.highlight,
        x + 0.35,
        3.48,
        panelWidth - 0.7,
        panelColor,
        usesModernRendererContract ? readableForeground(panelColor) : "FFFFFF",
      );
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
        fontFace: theme.bodyFont, fontSize: 13.5, color: cardTextColor,
        margin: 0, valign: "mid", fit: "shrink",
      });
    });
  });
}

function renderStructuredBrief(slide, data, index) {
  contentBase(slide, data, index);
  const sections = data.sections ?? [];
  const x = density.outerX;
  const y = density.contentTop;
  const width = W - density.outerX * 2;
  const height = density.contentBottom - density.contentTop;
  const totalCharacters = sections.reduce((sum, section) => sum + briefSectionCharacterCount(section), 0);
  const variant = String(data.variant ?? "auto").toLowerCase();
  const useMosaic = sections.length === 3
    && (variant === "editorial" || (variant === "auto" && totalCharacters < 600));

  if (useMosaic) {
    const leadHeight = 1.82;
    renderBriefLeadPanel(slide, sections[0], x, y, width, leadHeight, 0);
    const lowerY = y + leadHeight + density.gap;
    const lowerHeight = height - leadHeight - density.gap;
    const lowerWidth = (width - density.gap) / 2;
    renderBriefPanel(slide, sections[1], x, lowerY, lowerWidth, lowerHeight, 1);
    renderBriefPanel(slide, sections[2], x + lowerWidth + density.gap, lowerY, lowerWidth, lowerHeight, 2);
  } else {
    const panelWidth = (width - density.gap * (sections.length - 1)) / sections.length;
    sections.forEach((section, sectionIndex) => {
      renderBriefPanel(
        slide,
        section,
        x + sectionIndex * (panelWidth + density.gap),
        y,
        panelWidth,
        height,
        sectionIndex,
      );
    });
  }

  if (data.takeaway) {
    slide.addText(data.takeaway, {
      x: density.outerX,
      y: 6.93,
      w: W - density.outerX * 2,
      h: 0.22,
      fontFace: theme.bodyFont,
      fontSize: 9.5,
      bold: true,
      color: theme.secondaryTextOnBackground,
      margin: 0,
      align: "center",
      fit: "shrink",
    });
  }
}

function renderBriefLeadPanel(slide, section, x, y, width, height, sectionIndex) {
  const tone = toneColor(section.tone, sectionIndex);
  const padding = density.panelPadding;
  const bullets = section.bullets ?? [];
  const hasBullets = bullets.length > 0;
  const headingWidth = 2.45;
  card(slide, x, y, width, height, "FFFFFF");
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w: 0.08, h: height,
    fill: { color: tone },
    line: { color: tone, transparency: 100 },
  });
  slide.addText(section.heading, {
    x: x + padding + 0.08, y: y + 0.25, w: headingWidth - padding, h: 0.48,
    fontFace: theme.headingFont, fontSize: 15.5, bold: true,
    color: cardTextColor, margin: 0, valign: "mid", fit: "shrink",
  });
  if (section.highlight) {
    slide.addText(section.highlight, {
      x: x + padding + 0.08, y: y + 0.88, w: headingWidth - padding, h: 0.32,
      fontFace: theme.bodyFont, fontSize: 10.5, bold: true,
      color: readableToneOnSurface(tone), margin: 0, fit: "shrink",
    });
  }
  slide.addText(String(sectionIndex + 1).padStart(2, "0"), {
    x: x + padding + 0.08, y: y + height - 0.43, w: 0.5, h: 0.2,
    fontFace: theme.bodyFont, fontSize: 8.5, bold: true,
    color: readableToneOnSurface(tone), transparency: 20, margin: 0,
  });
  const dividerX = x + headingWidth + 0.15;
  slide.addShape(pptx.ShapeType.line, {
    x: dividerX, y: y + 0.22, w: 0, h: height - 0.44,
    line: { color: tone, transparency: 62, width: 0.75 },
  });
  const contentX = dividerX + 0.28;
  const contentWidth = width - (contentX - x) - padding;
  const bodyWidth = hasBullets ? contentWidth * 0.64 : contentWidth;
  if (section.body) {
    slide.addText(section.body, {
      x: contentX, y: y + 0.28, w: bodyWidth, h: height - 0.56,
      fontFace: theme.bodyFont, fontSize: briefBodyFontSize(section, 12.8),
      color: cardTextColor, margin: 0, valign: "mid", fit: "shrink",
    });
  }
  if (hasBullets) {
    const bulletX = contentX + bodyWidth + 0.3;
    slide.addShape(pptx.ShapeType.line, {
      x: bulletX - 0.16, y: y + 0.3, w: 0, h: height - 0.6,
      line: { color: styleProfile.borderColor, width: 0.55 },
    });
    addNativeBulletList(
      slide,
      bullets,
      bulletX,
      y + 0.3,
      x + width - padding - bulletX,
      height - 0.6,
      briefBulletFontSize(section, 11.2),
      "mid",
      cardTextColor,
    );
  }
}

function renderBriefPanel(slide, section, x, y, width, height, sectionIndex) {
  const tone = toneColor(section.tone, sectionIndex);
  const padding = density.panelPadding;
  const bullets = section.bullets ?? [];
  const sectionCharacters = briefSectionCharacterCount(section);
  const headingFontSize = sectionCharacters < 190 ? 15 : sectionCharacters < 270 ? 14 : 13;
  card(slide, x, y, width, height, "FFFFFF");
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w: width, h: currentDensityName === "detailed" ? 0.07 : 0.1,
    fill: { color: tone },
    line: { color: tone, transparency: 100 },
  });
  slide.addText(section.heading, {
    x: x + padding, y: y + 0.2, w: width - padding * 2 - 0.45, h: 0.52,
    fontFace: theme.headingFont, fontSize: headingFontSize, bold: true,
    color: cardTextColor, margin: 0, valign: "mid", fit: "shrink",
  });
  slide.addText(String(sectionIndex + 1).padStart(2, "0"), {
    x: x + width - padding - 0.36, y: y + 0.27, w: 0.36, h: 0.2,
    fontFace: theme.bodyFont, fontSize: 8.5, bold: true,
    color: readableToneOnSurface(tone), transparency: 20, margin: 0, align: "right",
  });
  const highlightY = y + 0.77;
  if (section.highlight) {
    slide.addText(section.highlight, {
      x: x + padding, y: highlightY, w: width - padding * 2, h: 0.32,
      fontFace: theme.bodyFont, fontSize: 10.5, bold: true,
      color: readableToneOnSurface(tone), margin: 0, fit: "shrink",
    });
  }
  const ruleY = section.highlight ? y + 1.15 : y + 0.88;
  slide.addShape(pptx.ShapeType.line, {
    x: x + padding, y: ruleY, w: width - padding * 2, h: 0,
    line: { color: tone, transparency: 58, width: 0.75 },
  });
  const textTop = ruleY + 0.16;
  const textBottom = y + height - padding;
  const availableHeight = textBottom - textTop;
  const hasBody = Boolean(section.body);
  const hasBullets = bullets.length > 0;
  const bodyFontSize = briefBodyFontSize(section, currentDensityName === "detailed" ? 11.2 : 11.8);
  const bodyCharactersPerLine = Math.max(16, Math.floor((width - padding * 2) * 7.2));
  const estimatedBodyLines = hasBody ? Math.ceil([...section.body].length / bodyCharactersPerLine) : 0;
  const estimatedBodyHeight = Math.max(0.7, estimatedBodyLines * bodyFontSize / 48 + 0.1);
  const bodyHeight = hasBody
    ? hasBullets ? Math.min(availableHeight - 0.68, estimatedBodyHeight) : availableHeight
    : 0;
  if (hasBody) {
    slide.addText(section.body, {
      x: x + padding, y: textTop, w: width - padding * 2, h: bodyHeight,
      fontFace: theme.bodyFont, fontSize: bodyFontSize,
      color: cardTextColor, margin: 0, valign: "top", fit: "shrink",
    });
  }
  if (hasBullets) {
    const bulletTop = hasBody ? textTop + bodyHeight + 0.1 : textTop;
    addNativeBulletList(
      slide,
      bullets,
      x + padding,
      bulletTop,
      width - padding * 2,
      Math.max(0.5, textBottom - bulletTop),
      briefBulletFontSize(section, currentDensityName === "detailed" ? 10.2 : 11),
      "top",
      cardTextColor,
    );
  }
}

function briefSectionCharacterCount(section) {
  return [...String(section.body ?? "")].length
    + (section.bullets ?? []).reduce((sum, item) => sum + [...String(item)].length, 0);
}

function briefBodyFontSize(section, fallback) {
  const characters = briefSectionCharacterCount(section);
  if (characters < 150) return Math.max(fallback, 13.5);
  if (characters < 230) return Math.max(fallback, 12.2);
  return fallback;
}

function briefBulletFontSize(section, fallback) {
  const characters = briefSectionCharacterCount(section);
  if (characters < 150) return Math.max(fallback, 12.2);
  if (characters < 230) return Math.max(fallback, 11.2);
  return fallback;
}

function renderScorecard(slide, data, index) {
  contentBase(slide, data, index);
  const scorecard = data.scorecard;
  const options = scorecard.options ?? [];
  const criteria = scorecard.criteria ?? [];
  const x = density.outerX;
  const y = density.contentTop;
  const width = W - density.outerX * 2;
  const tableBottom = data.takeaway ? 6.6 : density.contentBottom;
  const height = tableBottom - y;
  const criterionWidth = currentDensityName === "detailed" ? 1.72 : 1.92;
  const optionWidth = (width - criterionWidth) / options.length;
  const headerHeight = currentDensityName === "detailed" ? 0.68 : 0.78;
  const rowHeight = (height - headerHeight) / criteria.length;
  const scorecardTextColor = usesModernRendererContract ? theme.onSurface : theme.text;
  const scorecardMutedTextColor = usesModernRendererContract ? theme.mutedOnSurface : theme.muted;
  const scorecardBorderColor = usesModernRendererContract
    ? mixHex(theme.surface, theme.onSurface, 0.22)
    : "D9E0E8";
  const tableRows = [
    [
      {
        text: "評価軸",
        options: {
          fontFace: theme.headingFont,
          bold: true,
          color: theme.onPrimary,
          fill: { color: theme.primary },
          align: "center",
        },
      },
      ...options.map((option) => ({
        text: option.subtitle
          ? [
              { text: option.title, options: { fontFace: theme.headingFont, bold: true, breakLine: true, color: theme.onPrimary } },
              { text: option.subtitle, options: { fontSize: 9, color: theme.onPrimary } },
            ]
          : option.title,
        options: {
          fontFace: theme.headingFont,
          bold: true,
          color: theme.onPrimary,
          fill: { color: theme.primary },
          align: "center",
        },
      })),
    ],
    ...criteria.map((row, rowIndex) => [
      {
        text: row.criterion,
        options: {
          fontFace: theme.headingFont,
          bold: true,
          color: scorecardTextColor,
          fill: {
            color: usesModernRendererContract
              ? rowIndex % 2 === 0
                ? mixHex(theme.surface, theme.secondary, 0.1)
                : mixHex(theme.surface, theme.onSurface, 0.035)
              : rowIndex % 2 === 0 ? "EEF2F7" : "F7F9FC",
          },
          align: "left",
        },
      },
      ...row.cells.map((cell, cellIndex) => {
        const tone = toneColor(cell.tone, cellIndex);
        const cellFill = usesModernRendererContract ? mixHex(theme.surface, tone, 0.1) : tone;
        const ratingColor = usesModernRendererContract
          ? contrastRatio(tone, cellFill) >= 4.5 ? tone : theme.onSurface
          : tone;
        return {
          text: cell.detail
            ? [
                { text: `● ${cell.rating}`, options: { bold: true, breakLine: true, color: ratingColor } },
                { text: cell.detail, options: { fontSize: currentDensityName === "detailed" ? 9 : 9.5, color: scorecardMutedTextColor } },
              ]
            : `● ${cell.rating}`,
          options: {
            bold: !cell.detail,
            color: ratingColor,
            fill: usesModernRendererContract
              ? { color: cellFill }
              : { color: cellFill, transparency: 92 },
            align: "left",
          },
        };
      }),
    ]),
  ];

  slide.addTable(tableRows, {
    x, y, w: width, h: height,
    colW: [criterionWidth, ...options.map(() => optionWidth)],
    rowH: [headerHeight, ...criteria.map(() => rowHeight)],
    fontFace: theme.bodyFont,
    fontSize: currentDensityName === "detailed" ? 9.5 : 10.5,
    color: scorecardTextColor,
    margin: currentDensityName === "detailed" ? 0.08 : 0.11,
    // PptxGenJS serializes table-level "mid" literally as anchor="mid",
    // which is not a valid DrawingML value. "middle" is normalized to "ctr".
    valign: "middle",
    border: {
      type: "solid",
      color: scorecardBorderColor,
      width: 0.55,
    },
    autoFit: false,
    autoPage: false,
  });

  if (data.takeaway) {
    slide.addText(data.takeaway, {
      x,
      y: 6.72,
      w: width,
      h: 0.28,
      fontFace: theme.bodyFont,
      fontSize: 10,
      bold: true,
      color: theme.secondaryTextOnBackground,
      margin: 0,
      align: "center",
      fit: "shrink",
    });
  }
}

function renderDataTable(slide, data, index) {
  contentBase(slide, data, index);
  const table = data.dataTable;
  const columns = table.columns ?? [];
  const rows = table.rows ?? [];
  const x = density.outerX;
  const y = density.contentTop;
  const width = W - density.outerX * 2;
  const tableBottom = data.takeaway ? 6.58 : density.contentBottom;
  const height = tableBottom - y;
  const headerHeight = currentDensityName === "detailed" ? 0.52 : 0.62;
  const rowHeight = (height - headerHeight) / rows.length;
  const totalWeight = columns.reduce((sum, column) => sum + Number(column.widthWeight ?? 1), 0);
  const columnWidths = columns.map((column) =>
    width * Number(column.widthWeight ?? 1) / totalWeight);
  const bodyFontSize = currentDensityName === "detailed"
    ? 9.2
    : rows.length >= 7
      ? 9.4
      : 10.5;
  const tableRows = [
    columns.map((column) => ({
      text: column.header,
      options: {
        fontFace: theme.headingFont,
        bold: true,
        color: theme.onPrimary,
        fill: { color: theme.primary },
        align: normalizeTableAlignment(column.align),
      },
    })),
    ...rows.map((row, rowIndex) =>
      row.cells.map((cell, cellIndex) => {
        const tone = toneColor(cell.tone, cellIndex);
        const emphasized = cell.emphasize === true;
        const rowHeader = table.firstColumnIsHeader !== false && cellIndex === 0;
        const baseFill = rowIndex % 2 === 0
          ? mixHex(theme.surface, theme.onSurface, 0.035)
          : theme.surface;
        const cellFill = emphasized
          ? mixHex(theme.surface, tone, 0.1)
          : rowHeader ? mixHex(theme.surface, theme.secondary, 0.1) : baseFill;
        return {
          text: cell.text,
          options: {
            bold: emphasized || rowHeader,
            color: emphasized && contrastRatio(tone, cellFill) >= 4.5
              ? tone
              : theme.onSurface,
            fill: { color: cellFill },
            align: normalizeTableAlignment(columns[cellIndex].align),
          },
        };
      })),
  ];

  slide.addTable(tableRows, {
    x, y, w: width, h: height,
    colW: columnWidths,
    rowH: [headerHeight, ...rows.map(() => rowHeight)],
    fontFace: theme.bodyFont,
    fontSize: bodyFontSize,
    color: theme.onSurface,
    margin: currentDensityName === "detailed" ? 0.07 : 0.1,
    valign: "middle",
    border: {
      type: "solid",
      color: styleProfile.borderColor,
      width: Math.max(0.45, density.cardBorderWidth * styleProfile.borderWidthScale),
    },
    autoFit: false,
    autoPage: false,
  });

  if (data.takeaway) {
    slide.addText(data.takeaway, {
      x,
      y: 6.72,
      w: width,
      h: 0.28,
      fontFace: theme.bodyFont,
      fontSize: 10,
      bold: true,
      color: theme.secondaryTextOnBackground,
      margin: 0,
      align: "center",
      fit: "shrink",
    });
  }
}

function normalizeTableAlignment(value) {
  const alignment = String(value ?? "left").toLowerCase();
  return ["left", "center", "right"].includes(alignment) ? alignment : "left";
}

function renderMusicScore(slide, data, index) {
  contentBase(slide, data, index);
  const musicScore = data.musicScore;
  const measures = musicScore.measures ?? [];
  const splitIndex = measures.length <= 4 ? measures.length : Math.ceil(measures.length / 2);
  const systems = measures.length <= 4
    ? [measures]
    : [measures.slice(0, splitIndex), measures.slice(splitIndex)];
  const cardY = systems.length === 1 ? 2.08 : 1.76;
  const cardHeight = systems.length === 1 ? 4.45 : 5.02;
  card(slide, 0.58, cardY, W - 1.16, cardHeight, "FFFFFF");

  const metadata = [
    `Key ${musicScore.keySignature}`,
    musicScore.timeSignature,
    musicScore.tempoBpm ? `Tempo ${musicScore.tempoBpm} BPM` : null,
    String(musicScore.tuning).toLowerCase() === "low-g" ? "Low-G" : "High-G",
  ].filter(Boolean).join("  ·  ");
  slide.addText(metadata, {
    x: 0.92,
    y: cardY + 0.12,
    w: W - 1.84,
    h: 0.25,
    fontFace: theme.bodyFont,
    fontSize: 9.5,
    bold: true,
    color: cardMutedTextColor,
    margin: 0,
    align: "right",
    fit: "shrink",
  });

  const systemYPositions = systems.length === 1
    ? [cardY + 0.62]
    : [cardY + 0.5, cardY + 2.58];
  let measureOffset = 0;
  systems.forEach((systemMeasures, systemIndex) => {
    renderMusicSystem(
      slide,
      musicScore,
      systemMeasures,
      measureOffset,
      systemYPositions[systemIndex],
    );
    measureOffset += systemMeasures.length;
  });

  const legendY = systems.length === 1 ? cardY + 3.18 : cardY + 4.47;
  if (musicScore.colorFingerings && musicScore.showTablature) {
    renderFingerLegend(slide, legendY);
  }
  if (musicScore.caption) {
    slide.addText(musicScore.caption, {
      x: musicScore.colorFingerings && musicScore.showTablature ? 5.5 : 0.92,
      y: legendY - 0.02,
      w: musicScore.colorFingerings && musicScore.showTablature ? 6.85 : W - 1.84,
      h: 0.32,
      fontFace: theme.bodyFont,
      fontSize: 9.5,
      color: cardMutedTextColor,
      margin: 0,
      align: "right",
      fit: "shrink",
    });
  }
}

function renderMusicSystem(slide, musicScore, measures, measureOffset, systemY) {
  const showStaff = musicScore.showStandardNotation;
  const showTab = musicScore.showTablature;
  const staffSpacing = 0.125;
  const tabSpacing = 0.145;
  const staffTop = systemY + (showStaff ? 0.18 : 0);
  const staffBottom = staffTop + staffSpacing * 4;
  const tabTop = showTab
    ? systemY + (showStaff ? 1.08 : 0.3)
    : null;
  const tabBottom = showTab ? tabTop + tabSpacing * 3 : null;
  const prefixX = 0.82;
  const notationX = showStaff ? 1.84 : 1.52;
  const notationRight = 12.5;
  const notationWidth = notationRight - notationX;

  if (showStaff) {
    for (let lineIndex = 0; lineIndex < 5; lineIndex++) {
      const lineY = staffTop + lineIndex * staffSpacing;
      slide.addShape(pptx.ShapeType.line, {
        x: notationX,
        y: lineY,
        w: notationWidth,
        h: 0,
        line: { color: cardTextColor, transparency: 12, width: 0.75 },
      });
    }
    renderTrebleClef(slide, prefixX, staffTop, staffSpacing);
    renderTimeSignature(slide, musicScore.timeSignature, 1.55, staffTop, staffSpacing);
  }

  if (showTab) {
    for (let stringIndex = 0; stringIndex < 4; stringIndex++) {
      const lineY = tabTop + stringIndex * tabSpacing;
      slide.addShape(pptx.ShapeType.line, {
        x: notationX,
        y: lineY,
        w: notationWidth,
        h: 0,
        line: { color: cardMutedTextColor, transparency: 18, width: 0.75 },
      });
    }
    slide.addText("TAB", {
      x: prefixX - 0.04,
      y: tabTop + 0.06,
      w: 0.72,
      h: 0.26,
      fontFace: theme.bodyFont,
      fontSize: 10,
      bold: true,
      color: readableToneOnSurface(theme.secondary),
      margin: 0,
      align: "center",
      valign: "mid",
      fit: "shrink",
    });
  }

  const measureWidth = notationWidth / measures.length;
  measures.forEach((measure, measureIndex) => {
    const measureX = notationX + measureIndex * measureWidth;
    const measureRight = measureX + measureWidth;
    const displayedNumber = measure.number ?? measureOffset + measureIndex + 1;
    const labelY = showStaff ? staffTop - 0.32 : tabTop - 0.3;
    slide.addText(String(displayedNumber), {
      x: measureX + 0.06,
      y: labelY,
      w: 0.34,
      h: 0.2,
      fontFace: theme.bodyFont,
      fontSize: 7.5,
      bold: true,
      color: cardMutedTextColor,
      margin: 0,
      fit: "shrink",
    });

    if (measureIndex > 0) {
      renderMusicBarLine(slide, measureX, staffTop, staffBottom, tabTop, tabBottom, showStaff, showTab, 0.85);
    }

    const events = measure.events ?? [];
    const eventPositions = events.map((_, eventIndex) =>
      measureX + measureWidth * (eventIndex + 0.5) / events.length);
    events.forEach((musicEvent, eventIndex) => {
      const eventX = eventPositions[eventIndex];
      if (showStaff) {
        renderStandardMusicEvent(
          slide,
          musicEvent,
          eventX,
          eventPositions[eventIndex + 1],
          staffTop,
          staffBottom,
          staffSpacing,
        );
      }
      if (showTab) {
        renderTabMusicEvent(slide, musicScore, musicEvent, eventX, tabTop, tabSpacing);
      }
    });

    if (measureIndex === measures.length - 1) {
      renderMusicBarLine(slide, measureRight, staffTop, staffBottom, tabTop, tabBottom, showStaff, showTab, 1.35);
      slide.addShape(pptx.ShapeType.line, {
        x: measureRight - 0.055,
        y: showStaff ? staffTop : tabTop,
        w: 0,
        h: showStaff && showTab ? tabBottom - staffTop : showStaff ? staffBottom - staffTop : tabBottom - tabTop,
        line: { color: cardTextColor, width: 0.55, transparency: 20 },
      });
    }
  });
}

function renderMusicBarLine(slide, x, staffTop, staffBottom, tabTop, tabBottom, showStaff, showTab, width) {
  if (showStaff) {
    slide.addShape(pptx.ShapeType.line, {
      x, y: staffTop, w: 0, h: staffBottom - staffTop,
      line: { color: cardTextColor, width },
    });
  }
  if (showTab) {
    slide.addShape(pptx.ShapeType.line, {
      x, y: tabTop, w: 0, h: tabBottom - tabTop,
      line: { color: cardTextColor, width },
    });
  }
}

function addMusicLine(slide, x, y, width, height, line) {
  slide.addShape(pptx.ShapeType.line, {
    x: x + Math.min(0, width),
    y: y + Math.min(0, height),
    w: Math.abs(width),
    h: Math.abs(height),
    flipH: width < 0,
    flipV: height < 0,
    line,
  });
}

function renderTrebleClef(slide, x, staffTop, spacing) {
  addBravuraGlyph(slide, pptx, bravuraFont, MusicGlyph.gClef, {
    originX: x + 0.1,
    originY: staffTop + spacing * 3,
    scale: musicGlyphScale(spacing),
    color: readableToneOnSurface(theme.secondary),
  });
}

function renderTimeSignature(slide, timeSignature, centerX, staffTop, spacing) {
  const [numerator, denominator] = String(timeSignature).split("/");
  [numerator, denominator].forEach((value, rowIndex) => {
    const digits = [...value];
    const digitSpacing = 0.17;
    digits.forEach((digit, digitIndex) => {
      addBravuraGlyph(slide, pptx, bravuraFont, timeSignatureGlyph(digit), {
        centerX: centerX + (digitIndex - (digits.length - 1) / 2) * digitSpacing,
        centerY: staffTop + spacing * (rowIndex === 0 ? 1 : 3),
        scale: musicGlyphScale(spacing),
        color: cardTextColor,
      });
    });
  });
}

function renderStandardMusicEvent(slide, musicEvent, eventX, nextEventX, staffTop, staffBottom, staffSpacing) {
  const duration = String(musicEvent.duration).toLowerCase();
  if (musicEvent.annotation) {
    slide.addText(musicEvent.annotation, {
      x: eventX - 0.38,
      y: staffTop - 0.27,
      w: 0.76,
      h: 0.16,
      fontFace: theme.bodyFont,
      fontSize: 6.8,
      bold: true,
      color: readableToneOnSurface(theme.secondary),
      margin: 0,
      align: "center",
      fit: "shrink",
    });
  }

  if (musicEvent.rest) {
    renderMusicRest(slide, duration, eventX, staffTop, staffSpacing, musicEvent.dotted);
    return;
  }

  const notePositions = (musicEvent.notes ?? []).map((note) => ({
    note,
    step: musicPitchStep(note.pitch),
  })).map((item) => ({
    ...item,
    y: staffBottom - item.step * staffSpacing / 2,
  }));
  notePositions.forEach(({ note, step, y }) => {
    renderLedgerLines(slide, eventX, step, staffBottom, staffSpacing);
    addBravuraGlyph(slide, pptx, bravuraFont, musicNoteheadGlyph(duration), {
      centerX: eventX,
      centerY: y,
      scale: musicGlyphScale(staffSpacing),
      color: cardTextColor,
    });
    const accidentalGlyph = musicAccidentalGlyph(note.pitch);
    if (accidentalGlyph) {
      addBravuraGlyph(slide, pptx, bravuraFont, accidentalGlyph, {
        centerX: eventX - 0.14,
        centerY: y,
        scale: musicGlyphScale(staffSpacing),
        color: cardTextColor,
      });
    }
  });

  if (duration !== "whole") {
    renderMusicStemAndFlags(slide, duration, eventX, notePositions, staffSpacing);
  }
  if (musicEvent.dotted && notePositions.length > 0) {
    const dotY = notePositions.reduce((sum, item) => sum + item.y, 0) / notePositions.length;
    addBravuraGlyph(slide, pptx, bravuraFont, MusicGlyph.augmentationDot, {
      centerX: eventX + 0.12,
      centerY: dotY,
      scale: musicGlyphScale(staffSpacing),
      color: cardTextColor,
    });
  }
  if (musicEvent.tieToNext && nextEventX) {
    const lowestY = Math.max(...notePositions.map((item) => item.y));
    slide.addShape(pptx.ShapeType.arc, {
      x: eventX + 0.04,
      y: lowestY + 0.03,
      w: Math.max(0.15, nextEventX - eventX - 0.08),
      h: 0.18,
      rotate: 180,
      fill: { color: cardTextColor, transparency: 100 },
      line: { color: cardTextColor, width: 0.7 },
    });
  }
}

function renderMusicStemAndFlags(slide, duration, eventX, notePositions, staffSpacing) {
  const averageStep = notePositions.reduce((sum, item) => sum + item.step, 0) / notePositions.length;
  const stemUp = averageStep < 4;
  const minimumY = Math.min(...notePositions.map((item) => item.y));
  const maximumY = Math.max(...notePositions.map((item) => item.y));
  const stemX = eventX + (stemUp ? 0.066 : -0.066);
  const stemTop = stemUp ? minimumY - 0.42 : minimumY;
  const stemBottom = stemUp ? maximumY : maximumY + 0.42;
  slide.addShape(pptx.ShapeType.line, {
    x: stemX, y: stemTop, w: 0, h: stemBottom - stemTop,
    line: { color: cardTextColor, width: 0.9 },
  });

  const flagGlyph = musicFlagGlyph(duration, stemUp);
  if (flagGlyph) {
    addBravuraGlyph(slide, pptx, bravuraFont, flagGlyph, {
      originX: stemX,
      originY: stemUp ? stemTop : stemBottom,
      scale: musicGlyphScale(staffSpacing),
      color: cardTextColor,
    });
  }
}

function renderLedgerLines(slide, eventX, step, staffBottom, staffSpacing) {
  if (step < 0) {
    for (let ledgerStep = -2; ledgerStep >= step; ledgerStep -= 2) {
      const lineY = staffBottom - ledgerStep * staffSpacing / 2;
      slide.addShape(pptx.ShapeType.line, {
        x: eventX - 0.13, y: lineY, w: 0.26, h: 0,
        line: { color: cardTextColor, width: 0.7 },
      });
    }
  } else if (step > 8) {
    for (let ledgerStep = 10; ledgerStep <= step; ledgerStep += 2) {
      const lineY = staffBottom - ledgerStep * staffSpacing / 2;
      slide.addShape(pptx.ShapeType.line, {
        x: eventX - 0.13, y: lineY, w: 0.26, h: 0,
        line: { color: cardTextColor, width: 0.7 },
      });
    }
  }
}

function renderMusicRest(slide, duration, eventX, staffTop, staffSpacing, dotted) {
  const middleY = staffTop + staffSpacing * 2;
  const restGlyph = musicRestGlyph(duration);
  if (duration === "whole" || duration === "half") {
    addBravuraGlyph(slide, pptx, bravuraFont, restGlyph, {
      originX: eventX - 0.071,
      originY: middleY,
      scale: musicGlyphScale(staffSpacing),
      color: cardTextColor,
    });
  } else if (duration === "quarter") {
    addBravuraGlyph(slide, pptx, bravuraFont, restGlyph, {
      centerX: eventX,
      centerY: middleY,
      scale: musicGlyphScale(staffSpacing),
      color: cardTextColor,
    });
  } else {
    addBravuraGlyph(slide, pptx, bravuraFont, restGlyph, {
      centerX: eventX,
      centerY: middleY,
      scale: musicGlyphScale(staffSpacing),
      color: cardTextColor,
    });
  }
  if (dotted) {
    addBravuraGlyph(slide, pptx, bravuraFont, MusicGlyph.augmentationDot, {
      centerX: eventX + 0.12,
      centerY: middleY,
      scale: musicGlyphScale(staffSpacing),
      color: cardTextColor,
    });
  }
}

function renderTabMusicEvent(slide, musicScore, musicEvent, eventX, tabTop, tabSpacing) {
  if (musicEvent.rest) {
    slide.addText("rest", {
      x: eventX - 0.18, y: tabTop + tabSpacing - 0.02, w: 0.36, h: 0.18,
      fontFace: theme.bodyFont, fontSize: 6.5, bold: true, color: cardMutedTextColor,
      margin: 0, align: "center", fit: "shrink",
    });
    return;
  }

  (musicEvent.notes ?? []).forEach((note) => {
    const stringY = tabTop + (note.string - 1) * tabSpacing;
    const markerColor = musicScore.colorFingerings
      ? musicFingerColor(note)
      : "FFFFFF";
    const isOpenOrNeutral = !musicScore.colorFingerings || note.finger == null || note.finger === 0;
    const outline = isOpenOrNeutral ? cardMutedTextColor : markerColor;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: eventX - 0.095,
      y: stringY - 0.095,
      w: 0.19,
      h: 0.19,
      fill: { color: markerColor },
      line: { color: outline, width: isOpenOrNeutral ? 0.75 : 0.4 },
    });
    slide.addText(String(note.fret), {
      x: eventX - 0.095,
      y: stringY - 0.09,
      w: 0.19,
      h: 0.18,
      fontFace: theme.bodyFont,
      fontSize: note.fret >= 10 ? 6.8 : 7.8,
      bold: true,
      color: usesModernRendererContract
        ? readableForeground(markerColor)
        : isOpenOrNeutral || luminance(markerColor) > 0.56 ? theme.text : "FFFFFF",
      margin: 0,
      align: "center",
      valign: "mid",
      fit: "shrink",
    });
  });
}

function renderFingerLegend(slide, y) {
  const legendItems = [
    [0, "開放"],
    [1, "人差し指"],
    [2, "中指"],
    [3, "薬指"],
    [4, "小指"],
  ];
  slide.addText("左手", {
    x: 0.92, y, w: 0.45, h: 0.22,
    fontFace: theme.bodyFont, fontSize: 8, bold: true, color: cardMutedTextColor,
    margin: 0, valign: "mid",
  });
  legendItems.forEach(([finger, label], itemIndex) => {
    const x = 1.42 + itemIndex * 0.78;
    const color = musicFingerColor({ finger });
    slide.addShape(pptx.ShapeType.ellipse, {
      x, y: y + 0.02, w: 0.16, h: 0.16,
      fill: { color },
      line: { color: finger === 0 ? cardMutedTextColor : color, width: 0.6 },
    });
    slide.addText(label, {
      x: x + 0.21, y, w: 0.54, h: 0.2,
      fontFace: theme.bodyFont, fontSize: 7.2, color: cardMutedTextColor,
      margin: 0, valign: "mid", fit: "shrink",
    });
  });
}

function musicFingerColor(note) {
  if (note.color && /^#?[0-9a-f]{6}$/i.test(note.color)) {
    return String(note.color).replace(/^#/, "").toUpperCase();
  }
  const colors = {
    0: "FFFFFF",
    1: "E5484D",
    2: "2F6FED",
    3: "22A06B",
    4: "8E5AD7",
  };
  return colors[note.finger] ?? "FFFFFF";
}

function musicPitchStep(pitch) {
  const match = /^([A-Ga-g])(?:#|b)?([0-8])$/.exec(String(pitch));
  if (!match) return 0;
  const letterSteps = { C: 0, D: 1, E: 2, F: 3, G: 4, A: 5, B: 6 };
  const absoluteStep = Number(match[2]) * 7 + letterSteps[match[1].toUpperCase()];
  const e4Step = 4 * 7 + letterSteps.E;
  return absoluteStep - e4Step;
}

function musicGlyphScale(staffSpacing) {
  // SMuFL defines one staff space as 250 font units for a 1000 UPM font.
  return staffSpacing / 250;
}

function musicNoteheadGlyph(duration) {
  if (duration === "whole") return MusicGlyph.noteheadWhole;
  if (duration === "half") return MusicGlyph.noteheadHalf;
  return MusicGlyph.noteheadBlack;
}

function musicAccidentalGlyph(pitch) {
  if (String(pitch).includes("#")) return MusicGlyph.accidentalSharp;
  if (String(pitch).includes("b")) return MusicGlyph.accidentalFlat;
  return null;
}

function musicFlagGlyph(duration, stemUp) {
  if (duration === "eighth") {
    return stemUp ? MusicGlyph.flag8thUp : MusicGlyph.flag8thDown;
  }
  if (duration === "sixteenth") {
    return stemUp ? MusicGlyph.flag16thUp : MusicGlyph.flag16thDown;
  }
  return null;
}

function musicRestGlyph(duration) {
  const glyphs = {
    whole: MusicGlyph.restWhole,
    half: MusicGlyph.restHalf,
    quarter: MusicGlyph.restQuarter,
    eighth: MusicGlyph.rest8th,
    sixteenth: MusicGlyph.rest16th,
  };
  const glyph = glyphs[duration];
  if (!glyph) throw new Error(`Unsupported rest duration: ${duration}`);
  return glyph;
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
      fontFace: theme.bodyFont, fontSize: 12, bold: true,
      color: theme.secondaryTextOnBackground, margin: 0, align: "center", fit: "shrink",
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
      fontFace: theme.bodyFont, fontSize: 10.5, bold: true,
      color: theme.secondaryTextOnBackground, margin: 0, align: "center", fit: "shrink",
    });
    slide.addText(step.title, {
      x: textX,
      y: above ? 2.4 : 4.75,
      w: textWidth, h: 0.48,
      fontFace: theme.bodyFont, fontSize: 14, bold: true,
      color: theme.text, margin: 0, align: "center", valign: "mid", fit: "shrink",
    });
    if (step.description) {
      slide.addText(step.description, {
        x: textX,
        y: above ? 2.98 : 5.29,
        w: textWidth, h: 0.42,
        fontFace: theme.bodyFont, fontSize: 10.5, color: theme.muted,
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
  const chartRuleColor = usesModernRendererContract ? styleProfile.borderColor : "D9DFE8";
  const chartOptions = {
    x: 1.0, y: 2.23, w: chartWidth - 0.58, h: 4.18,
    showTitle: false,
    showLegend: chart.showLegend,
    legendPos: "b",
    legendFontFace: theme.bodyFont,
    legendFontSize: 10,
    ...(usesModernRendererContract ? { legendColor: cardMutedTextColor } : {}),
    chartColors: theme.dataSeries,
    showValue: chartType === "pie" || chartType === "doughnut",
    showPercent: chartType === "pie" || chartType === "doughnut",
    dataLabelPosition: "bestFit",
    catAxisLabelFontFace: theme.bodyFont,
    catAxisLabelFontSize: 10,
    catAxisLabelColor: cardMutedTextColor,
    valAxisLabelFontFace: theme.bodyFont,
    valAxisLabelFontSize: 10,
    valAxisLabelColor: cardMutedTextColor,
    valAxisLineColor: chartRuleColor,
    catAxisLineColor: chartRuleColor,
    valGridLine: { color: chartRuleColor, size: 1 },
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
  addDarkSlideStyle(slide, index);
  if (data.eyebrow) {
    slide.addText(data.eyebrow.toUpperCase(), {
      x: 0.78, y: 0.65, w: 8.4, h: 0.32,
      fontFace: theme.bodyFont, fontSize: scaled(10.5), bold: true,
      charSpacing: 1.6, color: theme.accentTextOnPrimary, margin: 0,
    });
  }
  slide.addText(data.title, {
    x: 0.78, y: 1.12, w: 6.8, h: 0.68,
    fontFace: theme.headingFont, fontSize: scaled(21), bold: true,
    color: theme.onPrimary, transparency: 18, margin: 0, fit: "shrink",
  });
  slide.addText(data.body, {
    x: 0.78, y: 2.0, w: 8.3, h: 3.3,
    fontFace: theme.headingFont,
    fontSize: scaled(titleSize(data.body, 34, 23) * styleProfile.titleScale),
    bold: true,
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
    fontFace: theme.bodyFont, fontSize: scaled(16), bold: true,
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
      fontFace: theme.bodyFont, fontSize: scaled(15), bold: true,
      color: theme.text, margin: 0, fit: "shrink",
    });
    if (quadrant.highlight) {
      slide.addText(quadrant.highlight, {
        x: cellX + cellWidth - 1.25, y: cellY + 0.22, w: 0.98, h: 0.28,
        fontFace: theme.bodyFont, fontSize: scaled(9), bold: true,
        color: readableToneOnBackground(colors[quadrantIndex], theme.background, theme.text),
        margin: 0, align: "right", fit: "shrink",
      });
    }
    addNativeBulletList(slide, quadrant.bullets, cellX + 0.25, cellY + 0.72, cellWidth - 0.5, cellHeight - 0.9, 11.5);
  });
  slide.addText(matrix.horizontalAxis, {
    x: 4.1, y: 6.48, w: 5.25, h: 0.26,
    fontFace: theme.bodyFont, fontSize: scaled(9.5), bold: true,
    color: theme.muted, margin: 0, align: "center", fit: "shrink",
  });
  slide.addText(matrix.verticalAxis, {
    x: -0.42, y: 4.0, w: 2.4, h: 0.32,
    fontFace: theme.bodyFont, fontSize: scaled(9.5), bold: true,
    color: theme.muted, margin: 0, align: "center", valign: "mid", rotate: 270, fit: "shrink",
  });
  if (data.takeaway) {
    slide.addText(data.takeaway, {
      x: 1.7, y: 6.78, w: 9.95, h: 0.2,
      fontFace: theme.bodyFont, fontSize: scaled(9.5), bold: true,
      color: theme.secondaryTextOnBackground, margin: 0, align: "center", fit: "shrink",
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
      fontFace: theme.bodyFont, fontSize: scaled(13.5), bold: true,
      color: usesModernRendererContract
        ? readableForeground(color)
        : luminance(color) > 0.56 ? theme.text : "FFFFFF",
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
      fontFace: theme.bodyFont, fontSize: scaled(12), bold: true,
      color: cardTextColor, margin: 0, fit: "shrink",
    });
    if (step.description) {
      slide.addText(step.description, {
        x: x + 0.2, y: y + 0.74, w: width - 0.4, h: 0.44,
        fontFace: theme.bodyFont, fontSize: scaled(9.5),
        color: cardMutedTextColor, margin: 0, align: "center", fit: "shrink",
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
      fontFace: theme.bodyFont, fontSize: scaled(11.5), bold: true,
      color: theme.secondaryTextOnBackground, margin: 0, align: "center", fit: "shrink",
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
  const chartRuleColor = usesModernRendererContract ? styleProfile.borderColor : "D9DFE8";
  slide.addChart(typeMap[chartType], series, {
    x: 0.98, y: 3.82, w: chartWidth - 0.56, h: 2.58,
    showLegend: chart.showLegend,
    legendPos: "b",
    legendFontFace: theme.bodyFont,
    legendFontSize: 9,
    ...(usesModernRendererContract ? { legendColor: cardMutedTextColor } : {}),
    chartColors: theme.dataSeries,
    showValue: chartType === "pie" || chartType === "doughnut",
    showPercent: chartType === "pie" || chartType === "doughnut",
    dataLabelPosition: "bestFit",
    catAxisLabelFontFace: theme.bodyFont,
    catAxisLabelFontSize: 9,
    catAxisLabelColor: cardMutedTextColor,
    valAxisLabelFontFace: theme.bodyFont,
    valAxisLabelFontSize: 9,
    valAxisLabelColor: cardMutedTextColor,
    valAxisLineColor: chartRuleColor,
    catAxisLineColor: chartRuleColor,
    valGridLine: { color: chartRuleColor, size: 1 },
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
    color: theme.secondaryTextOnBackground, transparency: 20, margin: 0,
  });
  if (data.eyebrow) {
    slide.addText(data.eyebrow.toUpperCase(), {
      x: 2.0, y: 0.88, w: 8.7, h: 0.32,
      fontFace: theme.bodyFont, fontSize: 11, bold: true,
      charSpacing: 1.6, color: theme.secondaryTextOnBackground, margin: 0,
    });
  }
  slide.addText(data.body, {
    x: 1.72, y: 1.63, w: 9.92, h: 3.52,
    fontFace: theme.headingFont, fontSize: titleSize(data.body, 28, 20), bold: true,
    color: theme.text, margin: 0, valign: "mid", align: "center", fit: "shrink",
  });
  if (data.attribution) {
    slide.addText(`— ${data.attribution}`, {
      x: 3.1, y: 5.53, w: 7.12, h: 0.5,
      fontFace: theme.bodyFont, fontSize: 14, color: theme.muted,
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
  addDarkSlideStyle(slide, index);
  slide.addText(data.title, {
    x: 1.25, y: hasBullets ? 0.88 : 1.73, w: 10.83, h: hasBullets ? 1.08 : 1.45,
    fontFace: theme.headingFont,
    fontSize: titleSize(data.title, hasBullets ? 31 : 35, 25) * styleProfile.titleScale,
    bold: true,
    color: theme.onPrimary, margin: 0, align: "center", valign: "mid", fit: "shrink",
  });
  if (data.subtitle || data.body) {
    slide.addText(data.subtitle ?? data.body, {
      x: 2.0, y: hasBullets ? 2.05 : 3.48, w: 9.33, h: hasBullets ? 0.72 : 1.2,
      fontFace: theme.bodyFont, fontSize: hasBullets ? 15 : 18, color: theme.onPrimary,
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
      fontFace: theme.bodyFont, color: theme.onPrimary,
      margin: 0.04, valign: "mid", fit: "shrink",
    });
  }
  if (data.takeaway) {
    pill(
      slide,
      data.takeaway,
      3.55,
      hasBullets ? 5.62 : 5.34,
      6.23,
      theme.accent,
      usesModernRendererContract ? theme.onAccent : theme.primary,
    );
  }
  footer(slide, index, true);
}

function contentBase(slide, data, index) {
  background(slide, theme.background);
  addMotif(slide, index, false, true);
  addContentStyle(slide, index);
  if (data.eyebrow) {
    pill(
      slide,
      data.eyebrow.toUpperCase(),
      density.outerX,
      currentDensityName === "detailed" ? 0.16 : 0.28,
      Math.min(3.4, 0.9 + data.eyebrow.length * 0.115),
      theme.secondary,
      usesModernRendererContract ? theme.onSecondary : "FFFFFF",
    );
  }
  slide.addText(data.title, {
    x: density.outerX,
    y: data.eyebrow ? (currentDensityName === "detailed" ? 0.68 : 0.84) : density.titleY,
    w: W - density.outerX * 2 - 0.95,
    h: density.titleHeight,
    fontFace: theme.headingFont,
    fontSize: scaled(titleSize(data.title, 25, 20) * styleProfile.titleScale),
    bold: true,
    color: theme.text, margin: 0, valign: "mid", fit: "shrink",
  });
  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: density.outerX, y: density.subtitleY, w: W - density.outerX * 2, h: 0.34,
      fontFace: theme.bodyFont, fontSize: scaled(12), color: theme.muted,
      margin: 0, fit: "shrink",
    });
  }
  if (!templateChrome) {
    slide.addText(String(index + 1).padStart(2, "0"), {
      x: 11.72, y: 0.42, w: 0.8, h: 0.42,
      fontFace: theme.bodyFont, fontSize: scaled(13), bold: true,
      color: theme.secondaryTextOnBackground, transparency: 20, margin: 0, align: "right",
    });
  }
  footer(slide, index, false);
}

function addDarkSlideStyle(slide, index) {
  if (!usesModernRendererContract) return;

  if (design.style === "editorial") {
    slide.addShape(pptx.ShapeType.line, {
      x: 0.76,
      y: 6.68,
      w: 5.2,
      h: 0,
      line: { color: theme.accent, width: 1.2, transparency: 18 },
    });
    return;
  }

  if (design.style === "bold") {
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 0.2,
      h: H,
      fill: { color: theme.accent },
      line: { color: theme.accent, transparency: 100 },
    });
    return;
  }

  if (design.style === "technical") {
    for (let lineIndex = 0; lineIndex < 5; lineIndex += 1) {
      slide.addShape(pptx.ShapeType.line, {
        x: 0.7,
        y: 1.2 + lineIndex * 1.16,
        w: W - 1.4,
        h: 0,
        line: { color: theme.onPrimary, width: 0.35, transparency: 91 },
      });
    }
    return;
  }

  if (design.style === "playful") {
    const colors = [theme.accent, theme.secondary, theme.onPrimary];
    colors.forEach((color, colorIndex) => {
      slide.addShape(pptx.ShapeType.ellipse, {
        x: 0.78 + colorIndex * 0.23,
        y: 6.62 + (index % 2) * 0.05,
        w: 0.13,
        h: 0.13,
        fill: { color, transparency: colorIndex === 2 ? 55 : 0 },
        line: { color, transparency: 100 },
      });
    });
  }
}

function addSectionStyle(slide, index) {
  if (!usesModernRendererContract) return;

  if (design.style === "executive") {
    slide.addShape(pptx.ShapeType.rect, {
      x: 4.95,
      y: 1.66,
      w: 1.24,
      h: 0.06,
      fill: { color: theme.secondary },
      line: { color: theme.secondary, transparency: 100 },
    });
    return;
  }

  if (design.style === "editorial") {
    slide.addShape(pptx.ShapeType.line, {
      x: 12.28,
      y: 1.32,
      w: 0,
      h: 4.52,
      line: { color: theme.accent, width: 1.15, transparency: 18 },
    });
    return;
  }

  if (design.style === "bold") {
    slide.addShape(pptx.ShapeType.rect, {
      x: 4.51,
      y: 0,
      w: W - 4.51,
      h: 0.16,
      fill: { color: index % 2 === 0 ? theme.secondary : theme.accent },
      line: { color: theme.secondary, transparency: 100 },
    });
    return;
  }

  if (design.style === "technical") {
    [0.78, 6.66].forEach((y) => {
      slide.addShape(pptx.ShapeType.line, {
        x: 4.78,
        y,
        w: 7.62,
        h: 0,
        line: { color: theme.primary, width: 0.4, transparency: 86 },
      });
    });
    return;
  }

  [theme.accent, theme.secondary].forEach((color, colorIndex) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 11.72 + colorIndex * 0.23,
      y: 6.52 + (index % 2) * 0.04,
      w: 0.13,
      h: 0.13,
      fill: { color },
      line: { color, transparency: 100 },
    });
  });
}

function addContentStyle(slide, index) {
  if (!usesModernRendererContract) return;

  if (design.style === "executive") {
    slide.addShape(pptx.ShapeType.rect, {
      x: density.outerX,
      y: density.contentTop - 0.08,
      w: 1.32,
      h: 0.055,
      fill: { color: theme.secondary },
      line: { color: theme.secondary, transparency: 100 },
    });
    return;
  }

  if (design.style === "editorial") {
    slide.addShape(pptx.ShapeType.line, {
      x: density.outerX,
      y: density.contentTop - 0.07,
      w: W - density.outerX * 2,
      h: 0,
      line: { color: theme.accent, width: 1.15, transparency: 18 },
    });
    return;
  }

  if (design.style === "bold") {
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 0.16,
      h: H,
      fill: { color: index % 2 === 0 ? theme.secondary : theme.accent },
      line: { color: theme.secondary, transparency: 100 },
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: density.outerX,
      y: density.contentTop - 0.09,
      w: 2.3,
      h: 0.08,
      fill: { color: theme.secondary },
      line: { color: theme.secondary, transparency: 100 },
    });
    return;
  }

  if (design.style === "technical") {
    for (let lineIndex = 0; lineIndex < 4; lineIndex += 1) {
      const y = density.contentTop + lineIndex * ((density.contentBottom - density.contentTop) / 3);
      slide.addShape(pptx.ShapeType.line, {
        x: density.outerX,
        y,
        w: W - density.outerX * 2,
        h: 0,
        line: { color: theme.primary, width: 0.35, transparency: 91 },
      });
    }
    return;
  }

  slide.addShape(pptx.ShapeType.ellipse, {
    x: density.outerX,
    y: density.contentTop - 0.08,
    w: 0.13,
    h: 0.13,
    fill: { color: theme.accent },
    line: { color: theme.accent, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: density.outerX + 0.2,
    y: density.contentTop - 0.08,
    w: 0.13,
    h: 0.13,
    fill: { color: theme.secondary },
    line: { color: theme.secondary, transparency: 100 },
  });
}

function background(slide, color) {
  slide.background = { color };
}

function footer(slide, index, inverse) {
  if (templateChrome) {
    return;
  }

  const color = inverse ? theme.onPrimary : theme.muted;
  slide.addText(spec.title, {
    x: 0.7, y: 7.12, w: 8.6, h: 0.18,
    fontFace: theme.bodyFont, fontSize: 7.5, color, transparency: 36, margin: 0,
  });
  slide.addText(`${index + 1} / ${totalSlides}`, {
    x: 11.6, y: 7.08, w: 1.03, h: 0.22,
    fontFace: theme.bodyFont, fontSize: 8.5, bold: true,
    color, transparency: 22, margin: 0, align: "right",
  });
}

function card(slide, x, y, w, h, color) {
  const shape = !usesModernRendererContract
    ? design.style === "technical" || currentDensityName === "detailed"
      ? pptx.ShapeType.rect
      : pptx.ShapeType.roundRect
    : styleProfile.cardShape === "rect" || currentDensityName === "detailed"
    ? pptx.ShapeType.rect
    : pptx.ShapeType.roundRect;
  const fillColor = usesModernRendererContract && color === "FFFFFF" ? styleProfile.surface : color;
  slide.addShape(shape, {
    x, y, w, h,
    rectRadius: 0.08,
    fill: { color: fillColor },
    line: {
      color: usesModernRendererContract ? styleProfile.borderColor : "D9E0E8",
      width: usesModernRendererContract
        ? density.cardBorderWidth * styleProfile.borderWidthScale
        : density.cardBorderWidth,
    },
    shadow: usesModernRendererContract
      ? !styleProfile.cardShadow || !density.cardShadow
        ? undefined
        : { type: "outer", color: "182230", opacity: design.style === "bold" ? 0.16 : 0.09, blur: 2, angle: 45, distance: 1 }
      : design.style === "editorial" || !density.cardShadow
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
    fontFace: theme.bodyFont, fontSize: 10.5, bold: true,
    color: usesModernRendererContract
      ? readableForeground(color)
      : luminance(color) > 0.56 ? theme.text : "FFFFFF",
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
    fontFace: theme.bodyFont, fontSize: 11, bold: true, color: textColor,
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
    fontFace: theme.bodyFont, fontSize: 9, bold: true,
    charSpacing: 1.2, color: theme.accentTextOnPrimary, margin: 0,
  });
  slide.addText(text, {
    x: x + 0.3, y: y + 0.93, w: w - 0.6, h: h - 1.35,
    fontFace: theme.headingFont, fontSize: 18, bold: true,
    color: theme.onPrimary, margin: 0, valign: "mid", fit: "shrink",
  });
}

function metricCard(slide, metric, x, y, w, h, index) {
  const tone = toneColor(metric.tone, index);
  const readableTone = readableToneOnSurface(tone);
  card(slide, x, y, w, h, "FFFFFF");
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + w - 0.78, y: y + 0.26, w: 0.42, h: 0.42,
    fill: { color: tone, transparency: 75 }, line: { color: tone, transparency: 100 },
  });
  slide.addText(metric.value, {
    x: x + 0.38, y: y + 0.34, w: w - 0.82, h: h > 2.5 ? 1.25 : 0.68,
    fontFace: theme.headingFont, fontSize: h > 2.5 ? 36 : 27, bold: true,
    color: readableTone, margin: 0, valign: "mid", fit: "shrink",
  });
  slide.addText(metric.label, {
    x: x + 0.4, y: y + (h > 2.5 ? 1.72 : 1.07), w: w - 0.8, h: 0.48,
    fontFace: theme.headingFont, fontSize: 15.5, bold: true,
    color: cardTextColor, margin: 0, fit: "shrink",
  });
  if (metric.detail) {
    slide.addText(metric.detail, {
      x: x + 0.4, y: y + (h > 2.5 ? 2.42 : 1.52), w: w - 0.8, h: h > 2.5 ? 0.72 : 0.32,
      fontFace: theme.bodyFont, fontSize: 11.5, color: cardMutedTextColor,
      margin: 0, fit: "shrink",
    });
  }
  slide.addText(String(index + 1).padStart(2, "0"), {
    x: x + w - 0.72, y: y + h - 0.44, w: 0.36, h: 0.22,
    fontFace: theme.bodyFont, fontSize: 8, bold: true,
    color: cardMutedTextColor, transparency: 25, margin: 0, align: "right",
  });
}

function processCard(slide, step, index, x, y, w, h) {
  card(slide, x, y, w, h, "FFFFFF");
  const label = step.label ?? String(index + 1).padStart(2, "0");
  const stepColor = index % 2 === 0 ? theme.secondary : theme.accent;
  if ([...label].length > 4) {
    pill(
      slide,
      label,
      x + 0.28,
      y + 0.28,
      Math.min(1.3, w - 0.56),
      stepColor,
      usesModernRendererContract ? readableForeground(stepColor) : "FFFFFF",
    );
  } else {
    badge(slide, label, x + 0.28, y + 0.28, 0.68, stepColor);
  }
  slide.addText(step.title, {
    x: x + 0.28, y: y + 1.15, w: w - 0.56, h: h > 2.5 ? 0.62 : 0.42,
    fontFace: theme.headingFont, fontSize: h > 2.5 ? ([...step.title].length > 8 ? 14.5 : 17) : 14.5, bold: true,
    color: cardTextColor, margin: 0, fit: "shrink",
  });
  if (step.description) {
    slide.addText(step.description, {
      x: x + 0.28, y: y + (h > 2.5 ? 1.95 : 1.5), w: w - 0.56, h: h > 2.5 ? 1.08 : 0.3,
      fontFace: theme.bodyFont, fontSize: h > 2.5 ? 12.5 : 10.5,
      color: cardMutedTextColor, margin: 0, valign: "top", fit: "shrink",
    });
  }
}

function smallMetricCard(slide, metric, x, y, w, h, index) {
  const tone = toneColor(metric.tone, index);
  const readableTone = readableToneOnSurface(tone);
  card(slide, x, y, w, h, "FFFFFF");
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + 0.22, y: y + 0.24, w: 0.42, h: 0.42,
    fill: { color: tone, transparency: 72 },
    line: { color: tone, transparency: 100 },
  });
  slide.addText(metric.value, {
    x: x + 0.78, y: y + 0.18, w: w - 1.02, h: 0.5,
    fontFace: theme.headingFont, fontSize: scaled(21), bold: true,
    color: readableTone, margin: 0, fit: "shrink",
  });
  slide.addText(metric.label, {
    x: x + 0.24, y: y + 0.78, w: w - 0.48, h: 0.3,
    fontFace: theme.headingFont, fontSize: scaled(10.5), bold: true,
    color: cardTextColor, margin: 0, fit: "shrink",
  });
  if (metric.detail) {
    slide.addText(metric.detail, {
      x: x + 0.24, y: y + 1.08, w: w - 0.48, h: 0.18,
      fontFace: theme.bodyFont, fontSize: 9, color: cardMutedTextColor,
      margin: 0, fit: "shrink",
    });
  }
}

function renderVisualCard(slide, item, x, y, w, h, index) {
  const tone = toneColor(item.tone, index);
  const readableTone = readableToneOnSurface(tone);
  const spacious = h > 2.5;
  const compact = h <= 1.6;
  card(slide, x, y, w, h, "FFFFFF");
  renderIcon(slide, item.icon, x + 0.28, y + (compact ? 0.34 : 0.28), compact ? 0.62 : spacious ? 0.88 : 0.66, tone, false);
  if (item.value) {
    slide.addText(item.value, {
      x: x + w - Math.min(2.1, w * 0.42), y: y + 0.29, w: Math.min(1.72, w * 0.36), h: 0.57,
      fontFace: theme.bodyFont, fontSize: scaled(spacious ? 22 : 17), bold: true,
      color: readableTone, margin: 0, align: "right", valign: "mid", fit: "shrink",
    });
  }
  if (compact) {
    slide.addText(item.title, {
      x: x + 1.12, y: y + 0.2, w: w - (item.value ? 3.0 : 1.42), h: 0.38,
      fontFace: theme.headingFont, fontSize: scaled(14), bold: true,
      color: cardTextColor, margin: 0, fit: "shrink",
    });
    if (item.description) {
      slide.addText(item.description, {
        x: x + 1.12, y: y + 0.68, w: w - 1.42, h: h - 0.84,
        fontFace: theme.bodyFont, fontSize: scaled(9.5), color: cardMutedTextColor,
        margin: 0, fit: "shrink",
      });
    }
    return;
  }

  slide.addText(item.title, {
    x: x + 0.3, y: y + (spacious ? 1.47 : 1.04), w: w - 0.6, h: spacious ? 0.62 : 0.42,
    fontFace: theme.headingFont, fontSize: scaled(spacious ? 19 : 15), bold: true,
    color: cardTextColor, margin: 0, fit: "shrink",
  });
  if (item.description) {
    slide.addText(item.description, {
      x: x + 0.3, y: y + (spacious ? 2.28 : 1.52), w: w - 0.6, h: spacious ? h - 2.65 : Math.max(0.3, h - 1.7),
      fontFace: theme.bodyFont, fontSize: scaled(spacious ? 13 : 10.5),
      color: cardMutedTextColor, margin: 0, valign: "top", fit: "shrink",
    });
  }
}

function addNativeBulletList(slide, items, x, y, w, h, fontSize, verticalAlign = "mid", textColor = theme.text) {
  slide.addText((items ?? []).map((item, itemIndex) => ({
    text: item,
    options: {
      bullet: true,
      breakLine: itemIndex < items.length - 1,
      paraSpaceAfter: currentDensityName === "detailed" ? 4 : 7,
      fontSize: scaled(fontSize),
      color: textColor,
    },
  })), {
    x, y, w, h,
    fontFace: theme.bodyFont, color: textColor,
    margin: 0.03, valign: verticalAlign, fit: "shrink",
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
  if (currentDensityName === "detailed" && subtle) return;
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
    positive: usesModernRendererContract ? theme.positive : "159D73",
    success: usesModernRendererContract ? theme.positive : "159D73",
    warning: usesModernRendererContract ? theme.warning : "E89222",
    critical: usesModernRendererContract ? theme.critical : "D92D20",
    danger: usesModernRendererContract ? theme.critical : "D92D20",
    negative: usesModernRendererContract ? theme.critical : "D92D20",
    risk: usesModernRendererContract ? theme.critical : "D92D20",
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

function mixHex(base, overlay, overlayRatio) {
  const ratio = Math.max(0, Math.min(1, overlayRatio));
  const component = (hex, offset) => Number.parseInt(hex.slice(offset, offset + 2), 16);
  return [0, 2, 4]
    .map((offset) => Math.round(component(base, offset) * (1 - ratio) + component(overlay, offset) * ratio)
      .toString(16)
      .padStart(2, "0"))
    .join("")
    .toUpperCase();
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

function readableForeground(fillColor) {
  const dark = "17213A";
  const light = "FFFFFF";
  return contrastRatio(dark, fillColor) >= contrastRatio(light, fillColor) ? dark : light;
}

function readableToneOnBackground(tone, backgroundColor, fallback) {
  if (!usesModernRendererContract || contrastRatio(tone, backgroundColor) >= 4.5) {
    return tone;
  }
  return contrastRatio(fallback, backgroundColor) >= 4.5
    ? fallback
    : readableForeground(backgroundColor);
}

function readableToneOnSurface(tone) {
  return !usesModernRendererContract || contrastRatio(tone, styleProfile.surface) >= 4.5
    ? tone
    : cardTextColor;
}
