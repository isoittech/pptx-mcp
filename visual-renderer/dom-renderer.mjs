import { writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { exportHtmlToPptx } from "dom-to-pptx/node";
import { renderApprovedReactIcon } from "./react-icons.mjs";

const SLIDE_WIDTH_INCHES = 13.333333;
const SLIDE_HEIGHT_INCHES = 7.5;
const VIEWPORT_WIDTH = 1600;
const VIEWPORT_HEIGHT = 900;

const presets = Object.freeze({
  midnight: ["14213D", "2D5BFF", "2ED3C6", "F5F7FB", "17213A", "667085"],
  aurora: ["312E81", "7C3AED", "14B8A6", "F7F5FF", "201A3D", "706A86"],
  sunset: ["3A1F3D", "F05A47", "F5B642", "FFF8F3", "2B1B2E", "7C687C"],
  forest: ["153B32", "207A63", "A7C957", "F5F8F4", "18332C", "64746F"],
  minimal: ["202124", "4F6BED", "FFB000", "F7F7F5", "202124", "6B6F76"],
  ocean: ["063970", "087CA7", "17C3B2", "F2F8FB", "102A43", "627D98"],
  berry: ["4A1942", "893168", "E8C547", "FBF7FA", "2E172B", "806579"],
  clay: ["5E3023", "C08552", "F3A712", "FCF8F2", "35251E", "806E65"],
  cyber: ["0B1020", "4F46E5", "22D3EE", "F3F6FF", "111827", "667085"],
});

export const domSupportedSlideKinds = Object.freeze(new Set([
  "title",
  "agenda",
  "section",
  "bullets",
  "metrics",
  "comparison",
  "process",
  "timeline",
  "statement",
  "cards",
  "matrix",
  "funnel",
  "roadmap",
  "quote",
  "closing",
  "structuredbrief",
  "structured_brief",
  "scorecard",
  "datatable",
  "data_table",
  "media",
  "coveragemap",
  "coverage_map",
  "transformationevidence",
  "transformation_evidence",
  "artifactshowcase",
  "artifact_showcase",
  "ganttschedule",
  "gantt_schedule",
]));

export function canRenderDeckWithDom(spec) {
  return Array.isArray(spec?.slides)
    && spec.slides.length > 0
    && spec.slides.every((slide) => domSupportedSlideKinds.has(normalizeKind(slide.kind)));
}

export async function renderDomDeck(spec, outputPath, imageAssets = {}) {
  if (!canRenderDeckWithDom(spec)) {
    throw new Error("The deck contains a slide kind that is not supported by the DOM renderer.");
  }

  const html = buildDomDeckHtml(spec, imageAssets);
  const htmlPath = join(dirname(outputPath), "visual-deck.html");
  await writeFile(htmlPath, html, { encoding: "utf8", flag: "wx" });
  const buffer = await exportHtmlToPptx(htmlPath, {
    selector: ".slide",
    injectBundle: true,
    browserWidth: VIEWPORT_WIDTH,
    browserHeight: VIEWPORT_HEIGHT,
    pptxOptions: {
      author: "PowerPoint MCP",
      title: safeText(spec.title),
      subject: safeText(spec.subject),
      width: SLIDE_WIDTH_INCHES,
      height: SLIDE_HEIGHT_INCHES,
      includePseudoElements: false,
      svgAsVector: true,
    },
  });
  await writeFile(outputPath, buffer, { flag: "wx" });
}

export function buildDomDeckHtml(spec, imageAssets = {}) {
  const theme = {
    ...resolveTheme(spec.theme),
    visualObjectAssets: Array.isArray(spec.visual_object_assets) ? spec.visual_object_assets : [],
  };
  const templateChrome = spec.templateChrome === true;
  const slideNumberOffset = Number.isInteger(spec.slideNumberOffset) && spec.slideNumberOffset >= 0
    ? spec.slideNumberOffset
    : 0;
  const deckTotalSlides = Number.isInteger(spec.deckTotalSlides) && spec.deckTotalSlides >= spec.slides.length
    ? spec.deckTotalSlides
    : spec.slides.length;
  const slides = spec.slides.map((slide, index) => renderSlide(
    slide,
    slideNumberOffset + index,
    deckTotalSlides,
    theme,
    templateChrome,
    imageAssets,
  )).join("\n");

  return `<!doctype html>
<html lang="${escapeAttribute(spec.language ?? "ja-JP")}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=${VIEWPORT_WIDTH},initial-scale=1">
<style>${baseCss(theme)}</style>
</head>
<body>${slides}</body>
</html>`;
}

function renderSlide(slide, index, totalSlides, theme, templateChrome, imageAssets) {
  const kind = normalizeKind(slide.kind);
  const content = renderers[kind](slide, index, theme, imageAssets);
  const notes = slide.speakerNotes
    ? `このスライドの狙い\n${safeText(slide.speakerNotes.purpose)}\n\nトークスクリプト\n${safeText(slide.speakerNotes.talkScript)}`
    : "";
  const backgroundClass = templateChrome ? " template-chrome" : "";
  const eyebrow = slide.eyebrow
    ? `<div class="eyebrow">${escapeHtml(slide.eyebrow)}</div>`
    : "";
  const footer = templateChrome || kind === "title" || kind === "section" || kind === "closing"
    ? ""
    : `<div class="footer"><span>${escapeHtml(safeText(slide.attribution))}</span><span>${index + 1} / ${totalSlides}</span></div>`;

  return `<section class="slide${backgroundClass}">
    ${notes ? `<template data-pptx-notes>${escapeHtml(notes)}</template>` : ""}
    <div class="accent-rail"></div>
    <header>${eyebrow}<h1>${escapeHtml(slide.title)}</h1>${slide.subtitle ? `<p>${escapeHtml(slide.subtitle)}</p>` : ""}</header>
    <main class="kind-${escapeAttribute(kind)}">${content}</main>
    ${renderPreparedVisualObjects(slide, index, theme)}
    ${footer}
  </section>`;
}

const renderers = Object.freeze({
  title: renderTitle,
  agenda: renderAgenda,
  section: renderSection,
  bullets: renderBullets,
  metrics: renderMetrics,
  comparison: renderComparison,
  process: renderProcess,
  timeline: renderTimeline,
  statement: renderStatement,
  cards: renderCards,
  matrix: renderMatrix,
  funnel: renderFunnel,
  roadmap: renderRoadmap,
  quote: renderQuote,
  closing: renderClosing,
  structuredbrief: renderStructuredBrief,
  structured_brief: renderStructuredBrief,
  scorecard: renderScorecard,
  datatable: renderDataTable,
  data_table: renderDataTable,
  media: renderMedia,
  coveragemap: renderCoverageMap,
  coverage_map: renderCoverageMap,
  transformationevidence: renderTransformationEvidence,
  transformation_evidence: renderTransformationEvidence,
  artifactshowcase: renderArtifactShowcase,
  artifact_showcase: renderArtifactShowcase,
  ganttschedule: renderGanttSchedule,
  gantt_schedule: renderGanttSchedule,
});

function renderTitle(slide) {
  const summary = slide.body ?? slide.takeaway ?? slide.subtitle ?? "";
  return `<div class="hero-grid">
    <div class="hero-copy">${summary ? `<p>${escapeHtml(summary)}</p>` : ""}</div>
    <div class="hero-mark"><div class="icon-shell">${renderApprovedReactIcon("insight", { color: "var(--accent)" })}</div></div>
  </div>`;
}

function renderAgenda(slide) {
  const items = normalizedItems(slide.bullets);
  return `<div class="agenda-list">${items.map((item, index) => `<div class="agenda-row"><span>${String(index + 1).padStart(2, "0")}</span><p>${escapeHtml(item)}</p></div>`).join("")}</div>`;
}

function renderSection(slide, index) {
  return `<div class="section-stage"><div class="section-number">${String(index + 1).padStart(2, "0")}</div><p>${escapeHtml(slide.body ?? slide.subtitle ?? "")}</p></div>`;
}

function renderBullets(slide) {
  const items = normalizedItems(slide.bullets);
  const split = String(slide.variant ?? "").toLowerCase() === "split" && items.length >= 4;
  const body = `<div class="bullet-list${split ? " split" : ""}">${items.map((item) => `<div class="bullet-row"><span></span><p>${escapeHtml(item)}</p></div>`).join("")}</div>`;
  return `${slide.body ? `<p class="lead">${escapeHtml(slide.body)}</p>` : ""}${body}${takeaway(slide.takeaway)}`;
}

function renderMetrics(slide) {
  const metrics = Array.isArray(slide.metrics) ? slide.metrics : [];
  const spotlight = String(slide.variant ?? "").toLowerCase() === "spotlight" && metrics.length >= 2;
  return `<div class="metric-grid count-${metrics.length}${spotlight ? " spotlight" : ""}">${metrics.map((metric, index) => `<div class="metric-card tone-${toneClass(metric.tone)}${spotlight && index === 0 ? " featured" : ""}"><div class="metric-value">${escapeHtml(metric.value)}</div><div class="metric-label">${escapeHtml(metric.label)}</div>${metric.detail ? `<p>${escapeHtml(metric.detail)}</p>` : ""}</div>`).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderComparison(slide) {
  const panels = Array.isArray(slide.panels) ? slide.panels : [];
  return `<div class="panel-grid count-${Math.min(panels.length, 3)}">${panels.map((panel, index) => `<article class="panel"><div class="panel-index">${String(index + 1).padStart(2, "0")}</div><h2>${escapeHtml(panel.title)}</h2>${panel.highlight ? `<strong>${escapeHtml(panel.highlight)}</strong>` : ""}<div class="compact-list">${normalizedItems(panel.bullets).map((item) => `<p>${escapeHtml(item)}</p>`).join("")}</div></article>`).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderProcess(slide) {
  const steps = Array.isArray(slide.steps) ? slide.steps : [];
  const loop = String(slide.variant ?? "").toLowerCase() === "loop";
  return `<div class="step-flow${loop ? " loop" : ""}">${steps.map((step, index) => `${index > 0 ? `<div class="connector">→</div>` : ""}<article class="step"><span>${escapeHtml(step.label ?? String(index + 1).padStart(2, "0"))}</span><h2>${escapeHtml(step.title)}</h2>${step.description ? `<p>${escapeHtml(step.description)}</p>` : ""}</article>`).join("")}</div>${loop ? `<div class="loop-label">↻ ITERATE</div>` : ""}${takeaway(slide.takeaway)}`;
}

function renderTimeline(slide) {
  const steps = Array.isArray(slide.steps) ? slide.steps : [];
  return `<div class="timeline" style="grid-template-columns:repeat(${steps.length},1fr)"><div class="timeline-line"></div>${steps.map((step, index) => `<article class="timeline-step"><span>${escapeHtml(step.label ?? String(index + 1))}</span><div></div><h2>${escapeHtml(step.title)}</h2>${step.description ? `<p>${escapeHtml(step.description)}</p>` : ""}</article>`).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderStatement(slide) {
  return `<div class="statement"><div class="statement-icon">${renderApprovedReactIcon("insight", { color: "var(--accent)" })}</div><p>${escapeHtml(slide.body ?? slide.subtitle ?? "")}</p>${slide.takeaway ? `<strong>${escapeHtml(slide.takeaway)}</strong>` : ""}</div>`;
}

function renderCards(slide) {
  const cards = Array.isArray(slide.cards) ? slide.cards : [];
  return `<div class="card-grid count-${Math.min(cards.length, 4)}">${cards.map((card) => `<article class="visual-card tone-${toneClass(card.tone)}"><div class="card-icon">${renderApprovedReactIcon(card.icon ?? "insight", { color: "var(--tone)" })}</div>${card.value ? `<strong>${escapeHtml(card.value)}</strong>` : ""}<h2>${escapeHtml(card.title)}</h2>${card.description ? `<p>${escapeHtml(card.description)}</p>` : ""}</article>`).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderMatrix(slide) {
  const matrix = slide.matrix ?? {};
  const quadrants = Array.isArray(matrix.quadrants) ? matrix.quadrants.slice(0, 4) : [];
  return `<div class="matrix-wrap"><div class="matrix-y">${escapeHtml(matrix.verticalAxis ?? "")}</div><div class="matrix-grid">${quadrants.map((quadrant) => `<article><h2>${escapeHtml(quadrant.title)}</h2>${quadrant.highlight ? `<strong>${escapeHtml(quadrant.highlight)}</strong>` : ""}${normalizedItems(quadrant.bullets).map((item) => `<p>${escapeHtml(item)}</p>`).join("")}</article>`).join("")}</div><div class="matrix-x">${escapeHtml(matrix.horizontalAxis ?? "")}</div></div>${takeaway(slide.takeaway)}`;
}

function renderFunnel(slide) {
  const steps = Array.isArray(slide.steps) ? slide.steps : [];
  return `<div class="funnel">${steps.map((step, index) => `<div style="width:${100 - index * Math.min(12, 44 / Math.max(1, steps.length - 1))}%"><span>${escapeHtml(step.label ?? String(index + 1))}</span><strong>${escapeHtml(step.title)}</strong>${step.description ? `<p>${escapeHtml(step.description)}</p>` : ""}</div>`).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderRoadmap(slide) {
  const steps = Array.isArray(slide.steps) ? slide.steps : [];
  return `<div class="roadmap" style="grid-template-columns:repeat(${steps.length},1fr)">${steps.map((step, index) => `<article><span>${escapeHtml(step.label ?? `PHASE ${index + 1}`)}</span><h2>${escapeHtml(step.title)}</h2>${step.description ? `<p>${escapeHtml(step.description)}</p>` : ""}</article>`).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderQuote(slide) {
  return `<div class="quote"><div>“</div><blockquote>${escapeHtml(slide.body ?? slide.subtitle ?? "")}</blockquote>${slide.attribution ? `<p>— ${escapeHtml(slide.attribution)}</p>` : ""}</div>`;
}

function renderClosing(slide) {
  const items = normalizedItems(slide.bullets).slice(0, 4);
  return `<div class="closing"><div class="closing-icon">${renderApprovedReactIcon("check", { color: "var(--accent)" })}</div>${slide.body ? `<p>${escapeHtml(slide.body)}</p>` : ""}${items.length ? `<div>${items.map((item) => `<span>${escapeHtml(item)}</span>`).join("")}</div>` : ""}</div>`;
}

function renderStructuredBrief(slide) {
  const sections = Array.isArray(slide.sections) ? slide.sections : [];
  return `<div class="brief-grid count-${sections.length}">${sections.map((section, index) => `<article class="tone-${toneClass(section.tone)}"><span>${String(index + 1).padStart(2, "0")}</span><h2>${escapeHtml(section.heading)}</h2>${section.highlight ? `<strong>${escapeHtml(section.highlight)}</strong>` : ""}${section.body ? `<p>${escapeHtml(section.body)}</p>` : ""}${normalizedItems(section.bullets).map((item) => `<div class="brief-bullet">${escapeHtml(item)}</div>`).join("")}</article>`).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderScorecard(slide) {
  const scorecard = slide.scorecard ?? {};
  const options = Array.isArray(scorecard.options) ? scorecard.options : [];
  const rows = Array.isArray(scorecard.criteria) ? scorecard.criteria : [];
  const columns = `1.15fr repeat(${options.length}, 1fr)`;
  return `<div class="data-grid scorecard" style="grid-template-columns:${columns}"><div class="table-head"></div>${options.map((option) => `<div class="table-head"><strong>${escapeHtml(option.title)}</strong>${option.subtitle ? `<span>${escapeHtml(option.subtitle)}</span>` : ""}</div>`).join("")}${rows.map((row) => `<div class="row-head">${escapeHtml(row.criterion)}</div>${(row.cells ?? []).map((cell) => `<div class="table-cell tone-${toneClass(cell.tone)}"><strong>${escapeHtml(cell.rating)}</strong>${cell.detail ? `<span>${escapeHtml(cell.detail)}</span>` : ""}</div>`).join("")}`).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderDataTable(slide) {
  const table = slide.dataTable ?? {};
  const columns = Array.isArray(table.columns) ? table.columns : [];
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const weights = columns.map((column) => `${Number(column.widthWeight ?? 1)}fr`).join(" ");
  return `<div class="data-grid" style="grid-template-columns:${escapeAttribute(weights)}">${columns.map((column) => `<div class="table-head align-${alignment(column.align)}">${escapeHtml(column.header)}</div>`).join("")}${rows.map((row) => (row.cells ?? []).map((cell, index) => `<div class="table-cell align-${alignment(columns[index]?.align)} tone-${toneClass(cell.tone)}${cell.emphasize ? " emphasized" : ""}">${escapeHtml(cell.text)}</div>`).join("")).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderMedia(slide, _index, _theme, imageAssets) {
  const media = slide.media ?? {};
  const asset = imageAssets[media.assetId];
  if (!asset) {
    throw new Error("The media slide references an unavailable server-owned image asset.");
  }
  const textFirst = String(media.textPosition ?? "left").toLowerCase() !== "right";
  const copy = `<div class="media-copy">${slide.body ? `<p>${escapeHtml(slide.body)}</p>` : ""}${normalizedItems(slide.bullets).map((item) => `<div>${escapeHtml(item)}</div>`).join("")}${media.caption ? `<small>${escapeHtml(media.caption)}</small>` : ""}</div>`;
  const image = `<figure><img src="${escapeAttribute(asset.data)}" alt="${escapeAttribute(asset.altText)}" style="object-fit:${String(media.cropIntent).toLowerCase() === "contain" ? "contain" : "cover"}"></figure>`;
  return `<div class="media-grid">${textFirst ? `${copy}${image}` : `${image}${copy}`}</div>${takeaway(slide.takeaway)}`;
}

function renderCoverageMap(slide) {
  const coverage = slide.coverageMap ?? {};
  const columns = Array.isArray(coverage.columns) ? coverage.columns : [];
  const groups = Array.isArray(coverage.groups) ? coverage.groups : [];
  const bars = Array.isArray(coverage.bars) ? coverage.bars : [];
  const flattenedRows = groups.flatMap((group) => (group.rows ?? []).map((row) => ({ group, row })));
  const rowNumbers = new Map(flattenedRows.map(({ row }, index) => [row.id, index + 2]));
  const groupStarts = new Map();
  let cursor = 2;
  for (const group of groups) {
    groupStarts.set(group.id, cursor);
    cursor += Array.isArray(group.rows) ? group.rows.length : 0;
  }
  const cells = flattenedRows.flatMap(({ row }) => columns.map((_, columnIndex) => `<div class="coverage-cell" style="grid-row:${rowNumbers.get(row.id)};grid-column:${columnIndex + 3}"></div>`)).join("");
  const groupLabels = groups.map((group) => `<div class="coverage-group tone-${toneClass(group.tone)}" style="${toneStyle(group.tone)}grid-row:${groupStarts.get(group.id)} / span ${group.rows.length};grid-column:1"><strong>${escapeHtml(group.label)}</strong>${group.subtitle ? `<span>${escapeHtml(group.subtitle)}</span>` : ""}</div>`).join("");
  const rowLabels = flattenedRows.map(({ row }) => `<div class="coverage-row-label" style="grid-row:${rowNumbers.get(row.id)};grid-column:2">${escapeHtml(row.label)}</div>`).join("");
  const spanBars = bars.map((bar) => `<div class="coverage-bar tone-${toneClass(bar.tone)}" data-local-id="${escapeAttribute(bar.id)}" style="${toneStyle(bar.tone)}grid-row:${rowNumbers.get(bar.rowId)};grid-column:${bar.startColumn + 2} / ${bar.endColumn + 3}">${escapeHtml(bar.label)}</div>`).join("");
  const header = columns.map((column, index) => `<div class="coverage-axis" style="grid-row:1;grid-column:${index + 3}">${column.groupLabel ? `<span>${escapeHtml(column.groupLabel)}</span>` : ""}<strong>${escapeHtml(column.label)}</strong></div>`).join("");
  return `<div class="coverage-stage">
    <div class="coverage-grid" style="--axis-count:${columns.length};--row-count:${flattenedRows.length}">
      <div class="coverage-corner" style="grid-row:1;grid-column:1 / 3">適用範囲</div>${header}${cells}${groupLabels}${rowLabels}${spanBars}
    </div>
    ${coverage.callout ? `<aside class="coverage-callout tone-${toneClass(coverage.callout.tone)}" style="${toneStyle(coverage.callout.tone)}">${escapeHtml(coverage.callout.text)}</aside>` : ""}
    ${renderChips(coverage.footerChips)}
  </div>${takeaway(slide.takeaway)}`;
}

function renderTransformationEvidence(slide) {
  const evidence = slide.transformationEvidence ?? {};
  const segments = Array.isArray(evidence.inputSegments) ? evidence.inputSegments : [];
  return `<div class="transformation-grid">
    <section class="evidence-input"><h2>${escapeHtml(evidence.inputHeading)}</h2>${evidence.inputCaption ? `<small>${escapeHtml(evidence.inputCaption)}</small>` : ""}<div class="tagged-copy">${segments.map((segment) => `<span class="tagged-segment tone-${toneClass(segment.tone)}" style="${toneStyle(segment.tone)}">${segment.tag ? `<b>${escapeHtml(segment.tag)}</b>` : ""}${escapeHtml(segment.text)}</span>`).join("")}</div></section>
    <div class="transformation-arrow" aria-hidden="true">${renderApprovedReactIcon("arrow", { color: "var(--secondary)" })}</div>
    <section class="evidence-output"><h2>${escapeHtml(evidence.outputHeading)}</h2><div class="output-copy">${escapeHtml(evidence.outputText)}</div><h3>検証根拠</h3>${renderDataTableMarkup(evidence.evidenceTable)}</section>
  </div>${takeaway(slide.takeaway)}`;
}

function renderArtifactShowcase(slide, _index, _theme, imageAssets) {
  const showcase = slide.artifactShowcase ?? {};
  const groups = Array.isArray(showcase.groups) ? showcase.groups : [];
  return `<div class="artifact-grid count-${groups.length}">${groups.map((group) => {
    const artifacts = Array.isArray(group.artifacts) ? group.artifacts : [];
    const previews = artifacts.map((artifact, index) => {
      const asset = imageAssets[artifact.assetId];
      if (!asset) throw new Error("The artifact showcase references an unavailable server-owned image asset.");
      return `<figure class="artifact-preview preview-${Math.min(index + 1, 4)}"><img src="${escapeAttribute(asset.data)}" alt="${escapeAttribute(asset.altText)}" style="object-fit:${String(artifact.cropIntent).toLowerCase() === "cover" ? "cover" : "contain"}">${artifact.label ? `<figcaption>${escapeHtml(artifact.label)}</figcaption>` : ""}</figure>`;
    }).join("");
    return `<article class="artifact-group"><h2>${escapeHtml(group.title)}</h2>${group.description ? `<p>${escapeHtml(group.description)}</p>` : ""}<div class="artifact-stack count-${artifacts.length}">${previews}</div></article>`;
  }).join("")}</div>${takeaway(slide.takeaway)}`;
}

function renderGanttSchedule(slide) {
  const gantt = slide.ganttSchedule ?? {};
  const columns = Array.isArray(gantt.columns) ? gantt.columns : [];
  const tasks = Array.isArray(gantt.tasks) ? gantt.tasks : [];
  const markers = Array.isArray(gantt.markers) ? gantt.markers : [];
  const header = columns.map((column, index) => `<div class="gantt-axis" style="grid-row:1;grid-column:${index + 3}">${column.groupLabel ? `<span>${escapeHtml(column.groupLabel)}</span>` : ""}<strong>${escapeHtml(column.label)}</strong></div>`).join("");
  const cells = tasks.flatMap((_, rowIndex) => columns.map((__, columnIndex) => `<div class="gantt-cell" style="grid-row:${rowIndex + 2};grid-column:${columnIndex + 3}"></div>`)).join("");
  const markerRanges = markers.map((marker) => `<div class="gantt-marker tone-${toneClass(marker.tone)}" style="${toneStyle(marker.tone)}grid-row:2 / ${tasks.length + 2};grid-column:${marker.startColumn + 2} / ${marker.endColumn + 3}"><span>${escapeHtml(marker.label)}</span></div>`).join("");
  const taskRows = tasks.map((task, index) => `<div class="gantt-category" style="grid-row:${index + 2};grid-column:1">${escapeHtml(task.category)}</div><div class="gantt-task" style="grid-row:${index + 2};grid-column:2"><strong>${escapeHtml(task.title)}</strong>${normalizedItems(task.details).length ? `<span>${normalizedItems(task.details).map(escapeHtml).join("／")}</span>` : ""}</div><div class="gantt-bar tone-${toneClass(task.tone)}" data-local-id="${escapeAttribute(task.id)}" style="${toneStyle(task.tone)}grid-row:${index + 2};grid-column:${task.startColumn + 2} / ${task.endColumn + 3}"></div>`).join("");
  return `<div class="gantt-stage">${gantt.effortLabel ? `<div class="effort-label">${escapeHtml(gantt.effortLabel)}</div>` : ""}<div class="gantt-grid" style="--axis-count:${columns.length};--task-count:${tasks.length}"><div class="gantt-head" style="grid-row:1;grid-column:1">カテゴリ</div><div class="gantt-head" style="grid-row:1;grid-column:2">主なタスク</div>${header}${cells}${markerRanges}${taskRows}</div>${renderChips(gantt.legend)}</div>${takeaway(slide.takeaway)}`;
}

function renderDataTableMarkup(table) {
  const columns = Array.isArray(table?.columns) ? table.columns : [];
  const rows = Array.isArray(table?.rows) ? table.rows : [];
  const weights = columns.map((column) => `${Number(column.widthWeight ?? 1)}fr`).join(" ");
  return `<div class="data-grid compact-evidence-table" style="grid-template-columns:${escapeAttribute(weights)}">${columns.map((column) => `<div class="table-head align-${alignment(column.align)}">${escapeHtml(column.header)}</div>`).join("")}${rows.map((row) => (row.cells ?? []).map((cell, index) => `<div class="table-cell align-${alignment(columns[index]?.align)} tone-${toneClass(cell.tone)}${cell.emphasize ? " emphasized" : ""}">${escapeHtml(cell.text)}</div>`).join("")).join("")}</div>`;
}

function renderChips(chips) {
  const safeChips = Array.isArray(chips) ? chips : [];
  if (safeChips.length === 0) return "";
  return `<div class="visual-chips">${safeChips.map((chip) => `<span class="tone-${toneClass(chip.tone)}" style="${toneStyle(chip.tone)}">${escapeHtml(chip.label)}</span>`).join("")}</div>`;
}

function renderPreparedVisualObjects(slide, index, theme) {
  const objects = Array.isArray(slide.visualObjects) ? slide.visualObjects : [];
  if (objects.length === 0 || !Array.isArray(theme.visualObjectAssets)) return "";
  const assetMap = new Map(theme.visualObjectAssets.map((asset) => [asset.asset_id, asset]));
  return objects.map((reference) => {
    const asset = assetMap.get(reference.assetId);
    const brief = asset?.brief;
    if (!brief || Number(brief.slideNumber) !== index + 1) return "";
    const recipe = String(brief.recipe ?? "auto").toLowerCase();
    const label = brief.label ? `<span>${escapeHtml(brief.label)}</span>` : "";
    const glyph = recipe === "directionalcue" ? "→" : "";
    return `<div class="visual-object recipe-${escapeAttribute(recipe)}">${glyph}${label}</div>`;
  }).join("");
}

function takeaway(value) {
  return value ? `<div class="takeaway"><span>KEY TAKEAWAY</span><p>${escapeHtml(value)}</p></div>` : "";
}

function resolveTheme(input = {}) {
  const preset = presets[String(input.preset ?? "midnight").toLowerCase()] ?? presets.midnight;
  const background = color(input.backgroundColor, preset[3]);
  const muted = color(input.mutedTextColor, preset[5]);
  return {
    primary: color(input.primaryColor, preset[0]),
    secondary: color(input.secondaryColor, preset[1]),
    accent: color(input.accentColor, preset[2]),
    background,
    text: color(input.textColor, preset[4]),
    muted,
    footer: relativeLuminance(background) < 0.34 ? "F5F7FB" : muted,
    surface: color(input.surfaceColor, "FFFFFF"),
    positive: color(input.positiveColor, "159D73"),
    warning: color(input.warningColor, "E89222"),
    critical: color(input.criticalColor, "D92D20"),
    headingFont: safeFont(input.headingFontFace ?? input.fontFace ?? "Noto Sans CJK JP"),
    bodyFont: safeFont(input.bodyFontFace ?? input.fontFace ?? "Noto Sans CJK JP"),
  };
}

function baseCss(theme) {
  return `
.metric-grid.count-2{grid-template-columns:repeat(2,1fr)!important}.metric-grid.count-3{grid-template-columns:repeat(3,1fr)!important}.metric-grid.count-4{grid-template-columns:repeat(4,1fr)!important}.metric-grid.spotlight.count-3{grid-template-columns:1.55fr 1fr 1fr!important}.metric-grid.spotlight.count-4{grid-template-columns:1.55fr repeat(3,1fr)!important}.card-grid.count-2{grid-template-columns:repeat(2,1fr)!important}
*{box-sizing:border-box;box-shadow:none!important;text-shadow:none!important}html,body{margin:0;padding:0;background:transparent}body{font-family:"${theme.bodyFont}",sans-serif;color:#${theme.text}}
.slide{--primary:#${theme.primary};--secondary:#${theme.secondary};--accent:#${theme.accent};--bg:#${theme.background};--surface:#${theme.surface};--text:#${theme.text};--muted:#${theme.muted};--footer:#${theme.footer};--positive:#${theme.positive};--warning:#${theme.warning};--critical:#${theme.critical};position:relative;width:${VIEWPORT_WIDTH}px;height:${VIEWPORT_HEIGHT}px;overflow:hidden;background:var(--bg);padding:70px 96px 62px}.slide.template-chrome{background:transparent}.accent-rail{position:absolute;left:0;top:0;width:18px;height:100%;background:var(--accent)}
header{height:154px;position:relative;z-index:2}header h1{font-family:"${theme.headingFont}",sans-serif;font-size:50px;line-height:1.08;letter-spacing:-1.2px;margin:8px 0 0;color:var(--primary);max-width:1320px}header p{font-size:22px;line-height:1.35;color:var(--muted);margin:12px 0 0}.eyebrow{font-size:17px;line-height:1;text-transform:uppercase;letter-spacing:2.4px;font-weight:700;color:var(--secondary)}main{height:588px;position:relative;z-index:2}.footer{position:absolute;left:96px;right:96px;bottom:24px;display:flex;justify-content:space-between;font-size:14px;color:var(--footer)}.footer>span:first-child{flex:1}.footer>span:last-child{display:block;width:88px;text-align:right;white-space:nowrap}
.hero-grid{display:grid;grid-template-columns:2fr 1fr;gap:72px;height:100%;align-items:center}.hero-copy p{font-size:29px;line-height:1.45;max-width:860px}.hero-mark{height:360px;border-radius:40px;background:var(--surface);border:2px solid color-mix(in srgb,var(--secondary) 20%,transparent);display:flex;align-items:center;justify-content:center}.icon-shell{width:170px;height:170px}.agenda-list{display:grid;grid-template-columns:1fr 1fr;gap:22px 32px}.agenda-row{display:flex;gap:28px;align-items:center;padding:22px 28px;background:var(--surface);border-radius:22px;border:1px solid color-mix(in srgb,var(--muted) 22%,transparent)}.agenda-row span{font-size:24px;font-weight:800;color:var(--secondary)}.agenda-row p{font-size:25px;margin:0;font-weight:600}.section-stage{height:100%;display:flex;align-items:center;gap:56px}.section-number{font-size:190px;line-height:1;font-weight:800;color:var(--accent);letter-spacing:-10px}.section-stage p{font-size:32px;line-height:1.45;max-width:780px}.lead{font-size:24px;line-height:1.45;margin:0 0 28px}.bullet-list{display:grid;gap:18px}.bullet-list.split{grid-template-columns:1fr 1fr;gap:18px 28px}.bullet-row{display:flex;gap:18px;align-items:flex-start;background:var(--surface);padding:18px 24px;border-radius:18px}.bullet-row>span{width:12px;height:12px;border-radius:50%;background:var(--accent);margin-top:10px;flex:none}.bullet-row p{font-size:23px;line-height:1.4;margin:0}.metric-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:24px}.metric-grid.spotlight{grid-template-columns:1.55fr repeat(3,1fr)}.metric-card{--tone:var(--secondary);min-height:288px;padding:34px 30px;border-radius:26px;background:var(--surface);border-top:9px solid var(--tone)}.metric-card.featured{background:color-mix(in srgb,var(--tone) 9%,var(--surface))}.metric-value{font-size:60px;line-height:1;font-weight:800;color:var(--tone);letter-spacing:-2px}.metric-label{font-size:24px;font-weight:700;margin-top:24px}.metric-card p{font-size:18px;line-height:1.4;color:var(--muted)}
.tone-accent{--tone:var(--accent)}.tone-primary{--tone:var(--primary)}.tone-secondary{--tone:var(--secondary)}.tone-positive{--tone:var(--positive)}.tone-warning{--tone:var(--warning)}.tone-critical{--tone:var(--critical)}.tone-neutral{--tone:var(--muted)}.panel-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:28px}.panel-grid.count-3{grid-template-columns:repeat(3,1fr)}.panel{position:relative;min-height:370px;padding:36px;background:var(--surface);border-radius:28px;border:1px solid color-mix(in srgb,var(--muted) 20%,transparent)}.panel-index{font-size:18px;font-weight:800;color:var(--accent)}.panel h2{font-size:29px;margin:14px 0}.panel strong{display:block;font-size:22px;color:var(--secondary);margin-bottom:16px}.compact-list p{font-size:19px;line-height:1.4;margin:12px 0;padding-left:18px;border-left:4px solid var(--accent)}.step-flow{display:flex;align-items:center;justify-content:center;height:410px}.step{width:230px;min-height:250px;padding:28px 24px;background:var(--surface);border-radius:26px;text-align:center;border:1px solid color-mix(in srgb,var(--muted) 18%,transparent)}.step span{font-size:17px;font-weight:800;color:var(--accent)}.step h2{font-size:24px;margin:22px 0 10px}.step p{font-size:17px;line-height:1.4;color:var(--muted)}.connector{font-size:42px;color:var(--secondary);padding:0 16px}.loop-label{text-align:center;color:var(--secondary);font-size:18px;font-weight:800;letter-spacing:2px}.timeline{position:relative;display:grid;grid-template-columns:repeat(5,1fr);gap:12px;padding-top:110px}.timeline-line{position:absolute;left:7%;right:7%;top:148px;height:6px;background:var(--secondary);border-radius:3px}.timeline-step{text-align:center;position:relative}.timeline-step>span{font-size:17px;font-weight:800;color:var(--secondary)}.timeline-step>div{width:30px;height:30px;border:8px solid var(--accent);background:var(--surface);border-radius:50%;margin:21px auto 28px}.timeline-step h2{font-size:21px;margin:0 10px 12px}.timeline-step p{font-size:16px;line-height:1.35;color:var(--muted);margin:0 10px}.statement{height:100%;display:grid;grid-template-columns:180px 1fr;align-items:center;gap:52px}.statement-icon{width:150px;height:150px}.statement p{font-family:"${theme.headingFont}",sans-serif;font-size:45px;line-height:1.3;font-weight:700;margin:0;color:var(--primary)}.statement strong{display:block;font-size:22px;color:var(--secondary);margin-top:28px}.card-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:24px}.card-grid.count-4{grid-template-columns:repeat(4,1fr)}.visual-card{min-height:355px;padding:28px;border-radius:26px;background:var(--surface);border-bottom:8px solid var(--tone)}.card-icon{width:70px;height:70px;color:var(--tone)}.visual-card>strong{display:block;font-size:38px;margin-top:20px;color:var(--tone)}.visual-card h2{font-size:24px;margin:20px 0 10px}.visual-card p{font-size:17px;line-height:1.45;color:var(--muted)}.matrix-wrap{position:relative;padding:0 0 44px 55px}.matrix-grid{display:grid;grid-template-columns:1fr 1fr;grid-template-rows:1fr 1fr;height:430px;border-left:5px solid var(--primary);border-bottom:5px solid var(--primary)}.matrix-grid article{padding:22px 28px;background:var(--surface);border:1px solid color-mix(in srgb,var(--muted) 18%,transparent)}.matrix-grid h2{font-size:22px;margin:0 0 10px}.matrix-grid strong{color:var(--secondary)}.matrix-grid p{font-size:16px;margin:7px 0}.matrix-y{position:absolute;left:0;top:180px;writing-mode:vertical-rl;transform:rotate(180deg);font-size:18px;font-weight:700}.matrix-x{text-align:center;margin-top:16px;font-size:18px;font-weight:700}.funnel{display:flex;flex-direction:column;align-items:center;gap:8px}.funnel>div{min-height:72px;padding:13px 28px;background:var(--secondary);color:white;text-align:center;clip-path:polygon(5% 0,95% 0,100% 100%,0 100%)}.funnel span{font-size:14px;font-weight:800;margin-right:18px}.funnel strong{font-size:22px}.funnel p{display:inline;font-size:15px;margin-left:18px}.roadmap{display:grid;grid-template-columns:repeat(4,1fr);gap:22px;padding-top:58px}.roadmap article{min-height:330px;background:var(--surface);border-radius:26px;padding:30px;border-top:9px solid var(--secondary)}.roadmap span{font-size:15px;font-weight:800;color:var(--accent)}.roadmap h2{font-size:24px;margin:24px 0 14px}.roadmap p{font-size:17px;line-height:1.45;color:var(--muted)}.quote{height:100%;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center}.quote>div{font-size:120px;line-height:.5;color:var(--accent)}.quote blockquote{font-family:"${theme.headingFont}",sans-serif;font-size:42px;line-height:1.35;max-width:1120px;margin:42px 0 24px}.quote p{font-size:20px;color:var(--muted)}.closing{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center}.closing-icon{width:110px;height:110px}.closing>p{font-size:28px;max-width:900px}.closing>div:last-child{display:flex;gap:18px;margin-top:24px}.closing span{background:var(--surface);padding:13px 22px;border-radius:999px;font-size:17px}.brief-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:24px}.brief-grid.count-2{grid-template-columns:repeat(2,1fr)}.brief-grid article{min-height:380px;background:var(--surface);padding:28px;border-radius:24px;border-top:8px solid var(--tone)}.brief-grid article>span{font-size:16px;font-weight:800;color:var(--tone)}.brief-grid h2{font-size:24px;margin:14px 0}.brief-grid strong{display:block;color:var(--tone);font-size:19px;margin-bottom:12px}.brief-grid p,.brief-bullet{font-size:16px;line-height:1.45;color:var(--muted)}.brief-bullet{margin-top:9px;padding-left:13px;border-left:3px solid var(--tone)}.data-grid{display:grid;gap:2px;background:color-mix(in srgb,var(--muted) 20%,transparent);border:2px solid color-mix(in srgb,var(--muted) 20%,transparent);border-radius:16px;overflow:hidden}.table-head,.row-head,.table-cell{background:var(--surface);padding:16px 18px;min-height:54px;font-size:16px;display:flex;flex-direction:column;justify-content:center}.table-head{background:var(--primary);color:white;font-weight:700}.table-head span,.table-cell span{font-size:13px;line-height:1.35;margin-top:5px}.row-head{font-weight:700}.table-cell strong{color:var(--tone)}.table-cell.emphasized{font-weight:800}.align-center{text-align:center;align-items:center}.align-right{text-align:right;align-items:flex-end}.media-grid{display:grid;grid-template-columns:1fr 1.15fr;gap:34px;height:430px}.media-copy{background:var(--surface);border-radius:25px;padding:32px}.media-copy>p{font-size:22px;line-height:1.5;margin:0 0 24px}.media-copy>div{font-size:18px;line-height:1.4;margin:14px 0;padding-left:16px;border-left:4px solid var(--accent)}.media-copy small{display:block;margin-top:24px;color:var(--muted)}figure{margin:0;border-radius:28px;overflow:hidden;background:var(--surface)}figure img{width:100%;height:100%;display:block}.takeaway{position:absolute;left:0;right:0;bottom:4px;min-height:70px;display:flex;align-items:center;gap:22px;padding:13px 24px;background:color-mix(in srgb,var(--accent) 12%,var(--surface));border-left:8px solid var(--accent);border-radius:10px}.takeaway span{font-size:13px;font-weight:800;letter-spacing:1px;color:var(--secondary)}.takeaway p{flex:1;min-width:0;font-size:18px;line-height:1.35;margin:0}.visual-object{position:absolute;z-index:3;color:var(--accent)}.visual-object span{font-size:15px;font-weight:700}.recipe-sectionrule{left:96px;right:96px;bottom:112px;border-top:5px solid var(--secondary);padding-top:8px}.recipe-focuscorners{right:84px;top:220px;width:320px;height:280px;border:7px solid var(--accent);background:transparent}.recipe-directionalcue{left:50%;top:55%;font-size:40px}.recipe-directionalcue:after{content:"→"}.recipe-growthpath{right:110px;top:250px;width:160px;height:160px;border-top:8px solid var(--positive);transform:rotate(-28deg)}.recipe-cyclecue{right:95px;bottom:100px;width:110px;height:110px;border:8px solid var(--secondary);border-radius:50%}.recipe-annotationpin{right:120px;top:300px;background:var(--surface);border:3px solid var(--primary);border-radius:18px;padding:12px 18px}
.coverage-stage,.gantt-stage{position:relative;height:510px}.coverage-grid{display:grid;grid-template-columns:180px 230px repeat(var(--axis-count),minmax(0,1fr));grid-template-rows:62px repeat(var(--row-count),52px);position:relative;border:1px solid color-mix(in srgb,var(--muted) 28%,transparent);background:var(--surface)}.coverage-corner,.coverage-axis,.gantt-head,.gantt-axis{display:flex;align-items:center;justify-content:center;padding:8px 10px;border-right:1px solid color-mix(in srgb,var(--muted) 28%,transparent);border-bottom:1px solid color-mix(in srgb,var(--muted) 28%,transparent);font-size:23px;font-weight:700}.coverage-corner{justify-content:flex-start;background:var(--primary);color:white}.coverage-axis,.gantt-axis{flex-direction:column;background:color-mix(in srgb,var(--primary) 7%,var(--surface))}.coverage-axis span,.gantt-axis span{font-size:18px;color:var(--muted)}.coverage-group{--tone:var(--accent);display:flex;flex-direction:column;align-items:center;justify-content:center;padding:8px;background:var(--tone);color:white;border-bottom:1px solid var(--surface);font-size:23px;text-align:center;z-index:2}.coverage-group span{font-size:18px}.coverage-row-label{display:flex;align-items:center;padding:6px 14px;border-right:1px solid color-mix(in srgb,var(--muted) 28%,transparent);border-bottom:1px solid color-mix(in srgb,var(--muted) 28%,transparent);font-size:22px;font-weight:700;z-index:2;background:var(--surface)}.coverage-cell,.gantt-cell{border-right:1px solid color-mix(in srgb,var(--muted) 22%,transparent);border-bottom:1px solid color-mix(in srgb,var(--muted) 22%,transparent);z-index:1}.coverage-bar{--tone:var(--secondary);align-self:center;min-height:35px;display:flex;align-items:center;justify-content:center;margin:0 6px;padding:3px 12px;border-radius:18px;background:color-mix(in srgb,var(--tone) 18%,var(--surface));border:2px solid var(--tone);color:var(--tone);font-size:22px;font-weight:700;z-index:3}.coverage-callout{--tone:var(--primary);position:absolute;right:8px;top:76px;max-width:440px;padding:14px 18px;background:var(--tone);color:white;border-radius:4px;font-size:22px;font-weight:700;z-index:5}.visual-chips{display:flex;gap:12px;align-items:center;margin-top:12px}.visual-chips span{--tone:var(--muted);display:inline-flex;align-items:center;min-height:36px;padding:4px 18px;border:2px solid var(--tone);border-radius:18px;color:var(--tone);background:var(--surface);font-size:20px;font-weight:700}
.transformation-grid{display:grid;grid-template-columns:1fr 84px 1.24fr;gap:20px;height:505px;align-items:stretch}.evidence-input,.evidence-output{background:color-mix(in srgb,var(--primary) 6%,var(--surface));border:2px solid color-mix(in srgb,var(--primary) 18%,transparent);padding:24px;overflow:hidden}.evidence-input h2,.evidence-output h2{font-size:28px;margin:0 0 12px}.evidence-input small{display:block;font-size:20px;color:var(--muted);margin-bottom:16px}.tagged-copy{font-size:23px;line-height:1.58;background:var(--surface);border:1px solid color-mix(in srgb,var(--muted) 35%,transparent);padding:20px;height:365px;overflow:hidden}.tagged-segment{--tone:var(--muted)}.tagged-segment b{display:inline-block;border-radius:13px;padding:1px 8px;margin:0 5px;background:var(--tone);color:white;font-size:18px}.transformation-arrow{display:flex;align-items:center;justify-content:center;color:var(--secondary)}.transformation-arrow svg{width:70px;height:70px}.output-copy{min-height:132px;background:var(--primary);color:white;padding:18px;border-radius:4px;font-size:22px;line-height:1.4}.evidence-output h3{font-size:21px;margin:13px 0 8px}.compact-evidence-table .table-head,.compact-evidence-table .table-cell{font-size:19px;min-height:43px;padding:8px 10px}.compact-evidence-table{border-radius:6px}
.artifact-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:24px;height:500px}.artifact-grid.count-1{grid-template-columns:1fr}.artifact-grid.count-2{grid-template-columns:repeat(2,1fr)}.artifact-group{position:relative;background:var(--surface);border:2px solid color-mix(in srgb,var(--muted) 22%,transparent);padding:24px;overflow:hidden}.artifact-group h2{font-size:27px;margin:0 0 8px}.artifact-group>p{font-size:20px;line-height:1.35;color:var(--muted);margin:0}.artifact-stack{position:absolute;left:28px;right:28px;top:112px;bottom:24px}.artifact-preview{position:absolute;width:78%;height:82%;left:11%;top:8%;margin:0;border:2px solid color-mix(in srgb,var(--muted) 42%,transparent);border-radius:2px;background:white;overflow:visible}.artifact-preview img{display:block;width:100%;height:100%;background:white}.artifact-preview figcaption{position:absolute;left:0;bottom:-34px;font-size:19px;font-weight:700;color:var(--text)}.artifact-preview.preview-1{transform:translate(-8%,-3%) rotate(-2deg)}.artifact-preview.preview-2{transform:translate(7%,3%) rotate(2deg)}.artifact-preview.preview-3{transform:translate(0,8%)}.artifact-preview.preview-4{transform:translate(10%,-1%)}
.effort-label{position:absolute;right:0;top:-48px;font-size:25px;font-weight:800}.gantt-grid{display:grid;grid-template-columns:235px 390px repeat(var(--axis-count),minmax(0,1fr));grid-template-rows:60px repeat(var(--task-count),52px);position:relative;border:1px solid color-mix(in srgb,var(--muted) 28%,transparent);background:var(--surface)}.gantt-head{justify-content:flex-start;background:var(--primary);color:white}.gantt-axis{font-size:20px;padding:5px}.gantt-axis span{font-size:16px}.gantt-category,.gantt-task{display:flex;align-items:center;padding:5px 10px;border-right:1px solid color-mix(in srgb,var(--muted) 28%,transparent);border-bottom:1px solid color-mix(in srgb,var(--muted) 28%,transparent);font-size:19px;background:var(--surface);z-index:2}.gantt-task{flex-direction:column;align-items:flex-start;justify-content:center}.gantt-task strong{font-size:20px}.gantt-task span{font-size:16px;color:var(--muted)}.gantt-bar{--tone:var(--secondary);align-self:center;height:22px;margin:0 3px;background:var(--tone);border-radius:2px;z-index:4}.gantt-marker{--tone:var(--muted);position:relative;background:color-mix(in srgb,var(--tone) 12%,transparent);border-left:2px solid var(--tone);border-right:2px solid var(--tone);z-index:3}.gantt-marker span{position:absolute;top:5px;left:5px;writing-mode:vertical-rl;font-size:15px;color:var(--tone);font-weight:700}
.coverage-axis span,.coverage-group span,.gantt-axis span,.gantt-task span,.gantt-marker span,.tagged-segment b{font-size:19px}
`;
}

function normalizedItems(items) {
  return Array.isArray(items) ? items.filter((item) => typeof item === "string" && item.trim()).map((item) => item.trim()) : [];
}

function normalizeKind(kind) {
  return String(kind ?? "").replaceAll("-", "_").toLowerCase();
}

function toneClass(tone) {
  const normalized = String(tone ?? "neutral").replace(/^#/, "").toLowerCase();
  if (["positive", "success"].includes(normalized)) return "positive";
  if (["critical", "danger", "negative", "risk"].includes(normalized)) return "critical";
  if (["accent", "primary", "secondary", "warning", "neutral", "muted"].includes(normalized)) {
    return normalized === "muted" ? "neutral" : normalized;
  }
  return "neutral";
}

function alignment(value) {
  const normalized = String(value ?? "left").toLowerCase();
  return ["left", "center", "right"].includes(normalized) ? normalized : "left";
}

function toneStyle(tone) {
  const normalized = String(tone ?? "").replace(/^#/, "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(normalized) ? `--tone:#${normalized};` : "";
}

function relativeLuminance(hexColor) {
  const channels = String(hexColor).match(/.{2}/g)?.map((value) => {
    const channel = Number.parseInt(value, 16) / 255;
    return channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4;
  });
  return channels?.length === 3
    ? channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722
    : 1;
}

function color(value, fallback) {
  const normalized = String(value ?? fallback).replace(/^#/, "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(normalized) ? normalized : fallback;
}

function safeFont(value) {
  return String(value ?? "Noto Sans CJK JP").replace(/["'<>\\]/g, "").slice(0, 96);
}

function safeText(value) {
  return typeof value === "string" ? value : "";
}

function escapeHtml(value) {
  return safeText(value).replace(/[&<>"']/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;",
  })[character]);
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/[\r\n]/g, "&#10;");
}
