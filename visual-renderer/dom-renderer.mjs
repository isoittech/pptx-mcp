import { writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { exportHtmlToPptx } from "dom-to-pptx/node";
import JSZip from "jszip";
import { parseFragment, serialize } from "parse5";
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
  "nativediagram",
  "native_diagram",
]));

export function canRenderDeckWithDom(spec) {
  return Array.isArray(spec?.slides)
    && spec.slides.length > 0
    && spec.slides.every(canRenderSlideWithDom);
}

export function canRenderSlideWithDom(slide) {
  if (slide?.authoredHtml && typeof slide.authoredHtml.html === "string" && typeof slide.authoredHtml.css === "string") {
    return true;
  }
  const kind = normalizeKind(slide?.kind);
  if (!domSupportedSlideKinds.has(kind)) return false;
  if (kind !== "nativediagram" && kind !== "native_diagram") return true;
  return ["tree", "flow"].includes(String(slide?.diagram?.kind ?? "").toLowerCase());
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
  const normalizedBuffer = await normalizeDomListTextInsets(buffer);
  await writeFile(outputPath, normalizedBuffer, { flag: "wx" });
}

// dom-to-pptx 2.1.1 supplies text margins as [top, right, bottom, left], while
// its bundled PptxGenJS 4.0.1 consumes arrays as [left, right, bottom, top].
// The mismatch turns a normal UL/OL left indent into a large top inset. Swap
// those two generated DrawingML attributes only for native list text boxes;
// non-list text boxes and intentionally separate wrapper padding stay intact.
export async function normalizeDomListTextInsets(pptxBuffer) {
  const archive = await JSZip.loadAsync(pptxBuffer);
  const slidePaths = Object.keys(archive.files)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/u.test(path));
  let changed = false;

  for (const slidePath of slidePaths) {
    const entry = archive.file(slidePath);
    if (!entry) continue;
    const xml = await entry.async("string");
    const normalizedXml = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/gu, (shape) => {
      if (!/<a:bu(?:Char|AutoNum)\b/u.test(shape)) return shape;
      return shape.replace(/<a:bodyPr\b[^>]*>/u, (bodyProperties) => {
        const leftInset = bodyProperties.match(/\blIns="(\d+)"/u)?.[1];
        const topInset = bodyProperties.match(/\btIns="(\d+)"/u)?.[1];
        if (leftInset === undefined || topInset === undefined) return bodyProperties;
        changed = true;
        return bodyProperties
          .replace(/\blIns="\d+"/u, `lIns="${topInset}"`)
          .replace(/\btIns="\d+"/u, `tIns="${leftInset}"`);
      });
    });
    if (normalizedXml !== xml) archive.file(slidePath, normalizedXml);
  }

  if (!changed) return pptxBuffer;
  return archive.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

export function buildDomDeckHtml(spec, imageAssets = {}) {
  const theme = {
    ...resolveTheme(spec.theme),
    visualObjectAssets: Array.isArray(spec.visual_object_assets) ? spec.visual_object_assets : [],
  };
  const templateChrome = spec.templateChrome === true;
  const defaultTemplateCoverOverlay = spec.defaultTemplateCoverOverlay === true;
  const defaultTemplateBodyStyle = spec.defaultTemplateBodyStyle === true;
  const slideNumberOffset = Number.isInteger(spec.slideNumberOffset) && spec.slideNumberOffset >= 0
    ? spec.slideNumberOffset
    : 0;
  const deckTotalSlides = Number.isInteger(spec.deckTotalSlides) && spec.deckTotalSlides >= spec.slides.length
    ? spec.deckTotalSlides
    : spec.slides.length;
  const usesModelAuthoredHtml = String(spec.rendererContract ?? "").toLowerCase() === "visual-v7-author-html";
  if (usesModelAuthoredHtml && spec.slides.some((slide) => !slide?.authoredHtml)) {
    throw new Error("Every visual-v7-author-html slide must contain authoredHtml.");
  }
  const slides = spec.slides.map((slide, index) => renderSlide(
    slide,
    slideNumberOffset + index,
    deckTotalSlides,
    theme,
    templateChrome,
    defaultTemplateCoverOverlay,
    defaultTemplateBodyStyle,
    imageAssets,
  )).join("\n");
  const css = usesModelAuthoredHtml
    ? `${modelAuthoredBaseCss(theme, templateChrome)}\n${spec.slides.map((slide, index) => validateAndScopeAuthoredCss(
      slide.authoredHtml.css,
      slideNumberOffset + index,
      {
        requireDefaultCoverContract: defaultTemplateCoverOverlay && slideNumberOffset + index === 0,
        requireDefaultBodyContract: defaultTemplateBodyStyle && slideNumberOffset + index > 0,
        roleClassNames: collectAuthoredRoleClassNames(slide.authoredHtml.html),
        roleTagNames: collectAuthoredRoleTagNames(slide.authoredHtml.html),
      },
    )).join("\n")}${modelAuthoredComplianceCss(spec.slides, slideNumberOffset, defaultTemplateCoverOverlay, defaultTemplateBodyStyle, templateChrome)}`
    : baseCss(theme);

  return `<!doctype html>
<html lang="${escapeAttribute(spec.language ?? "ja-JP")}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=${VIEWPORT_WIDTH},initial-scale=1">
<style>${css}</style>
</head>
<body>${slides}</body>
</html>`;
}

function renderSlide(
  slide,
  index,
  totalSlides,
  theme,
  templateChrome,
  defaultTemplateCoverOverlay,
  defaultTemplateBodyStyle,
  imageAssets,
) {
  if (slide.authoredHtml) {
    return renderModelAuthoredSlide(
      slide,
      index,
      theme,
      templateChrome,
      defaultTemplateCoverOverlay,
      defaultTemplateBodyStyle,
      imageAssets,
    );
  }
  const kind = normalizeKind(slide.kind);
  const content = renderers[kind](slide, index, theme, imageAssets);
  const notes = slide.speakerNotes
    ? `このスライドの狙い\n${safeText(slide.speakerNotes.purpose)}\n\nトークスクリプト\n${safeText(slide.speakerNotes.talkScript)}`
    : "";
  const usesDefaultCover = defaultTemplateCoverOverlay
    && index === 0
    && (kind === "title" || kind === "section");
  const usesDefaultBody = defaultTemplateBodyStyle && index > 0;
  const backgroundClass = `${templateChrome ? " template-chrome" : ""}${usesDefaultCover ? " default-template-cover" : ""}${usesDefaultBody ? " default-template-body" : ""}${usesDefaultBody && slide.subtitle ? " has-header-claim" : ""}`;
  const eyebrow = slide.eyebrow
    ? `<div class="eyebrow">${escapeHtml(slide.eyebrow)}</div>`
    : "";
  const subtitle = slide.subtitle
    ? usesDefaultBody
      ? `<ul class="header-claim"><li>${escapeHtml(slide.subtitle)}</li></ul>`
      : `<p>${escapeHtml(slide.subtitle)}</p>`
    : "";
  const footer = templateChrome || kind === "title" || kind === "section" || kind === "closing"
    ? ""
    : `<div class="footer"><span>${escapeHtml(safeText(slide.attribution))}</span><span>${index + 1} / ${totalSlides}</span></div>`;

  return `<section class="slide${backgroundClass}">
    ${notes ? `<template data-pptx-notes>${escapeHtml(notes)}</template>` : ""}
    <div class="accent-rail"></div>
    <header>${eyebrow}<h1>${escapeHtml(slide.title)}</h1>${subtitle}</header>
    <main class="kind-${escapeAttribute(kind)}">${content}</main>
    ${renderPreparedVisualObjects(slide, index, theme)}
    ${footer}
  </section>`;
}

const allowedAuthoredHtmlTags = new Set([
  "div", "section", "article", "header", "footer", "main", "aside",
  "h1", "h2", "h3", "h4", "h5", "h6", "p", "blockquote", "span", "strong", "b", "em", "i", "small",
  "ul", "ol", "li", "dl", "dt", "dd", "table", "colgroup", "col", "thead", "tbody", "tfoot", "tr", "th", "td",
  "figure", "figcaption", "img", "br", "hr",
]);
const allowedAuthoredHtmlAttributes = new Set([
  "class", "id", "style", "title", "alt", "role", "aria-label", "aria-hidden",
  "colspan", "rowspan", "width", "height", "data-pptx-icon", "data-pptx-asset", "data-pptx-role",
]);
const allowedAuthoredRoles = new Set(["cover-title", "cover-subtitle", "body-title", "body-claim", "source-meta"]);
const authoredRoleFontSizes = Object.freeze({
  "cover-title": 54,
  "cover-subtitle": 27,
  "body-title": 50,
  "body-claim": 27,
});

function renderModelAuthoredSlide(
  slide,
  index,
  _theme,
  _templateChrome,
  defaultTemplateCoverOverlay,
  defaultTemplateBodyStyle,
  imageAssets,
) {
  const notes = slide.speakerNotes
    ? `このスライドの狙い\n${safeText(slide.speakerNotes.purpose)}\n\nトークスクリプト\n${safeText(slide.speakerNotes.talkScript)}`
    : "";
  const authored = slide.authoredHtml ?? {};
  const fragment = sanitizeAuthoredHtmlFragment(
    authored.html,
    Array.isArray(authored.assetIds) ? authored.assetIds : [],
    imageAssets,
    {
      requireDefaultCoverContract: defaultTemplateCoverOverlay && index === 0,
      expectedCoverTitle: slide.title,
      expectedCoverSubtitle: slide.subtitle,
      requireDefaultBodyContract: defaultTemplateBodyStyle && index > 0,
      expectedBodyTitle: slide.title,
      expectedBodyClaim: slide.subtitle,
    },
  );
  return `<section class="slide" data-author-slide="${index + 1}">
    ${notes ? `<template data-pptx-notes>${escapeHtml(notes)}</template>` : ""}
    ${fragment}
  </section>`;
}

export function sanitizeAuthoredHtmlFragment(
  html,
  declaredAssetIds = [],
  imageAssets = {},
  {
    requireDefaultCoverContract = false,
    expectedCoverTitle = "",
    expectedCoverSubtitle = "",
    requireDefaultBodyContract = false,
    expectedBodyTitle = "",
    expectedBodyClaim = "",
  } = {},
) {
  if (typeof html !== "string" || html.length < 1 || html.length > 24_000) {
    throw new Error("Model-authored HTML must contain between 1 and 24000 characters.");
  }
  assertNoUnsafeResourceSyntax(html, "HTML");
  const fragment = parseFragment(html);
  const declared = new Set(declaredAssetIds);
  const referenced = new Set();
  const roleNodes = new Map();
  let elementCount = 0;
  let iconCount = 0;

  const visit = (node) => {
    if (node.nodeName === "#comment") {
      throw new Error("Model-authored HTML comments are not accepted.");
    }
    if (!node.tagName) {
      for (const child of node.childNodes ?? []) visit(child);
      return;
    }

    elementCount += 1;
    if (elementCount > 500) {
      throw new Error("Model-authored HTML may contain at most 500 elements per slide.");
    }
    const tag = String(node.tagName).toLowerCase();
    if (!allowedAuthoredHtmlTags.has(tag)) {
      throw new Error(`Model-authored HTML tag is not allowed: ${tag}`);
    }

    const attributes = new Map();
    for (const attribute of node.attrs ?? []) {
      const name = String(attribute.name).toLowerCase();
      if (!allowedAuthoredHtmlAttributes.has(name)) {
        throw new Error(`Model-authored HTML attribute is not allowed: ${name}`);
      }
      if (attributes.has(name)) {
        throw new Error(`Model-authored HTML contains a duplicate attribute: ${name}`);
      }
      attributes.set(name, attribute.value);
    }
    const authoredRole = String(attributes.get("data-pptx-role") ?? "").trim().toLowerCase();
    if (authoredRole) {
      if (!allowedAuthoredRoles.has(authoredRole)) {
        throw new Error(`Model-authored HTML role is not allowed: ${authoredRole}`);
      }
      const roleAttribute = node.attrs.find((attribute) => attribute.name === "data-pptx-role");
      if (roleAttribute) roleAttribute.value = authoredRole;
      const entries = roleNodes.get(authoredRole) ?? [];
      entries.push(node);
      roleNodes.set(authoredRole, entries);
    }
    if (attributes.has("style")) {
      validateSafeCssDeclarations(attributes.get("style"), "inline style");
      validateAuthoredFontSizes(
        attributes.get("style"),
        authoredRole || null,
        "inline style",
      );
    }

    const iconName = attributes.get("data-pptx-icon");
    if (iconName) {
      iconCount += 1;
      if (tag === "img") throw new Error("data-pptx-icon must be placed on a container element, not img.");
      const iconFragment = parseFragment(renderApprovedReactIcon(iconName, { color: "currentColor" }));
      node.childNodes = iconFragment.childNodes ?? [];
      for (const child of node.childNodes) child.parentNode = node;
      node.attrs = (node.attrs ?? []).filter((attribute) => attribute.name !== "data-pptx-icon");
      const classAttribute = node.attrs.find((attribute) => attribute.name === "class");
      if (classAttribute) classAttribute.value = `${classAttribute.value} pptx-icon`.trim();
      else node.attrs.push({ name: "class", value: "pptx-icon" });
    }

    const assetId = attributes.get("data-pptx-asset");
    if (tag === "img") {
      if (!assetId || !declared.has(assetId) || !imageAssets[assetId]?.data) {
        throw new Error("Every model-authored img must reference one declared, server-verified data-pptx-asset ID.");
      }
      if (!attributes.get("alt")?.trim()) {
        throw new Error("Every model-authored img must have non-empty alt text.");
      }
      referenced.add(assetId);
      node.attrs = (node.attrs ?? []).filter((attribute) => attribute.name !== "data-pptx-asset");
      node.attrs.push({ name: "src", value: imageAssets[assetId].data });
    } else if (assetId) {
      throw new Error("data-pptx-asset is valid only on img elements.");
    }

    if (iconName) return;
    for (const child of node.childNodes ?? []) visit(child);
  };
  visit(fragment);

  if (declared.size !== referenced.size || [...declared].some((assetId) => !referenced.has(assetId))) {
    throw new Error("authoredHtml.assetIds must exactly match the data-pptx-asset references in HTML.");
  }
  if (requireDefaultCoverContract) {
    const titles = roleNodes.get("cover-title") ?? [];
    const subtitles = roleNodes.get("cover-subtitle") ?? [];
    if (titles.length !== 1 || !["h1", "h2"].includes(String(titles[0]?.tagName ?? "").toLowerCase())) {
      throw new Error("A default-template cover must contain exactly one h1 or h2 with data-pptx-role=\"cover-title\".");
    }
    if (subtitles.length !== 1 || String(subtitles[0]?.tagName ?? "").toLowerCase() !== "p") {
      throw new Error("A default-template cover must contain exactly one p with data-pptx-role=\"cover-subtitle\".");
    }
    const visibleTitle = normalizeVisibleText(textContent(titles[0]));
    const visibleSubtitle = normalizeVisibleText(textContent(subtitles[0]));
    const expectedTitle = normalizeVisibleText(expectedCoverTitle);
    const expectedSubtitle = normalizeVisibleText(expectedCoverSubtitle);
    if (!expectedTitle || !expectedSubtitle
        || visibleTitle !== expectedTitle
        || visibleSubtitle !== expectedSubtitle) {
      throw new Error("The default-template cover title and subtitle must be non-empty and exactly match slide.title and slide.subtitle.");
    }
    if (normalizeVisibleText(textContent(fragment)) !== normalizeVisibleText(`${expectedTitle}${expectedSubtitle}`)
        || declared.size > 0
        || iconCount > 0) {
      throw new Error("A default-template cover may contain only its title and subtitle, without extra visible copy, images, or icons.");
    }
  }
  if (requireDefaultBodyContract) {
    const titles = roleNodes.get("body-title") ?? [];
    const claims = roleNodes.get("body-claim") ?? [];
    if (titles.length !== 1 || !["h1", "h2"].includes(String(titles[0]?.tagName ?? "").toLowerCase())) {
      throw new Error("A default-template body slide must contain exactly one h1 or h2 with data-pptx-role=\"body-title\".");
    }
    if (claims.length !== 1 || String(claims[0]?.tagName ?? "").toLowerCase() !== "ul") {
      throw new Error("A default-template body slide must contain exactly one ul with data-pptx-role=\"body-claim\".");
    }
    const claimItems = (claims[0].childNodes ?? []).filter((child) => child.tagName === "li");
    if (claimItems.length !== 1 || !textContent(claimItems[0]).trim()) {
      throw new Error("The default-template body claim must contain exactly one non-empty native list item.");
    }
    const visibleTitle = normalizeVisibleText(textContent(titles[0]));
    const visibleClaim = normalizeVisibleText(textContent(claimItems[0]));
    if (!visibleTitle) {
      throw new Error("The default-template body title must not be empty.");
    }
    if (visibleTitle !== normalizeVisibleText(expectedBodyTitle)
        || visibleClaim !== normalizeVisibleText(expectedBodyClaim)) {
      throw new Error("The default-template body title and claim must exactly match slide.title and slide.subtitle.");
    }
  }
  return serialize(fragment);
}

export function validateAndScopeAuthoredCss(css, slideIndex, options = {}) {
  if (typeof css !== "string" || css.length < 1 || css.length > 16_000) {
    throw new Error("Model-authored CSS must contain between 1 and 16000 characters.");
  }
  assertNoUnsafeResourceSyntax(css, "CSS");
  if (css.includes("/*") || css.includes("*/") || css.includes("@") || /::?(?:before|after)\b/iu.test(css)) {
    throw new Error("Model-authored CSS may not use comments, @rules, or pseudo-element content.");
  }
  validateSafeCssDeclarations(css, "CSS");

  let cursor = 0;
  let scoped = "";
  const declaredRoleFontSizes = new Set();
  const rulePattern = /([^{}]+)\{([^{}]*)\}/gu;
  for (const match of css.matchAll(rulePattern)) {
    if (match.index !== cursor && css.slice(cursor, match.index).trim()) {
      throw new Error("Model-authored CSS must contain only flat style rules.");
    }
    const selectors = match[1].split(",").map((selector) => selector.trim());
    if (selectors.length === 0 || selectors.some((selector) => !/^\.slide(?:$|[\s.#:[>+~])/u.test(selector))) {
      throw new Error("Every model-authored CSS selector must be scoped below .slide.");
    }
    const hasFontSize = /(?:^|;)\s*font-size\s*:/iu.test(match[2]);
    const authoredRole = hasFontSize
      ? classifyAuthoredFontRole(selectors, options.roleClassNames, options.roleTagNames)
      : null;
    validateAuthoredFontSizes(match[2], authoredRole, "CSS");
    if (authoredRole) declaredRoleFontSizes.add(authoredRole);
    const scope = `.slide[data-author-slide="${slideIndex + 1}"]`;
    scoped += `${selectors.map((selector) => selector.replace(/^\.slide/u, scope)).join(",")}{${match[2]}}\n`;
    cursor = match.index + match[0].length;
  }
  if (cursor === 0 || css.slice(cursor).trim()) {
    throw new Error("Model-authored CSS contains an incomplete or nested rule.");
  }
  if (options.requireDefaultBodyContract
      && (!declaredRoleFontSizes.has("body-title") || !declaredRoleFontSizes.has("body-claim"))) {
    throw new Error("Default-template body CSS must explicitly size body-title at 50px and body-claim plus its li at 27px so the authored layout matches the final PowerPoint.");
  }
  if (options.requireDefaultCoverContract
      && (!declaredRoleFontSizes.has("cover-title") || !declaredRoleFontSizes.has("cover-subtitle"))) {
    throw new Error("Default-template cover CSS must explicitly size cover-title at 54px and cover-subtitle at 27px so the authored layout matches the final PowerPoint.");
  }
  return scoped;
}

function classifyAuthoredFontRole(selectors, roleClassNames = {}, roleTagNames = {}) {
  const roleSelector = (role) => new RegExp(
    `\\[data-pptx-role=(?:"${role}"|'${role}'|${role})\\]\\s*$`,
    "u",
  );
  const classSelector = (role, selector) => (roleClassNames[role] ?? []).some((className) => {
    const escapedClassName = escapeRegExp(className);
    if (role === "body-claim") {
      return new RegExp(`\\.${escapedClassName}(?:\\s+li)?\\s*$`, "u").test(selector);
    }
    return new RegExp(`\\.${escapedClassName}\\s*$`, "u").test(selector);
  });
  const tagSelector = (role, selector) => (roleTagNames[role] ?? []).some((tagName) =>
    new RegExp(`(?:^|[\\s>+~])${escapeRegExp(tagName)}\\s*$`, "u").test(selector));
  const roles = selectors.map((selector) => {
    if (roleSelector("source-meta").test(selector) || classSelector("source-meta", selector) || tagSelector("source-meta", selector)) return "source-meta";
    if (roleSelector("cover-title").test(selector) || classSelector("cover-title", selector) || tagSelector("cover-title", selector)) return "cover-title";
    if (roleSelector("cover-subtitle").test(selector) || classSelector("cover-subtitle", selector) || tagSelector("cover-subtitle", selector)) return "cover-subtitle";
    if (roleSelector("body-title").test(selector) || classSelector("body-title", selector) || tagSelector("body-title", selector)) return "body-title";
    if (roleSelector("body-claim").test(selector)
        || /\[data-pptx-role=(?:"body-claim"|'body-claim'|body-claim)\]\s+li\s*$/u.test(selector)
        || classSelector("body-claim", selector)
        || tagSelector("body-claim", selector)) {
      return "body-claim";
    }
    return null;
  });
  const matchedRoles = [...new Set(roles.filter(Boolean))];
  if (matchedRoles.length === 0) return null;
  if (matchedRoles.length !== 1 || roles.some((role) => role === null)) {
    throw new Error("A CSS font-size rule for a protected PowerPoint role may not be grouped with unrelated selectors or another protected role.");
  }
  return matchedRoles[0];
}

function collectAuthoredRoleClassNames(html) {
  const fragment = parseFragment(String(html ?? ""));
  const classUsages = new Map();
  const visit = (node) => {
    if (node?.tagName) {
      const role = String((node.attrs ?? []).find((attribute) => attribute.name === "data-pptx-role")?.value ?? "")
        .trim()
        .toLowerCase();
      const classes = String((node.attrs ?? []).find((attribute) => attribute.name === "class")?.value ?? "")
        .split(/\s+/u)
        .filter((className) => /^[A-Za-z_][A-Za-z0-9_-]*$/u.test(className));
      for (const className of classes) {
        const roles = classUsages.get(className) ?? new Set();
        roles.add(role || null);
        classUsages.set(className, roles);
      }
    }
    for (const child of node?.childNodes ?? []) visit(child);
  };
  visit(fragment);

  const result = {};
  for (const [className, roles] of classUsages.entries()) {
    if (roles.size !== 1) continue;
    const [role] = roles;
    if (!allowedAuthoredRoles.has(role)) continue;
    result[role] ??= [];
    result[role].push(className);
  }
  return result;
}

function collectAuthoredRoleTagNames(html) {
  const fragment = parseFragment(String(html ?? ""));
  const usages = new Map();
  const visit = (node) => {
    if (node?.tagName) {
      const tagName = String(node.tagName).toLowerCase();
      const role = String((node.attrs ?? []).find((attribute) => attribute.name === "data-pptx-role")?.value ?? "")
        .trim()
        .toLowerCase();
      const usage = usages.get(tagName) ?? { count: 0, roles: new Set() };
      usage.count += 1;
      usage.roles.add(role || null);
      usages.set(tagName, usage);
    }
    for (const child of node?.childNodes ?? []) visit(child);
  };
  visit(fragment);

  const result = {};
  for (const [tagName, usage] of usages.entries()) {
    if (usage.count !== 1 || usage.roles.size !== 1) continue;
    const [role] = usage.roles;
    if (!allowedAuthoredRoles.has(role)) continue;
    result[role] ??= [];
    result[role].push(tagName);
  }
  return result;
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

function validateAuthoredFontSizes(declarations, authoredRole, label) {
  for (const match of String(declarations).matchAll(/(?:^|;)\s*font-size\s*:\s*([^;}]*)/giu)) {
    const value = match[1].trim();
    const parsed = /^(\d+(?:\.\d+)?)px(?:\s*!important)?$/iu.exec(value);
    if (!parsed) {
      throw new Error(`Model-authored ${label} font-size must use a literal px value.`);
    }
    const pixels = Number(parsed[1]);
    const required = authoredRoleFontSizes[authoredRole];
    if (required !== undefined && pixels !== required) {
      throw new Error(`Model-authored ${label} font-size for ${authoredRole} must be exactly ${required}px so the authored layout matches the final PowerPoint size.`);
    }
    const sourceMetadataOnly = authoredRole === "source-meta";
    const minimum = sourceMetadataOnly ? 20 : 24;
    if (!Number.isFinite(pixels) || pixels < minimum) {
      throw new Error(`Model-authored ${label} font-size must be at least ${minimum}px${sourceMetadataOnly ? " for source metadata" : ""}.`);
    }
  }
}

function textContent(node) {
  if (!node) return "";
  if (typeof node.value === "string") return node.value;
  return (node.childNodes ?? []).map(textContent).join("");
}

function normalizeVisibleText(value) {
  return String(value ?? "").replace(/\s+/gu, " ").trim();
}

function validateSafeCssDeclarations(value, label) {
  if (/url\s*\(|expression\s*\(|(?:^|[;{])\s*(?:animation|transition|behavior|-moz-binding|content)\s*:|position\s*:\s*fixed|(?:box|text)-shadow\s*:/iu.test(value)) {
    throw new Error(`Model-authored ${label} contains a disallowed CSS feature.`);
  }
}

function assertNoUnsafeResourceSyntax(value, label) {
  if (/(?:https?|file|javascript|data):|(?:^|[\s'"(])\/\/|\\\\/iu.test(value)) {
    throw new Error(`Model-authored ${label} may not reference remote resources, data URIs, or local paths.`);
  }
  if (/<\/?(?:html|head|body|style|script|link|meta|base|iframe|object|embed|form|input|button|textarea|select|video|audio|canvas|svg)\b/iu.test(value)
    || /\son[a-z][a-z0-9_-]*\s*=/iu.test(value)
    || /\s(?:src|href|xlink:href)\s*=/iu.test(value)) {
    throw new Error(`Model-authored ${label} contains executable markup or an unmanaged resource attribute.`);
  }
}

function modelAuthoredBaseCss(theme, templateChrome) {
  return `
*{box-sizing:border-box;box-shadow:none!important;text-shadow:none!important}
html,body{margin:0;padding:0;background:transparent}
body{font-family:"${theme.bodyFont}",sans-serif;color:#${theme.text}}
.slide{--primary:#${theme.primary};--secondary:#${theme.secondary};--accent:#${theme.accent};--bg:#${theme.background};--surface:#${theme.surface};--text:#${theme.text};--muted:#${theme.muted};--footer:#${theme.footer};--positive:#${theme.positive};--warning:#${theme.warning};--critical:#${theme.critical};position:relative;width:${VIEWPORT_WIDTH}px;height:${VIEWPORT_HEIGHT}px;overflow:hidden;background:${templateChrome ? "transparent" : "var(--bg)"};font-family:"${theme.bodyFont}",sans-serif;font-size:24px;color:var(--text)}
.slide .pptx-icon svg{display:block;width:100%;height:100%}
.slide img{display:block;max-width:100%;max-height:100%}
`;
}

function modelAuthoredComplianceCss(slides, slideNumberOffset, defaultTemplateCoverOverlay, defaultTemplateBodyStyle, templateChrome) {
  const rules = [];
  for (let index = 0; index < slides.length; index += 1) {
    const number = slideNumberOffset + index + 1;
    const scope = `.slide[data-author-slide="${number}"]`;
    if (templateChrome) rules.push(`${scope}{background:transparent!important}`);
    rules.push(`${scope} [data-pptx-role="source-meta"]{font-size:20px!important}`);
    if (defaultTemplateCoverOverlay && number === 1) {
      rules.push(`${scope} div,${scope} section,${scope} article,${scope} header,${scope} main,${scope} aside{background:transparent!important;border:0!important}`);
      rules.push(`${scope} [data-pptx-role="cover-title"]{font-size:54px!important;line-height:1.12!important;margin:0!important;color:#fff!important;font-weight:800!important;background:transparent!important;border:0!important}`);
      rules.push(`${scope} [data-pptx-role="cover-subtitle"]{font-size:27px!important;line-height:1.4!important;margin:14px 0 0!important;color:#f5f7fb!important;font-weight:600!important;background:transparent!important;border:0!important}`);
    }
    if (defaultTemplateBodyStyle && number > 1) {
      rules.push(`${scope} [data-pptx-role="body-title"]{font-size:50px!important;line-height:1.1!important;margin:0!important;color:var(--secondary)!important;font-weight:800!important}`);
      rules.push(`${scope} [data-pptx-role="body-claim"]{font-size:27px!important;margin:8px 0 0!important;padding-left:0!important;color:var(--secondary)!important;font-weight:700!important}`);
      rules.push(`${scope} [data-pptx-role="body-claim"] li{font-size:27px!important;color:var(--secondary)!important;font-weight:700!important}`);
    }
  }
  return rules.length > 0 ? `\n${rules.join("\n")}` : "";
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
  nativediagram: renderNativeDiagram,
  native_diagram: renderNativeDiagram,
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
  const columns = split
    ? [items.slice(0, Math.ceil(items.length / 2)), items.slice(Math.ceil(items.length / 2))]
    : [items];
  const body = `<div class="bullet-list${split ? " split" : ""}">${columns.map((column) => `<ul class="bullet-column">${column.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`).join("")}</div>`;
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

const DIAGRAM_WIDTH = 1408;
const DIAGRAM_HEIGHT = 520;
const DIAGRAM_MINIMUM_GAP = 32;

export function layoutDirectedDiagram(diagram) {
  const nodes = Array.isArray(diagram?.nodes) ? diagram.nodes : [];
  const edges = Array.isArray(diagram?.edges) ? diagram.edges : [];
  const nodeById = new Map(nodes.map((node) => [node.id, node]));
  const incoming = new Map(nodes.map((node) => [node.id, 0]));
  for (const edge of edges) {
    if (nodeById.has(edge.from) && nodeById.has(edge.to)) {
      incoming.set(edge.to, (incoming.get(edge.to) ?? 0) + 1);
    }
  }
  const levels = new Map(nodes.filter((node) => incoming.get(node.id) === 0).map((node) => [node.id, 0]));
  for (let pass = 0; pass < nodes.length; pass += 1) {
    for (const edge of edges) {
      if (!levels.has(edge.from) || !nodeById.has(edge.to)) continue;
      levels.set(edge.to, Math.max(levels.get(edge.to) ?? 0, levels.get(edge.from) + 1));
    }
  }
  nodes.forEach((node, index) => {
    if (!levels.has(node.id)) levels.set(node.id, index);
  });

  const maximumLevel = Math.max(...levels.values(), 0);
  const vertical = String(diagram?.direction ?? "leftToRight").toLowerCase() === "toptobottom";
  const nodeWidth = vertical ? 330 : 340;
  const nodeHeight = vertical ? 94 : 104;
  const margin = 24;
  const positions = new Map();

  for (let level = 0; level <= maximumLevel; level += 1) {
    const levelNodes = nodes.filter((node) => levels.get(node.id) === level);
    if (vertical) {
      const available = DIAGRAM_WIDTH - margin * 2;
      const width = Math.min(nodeWidth, (available - DIAGRAM_MINIMUM_GAP * Math.max(0, levelNodes.length - 1)) / Math.max(1, levelNodes.length));
      const groupWidth = width * levelNodes.length + DIAGRAM_MINIMUM_GAP * Math.max(0, levelNodes.length - 1);
      const y = maximumLevel === 0
        ? (DIAGRAM_HEIGHT - nodeHeight) / 2
        : margin + level * (DIAGRAM_HEIGHT - margin * 2 - nodeHeight) / maximumLevel;
      levelNodes.forEach((node, itemIndex) => {
        positions.set(node.id, {
          x: (DIAGRAM_WIDTH - groupWidth) / 2 + itemIndex * (width + DIAGRAM_MINIMUM_GAP),
          y,
          width,
          height: nodeHeight,
        });
      });
      continue;
    }

    const x = maximumLevel === 0
      ? (DIAGRAM_WIDTH - nodeWidth) / 2
      : margin + level * (DIAGRAM_WIDTH - margin * 2 - nodeWidth) / maximumLevel;
    const gap = (DIAGRAM_HEIGHT - nodeHeight * levelNodes.length) / (levelNodes.length + 1);
    if (levelNodes.length > 1 && gap < DIAGRAM_MINIMUM_GAP) {
      throw new Error("The directed diagram cannot preserve the required spacing between nodes.");
    }
    levelNodes.forEach((node, itemIndex) => {
      positions.set(node.id, {
        x,
        y: gap + itemIndex * (nodeHeight + gap),
        width: nodeWidth,
        height: nodeHeight,
      });
    });
  }

  return { positions, edges, vertical };
}

function renderNativeDiagram(slide, index) {
  const diagram = slide.diagram ?? {};
  const kind = String(diagram.kind ?? "").toLowerCase();
  if (!["tree", "flow"].includes(kind)) {
    throw new Error("The DOM renderer currently supports tree and flow NativeDiagram pages only.");
  }
  const nodes = Array.isArray(diagram.nodes) ? diagram.nodes : [];
  const layout = layoutDirectedDiagram(diagram);
  const markerId = `diagram-arrow-${index + 1}`;
  const connectorMarkup = layout.edges.map((edge) => {
    const from = layout.positions.get(edge.from);
    const to = layout.positions.get(edge.to);
    if (!from || !to) return "";
    const x1 = layout.vertical ? from.x + from.width / 2 : from.x + from.width;
    const y1 = layout.vertical ? from.y + from.height : from.y + from.height / 2;
    const x2 = layout.vertical ? to.x + to.width / 2 : to.x;
    const y2 = layout.vertical ? to.y : to.y + to.height / 2;
    return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" marker-end="url(#${markerId})"/>`;
  }).join("");
  const edgeLabels = layout.edges.map((edge) => {
    if (!edge.label) return "";
    const from = layout.positions.get(edge.from);
    const to = layout.positions.get(edge.to);
    if (!from || !to) return "";
    const x1 = layout.vertical ? from.x + from.width / 2 : from.x + from.width;
    const y1 = layout.vertical ? from.y + from.height : from.y + from.height / 2;
    const x2 = layout.vertical ? to.x + to.width / 2 : to.x;
    const y2 = layout.vertical ? to.y : to.y + to.height / 2;
    const width = 188;
    const height = 42;
    const left = Math.max(0, Math.min(DIAGRAM_WIDTH - width, (x1 + x2 - width) / 2));
    const top = Math.max(0, Math.min(DIAGRAM_HEIGHT - height, (y1 + y2 - height) / 2));
    return `<div class="diagram-edge-label" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px">${escapeHtml(edge.label)}</div>`;
  }).join("");
  const nodeMarkup = nodes.map((node) => {
    const position = layout.positions.get(node.id);
    if (!position) return "";
    const emphasis = node.emphasize ? " emphasized" : "";
    return `<article class="diagram-node tone-${toneClass(node.tone)}${emphasis}" style="${toneStyle(node.tone)}left:${position.x}px;top:${position.y}px;width:${position.width}px;height:${position.height}px"><strong>${escapeHtml(node.label)}</strong>${node.description ? `<p>${escapeHtml(node.description)}</p>` : ""}</article>`;
  }).join("");
  return `<div class="diagram-stage kind-${escapeAttribute(kind)}"><svg class="diagram-connectors" viewBox="0 0 ${DIAGRAM_WIDTH} ${DIAGRAM_HEIGHT}" aria-hidden="true"><defs><marker id="${markerId}" markerWidth="10" markerHeight="10" refX="9" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z"/></marker></defs>${connectorMarkup}</svg>${nodeMarkup}${edgeLabels}</div>${takeaway(slide.takeaway)}`;
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
header{height:154px;position:relative;z-index:2}header h1{font-family:"${theme.headingFont}",sans-serif;font-size:50px;line-height:1.08;letter-spacing:-1.2px;margin:8px 0 0;color:var(--primary);max-width:1320px}header p{font-size:24px;line-height:1.35;color:var(--muted);margin:12px 0 0}.eyebrow{font-size:24px;line-height:1;text-transform:uppercase;letter-spacing:2.4px;font-weight:700;color:var(--secondary)}main{height:588px;position:relative;z-index:2}.footer{position:absolute;left:96px;right:96px;bottom:24px;display:flex;justify-content:space-between;font-size:14px;color:var(--footer)}.footer>span:first-child{flex:1}.footer>span:last-child{display:block;width:88px;text-align:right;white-space:nowrap}
.slide.default-template-body header h1{font-size:50px;color:var(--secondary);font-weight:800}.header-claim{margin:10px 0 0;padding-left:34px;color:var(--secondary);font-size:27px;line-height:1.3;font-weight:700}.header-claim li{margin:0;padding-left:4px}.header-claim li::marker{font-size:19px}.slide.has-header-claim header{height:176px}.slide.has-header-claim main{height:566px}
.slide.default-template-cover{padding:246px 112px 138px}.default-template-cover .accent-rail{display:none}.default-template-cover header{height:auto;max-width:1230px}.default-template-cover header h1{margin:0;color:#fff;font-size:54px;line-height:1.12;max-width:1230px}.default-template-cover header p,.default-template-cover .eyebrow{color:#f5f7fb}.default-template-cover main{height:auto;margin-top:34px}.default-template-cover .section-stage{height:auto;display:block}.default-template-cover .section-number{display:none}.default-template-cover .section-stage p,.default-template-cover .hero-copy p{color:#f5f7fb;font-size:27px;line-height:1.45;max-width:1040px;margin:0}.default-template-cover .hero-grid{display:block}.default-template-cover .hero-mark{display:none}
.hero-grid{display:grid;grid-template-columns:2fr 1fr;gap:72px;height:100%;align-items:center}.hero-copy p{font-size:29px;line-height:1.45;max-width:860px}.hero-mark{height:360px;border-radius:40px;background:var(--surface);border:2px solid color-mix(in srgb,var(--secondary) 20%,transparent);display:flex;align-items:center;justify-content:center}.icon-shell{width:170px;height:170px}.agenda-list{display:grid;grid-template-columns:1fr 1fr;gap:22px 32px}.agenda-row{display:flex;gap:28px;align-items:center;padding:22px 28px;background:var(--surface);border-radius:22px;border:1px solid color-mix(in srgb,var(--muted) 22%,transparent)}.agenda-row span{font-size:24px;font-weight:800;color:var(--secondary)}.agenda-row p{font-size:25px;margin:0;font-weight:600}.section-stage{height:100%;display:flex;align-items:center;gap:56px}.section-number{font-size:190px;line-height:1;font-weight:800;color:var(--accent);letter-spacing:-10px}.section-stage p{font-size:32px;line-height:1.45;max-width:780px}.lead{font-size:24px;line-height:1.45;margin:0 0 28px}.bullet-list{display:grid;grid-template-columns:1fr;gap:28px;height:100%}.bullet-list.split{grid-template-columns:1fr 1fr}.bullet-column{display:block;height:100%;background:var(--surface);padding:22px 30px 22px 58px;border-radius:18px;margin:0;list-style-type:disc;list-style-position:outside}.bullet-column li{font-size:24px;line-height:1.3;margin:0 0 13px;padding-left:5px}.bullet-column li:last-child{margin-bottom:0}.bullet-column li::marker{color:var(--accent);font-size:18px}.metric-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:24px}.metric-grid.spotlight{grid-template-columns:1.55fr repeat(3,1fr)}.metric-card{--tone:var(--secondary);min-height:288px;padding:34px 30px;border-radius:26px;background:var(--surface);border-top:9px solid var(--tone)}.metric-card.featured{background:color-mix(in srgb,var(--tone) 9%,var(--surface))}.metric-value{font-size:60px;line-height:1;font-weight:800;color:var(--tone);letter-spacing:-2px}.metric-label{font-size:24px;font-weight:700;margin-top:24px}.metric-card p{font-size:24px;line-height:1.4;color:var(--muted)}
.tone-accent{--tone:var(--accent)}.tone-primary{--tone:var(--primary)}.tone-secondary{--tone:var(--secondary)}.tone-positive{--tone:var(--positive)}.tone-warning{--tone:var(--warning)}.tone-critical{--tone:var(--critical)}.tone-neutral{--tone:var(--muted)}.panel-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:28px}.panel-grid.count-3{grid-template-columns:repeat(3,1fr)}.panel{position:relative;min-height:370px;padding:36px;background:var(--surface);border-radius:28px;border:1px solid color-mix(in srgb,var(--muted) 20%,transparent)}.panel-index{font-size:18px;font-weight:800;color:var(--accent)}.panel h2{font-size:29px;margin:14px 0}.panel strong{display:block;font-size:22px;color:var(--secondary);margin-bottom:16px}.compact-list p{font-size:19px;line-height:1.4;margin:12px 0;padding-left:18px;border-left:4px solid var(--accent)}.step-flow{display:flex;align-items:center;justify-content:center;height:410px}.step{width:230px;min-height:250px;padding:28px 24px;background:var(--surface);border-radius:26px;text-align:center;border:1px solid color-mix(in srgb,var(--muted) 18%,transparent)}.step span{font-size:17px;font-weight:800;color:var(--accent)}.step h2{font-size:24px;margin:22px 0 10px}.step p{font-size:17px;line-height:1.4;color:var(--muted)}.connector{font-size:42px;color:var(--secondary);padding:0 16px}.loop-label{text-align:center;color:var(--secondary);font-size:18px;font-weight:800;letter-spacing:2px}.timeline{position:relative;display:grid;grid-template-columns:repeat(5,1fr);gap:12px;padding-top:110px}.timeline-line{position:absolute;left:7%;right:7%;top:148px;height:6px;background:var(--secondary);border-radius:3px}.timeline-step{text-align:center;position:relative}.timeline-step>span{font-size:17px;font-weight:800;color:var(--secondary)}.timeline-step>div{width:30px;height:30px;border:8px solid var(--accent);background:var(--surface);border-radius:50%;margin:21px auto 28px}.timeline-step h2{font-size:21px;margin:0 10px 12px}.timeline-step p{font-size:16px;line-height:1.35;color:var(--muted);margin:0 10px}.statement{height:100%;display:grid;grid-template-columns:180px 1fr;align-items:center;gap:52px}.statement-icon{width:150px;height:150px}.statement p{font-family:"${theme.headingFont}",sans-serif;font-size:45px;line-height:1.3;font-weight:700;margin:0;color:var(--primary)}.statement strong{display:block;font-size:22px;color:var(--secondary);margin-top:28px}.card-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:24px}.card-grid.count-4{grid-template-columns:repeat(4,1fr)}.visual-card{min-height:355px;padding:28px;border-radius:26px;background:var(--surface);border-bottom:8px solid var(--tone)}.card-icon{width:70px;height:70px;color:var(--tone)}.visual-card>strong{display:block;font-size:38px;margin-top:20px;color:var(--tone)}.visual-card h2{font-size:24px;margin:20px 0 10px}.visual-card p{font-size:17px;line-height:1.45;color:var(--muted)}.matrix-wrap{position:relative;padding:0 0 44px 55px}.matrix-grid{display:grid;grid-template-columns:1fr 1fr;grid-template-rows:1fr 1fr;height:430px;border-left:5px solid var(--primary);border-bottom:5px solid var(--primary)}.matrix-grid article{padding:22px 28px;background:var(--surface);border:1px solid color-mix(in srgb,var(--muted) 18%,transparent)}.matrix-grid h2{font-size:22px;margin:0 0 10px}.matrix-grid strong{color:var(--secondary)}.matrix-grid p{font-size:16px;margin:7px 0}.matrix-y{position:absolute;left:0;top:180px;writing-mode:vertical-rl;transform:rotate(180deg);font-size:18px;font-weight:700}.matrix-x{text-align:center;margin-top:16px;font-size:18px;font-weight:700}.funnel{display:flex;flex-direction:column;align-items:center;gap:8px}.funnel>div{min-height:72px;padding:13px 28px;background:var(--secondary);color:white;text-align:center;clip-path:polygon(5% 0,95% 0,100% 100%,0 100%)}.funnel span{font-size:14px;font-weight:800;margin-right:18px}.funnel strong{font-size:22px}.funnel p{display:inline;font-size:15px;margin-left:18px}.roadmap{display:grid;grid-template-columns:repeat(4,1fr);gap:22px;padding-top:58px}.roadmap article{min-height:330px;background:var(--surface);border-radius:26px;padding:30px;border-top:9px solid var(--secondary)}.roadmap span{font-size:15px;font-weight:800;color:var(--accent)}.roadmap h2{font-size:24px;margin:24px 0 14px}.roadmap p{font-size:17px;line-height:1.45;color:var(--muted)}.quote{height:100%;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center}.quote>div{font-size:120px;line-height:.5;color:var(--accent)}.quote blockquote{font-family:"${theme.headingFont}",sans-serif;font-size:42px;line-height:1.35;max-width:1120px;margin:42px 0 24px}.quote p{font-size:20px;color:var(--muted)}.closing{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center}.closing-icon{width:110px;height:110px}.closing>p{font-size:28px;max-width:900px}.closing>div:last-child{display:flex;gap:18px;margin-top:24px}.closing span{background:var(--surface);padding:13px 22px;border-radius:999px;font-size:17px}.brief-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:24px}.brief-grid.count-2{grid-template-columns:repeat(2,1fr)}.brief-grid article{min-height:380px;background:var(--surface);padding:28px;border-radius:24px;border-top:8px solid var(--tone)}.brief-grid article>span{font-size:16px;font-weight:800;color:var(--tone)}.brief-grid h2{font-size:24px;margin:14px 0}.brief-grid strong{display:block;color:var(--tone);font-size:19px;margin-bottom:12px}.brief-grid p,.brief-bullet{font-size:16px;line-height:1.45;color:var(--muted)}.brief-bullet{margin-top:9px;padding-left:13px;border-left:3px solid var(--tone)}.data-grid{display:grid;gap:2px;background:color-mix(in srgb,var(--muted) 20%,transparent);border:2px solid color-mix(in srgb,var(--muted) 20%,transparent);border-radius:16px;overflow:hidden}.table-head,.row-head,.table-cell{background:var(--surface);padding:16px 18px;min-height:54px;font-size:16px;display:flex;flex-direction:column;justify-content:center}.table-head{background:var(--primary);color:white;font-weight:700}.table-head span,.table-cell span{font-size:13px;line-height:1.35;margin-top:5px}.row-head{font-weight:700}.table-cell strong{color:var(--tone)}.table-cell.emphasized{font-weight:800}.align-center{text-align:center;align-items:center}.align-right{text-align:right;align-items:flex-end}.media-grid{display:grid;grid-template-columns:1fr 1.15fr;gap:34px;height:430px}.media-copy{background:var(--surface);border-radius:25px;padding:32px}.media-copy>p{font-size:22px;line-height:1.5;margin:0 0 24px}.media-copy>div{font-size:18px;line-height:1.4;margin:14px 0;padding-left:16px;border-left:4px solid var(--accent)}.media-copy small{display:block;margin-top:24px;color:var(--muted)}figure{margin:0;border-radius:28px;overflow:hidden;background:var(--surface)}figure img{width:100%;height:100%;display:block}.takeaway{position:absolute;left:0;right:0;bottom:4px;min-height:70px;display:flex;align-items:center;gap:22px;padding:13px 24px;background:color-mix(in srgb,var(--accent) 12%,var(--surface));border-left:8px solid var(--accent);border-radius:10px}.takeaway span{font-size:13px;font-weight:800;letter-spacing:1px;color:var(--secondary)}.takeaway p{flex:1;min-width:0;font-size:18px;line-height:1.35;margin:0}.visual-object{position:absolute;z-index:3;color:var(--accent)}.visual-object span{font-size:15px;font-weight:700}.recipe-sectionrule{left:96px;right:96px;bottom:112px;border-top:5px solid var(--secondary);padding-top:8px}.recipe-focuscorners{right:84px;top:220px;width:320px;height:280px;border:7px solid var(--accent);background:transparent}.recipe-directionalcue{left:50%;top:55%;font-size:40px}.recipe-directionalcue:after{content:"→"}.recipe-growthpath{right:110px;top:250px;width:160px;height:160px;border-top:8px solid var(--positive);transform:rotate(-28deg)}.recipe-cyclecue{right:95px;bottom:100px;width:110px;height:110px;border:8px solid var(--secondary);border-radius:50%}.recipe-annotationpin{right:120px;top:300px;background:var(--surface);border:3px solid var(--primary);border-radius:18px;padding:12px 18px}
.coverage-stage,.gantt-stage{position:relative;height:510px}.coverage-grid{display:grid;grid-template-columns:180px 230px repeat(var(--axis-count),minmax(0,1fr));grid-template-rows:62px repeat(var(--row-count),52px);position:relative;border:1px solid color-mix(in srgb,var(--muted) 28%,transparent);background:var(--surface)}.coverage-corner,.coverage-axis,.gantt-head,.gantt-axis{display:flex;align-items:center;justify-content:center;padding:8px 10px;border-right:1px solid color-mix(in srgb,var(--muted) 28%,transparent);border-bottom:1px solid color-mix(in srgb,var(--muted) 28%,transparent);font-size:23px;font-weight:700}.coverage-corner{justify-content:flex-start;background:var(--primary);color:white}.coverage-axis,.gantt-axis{flex-direction:column;background:color-mix(in srgb,var(--primary) 7%,var(--surface))}.coverage-axis span,.gantt-axis span{font-size:18px;color:var(--muted)}.coverage-group{--tone:var(--accent);display:flex;flex-direction:column;align-items:center;justify-content:center;padding:8px;background:var(--tone);color:white;border-bottom:1px solid var(--surface);font-size:23px;text-align:center;z-index:2}.coverage-group span{font-size:18px}.coverage-row-label{display:flex;align-items:center;padding:6px 14px;border-right:1px solid color-mix(in srgb,var(--muted) 28%,transparent);border-bottom:1px solid color-mix(in srgb,var(--muted) 28%,transparent);font-size:22px;font-weight:700;z-index:2;background:var(--surface)}.coverage-cell,.gantt-cell{border-right:1px solid color-mix(in srgb,var(--muted) 22%,transparent);border-bottom:1px solid color-mix(in srgb,var(--muted) 22%,transparent);z-index:1}.coverage-bar{--tone:var(--secondary);align-self:center;min-height:35px;display:flex;align-items:center;justify-content:center;margin:0 6px;padding:3px 12px;border-radius:18px;background:color-mix(in srgb,var(--tone) 18%,var(--surface));border:2px solid var(--tone);color:var(--tone);font-size:22px;font-weight:700;z-index:3}.coverage-callout{--tone:var(--primary);position:absolute;right:8px;top:76px;max-width:440px;padding:14px 18px;background:var(--tone);color:white;border-radius:4px;font-size:22px;font-weight:700;z-index:5}.visual-chips{display:flex;gap:12px;align-items:center;margin-top:12px}.visual-chips span{--tone:var(--muted);display:inline-flex;align-items:center;min-height:36px;padding:4px 18px;border:2px solid var(--tone);border-radius:18px;color:var(--tone);background:var(--surface);font-size:20px;font-weight:700}
.transformation-grid{display:grid;grid-template-columns:1fr 84px 1.24fr;gap:20px;height:505px;align-items:stretch}.evidence-input,.evidence-output{background:color-mix(in srgb,var(--primary) 6%,var(--surface));border:2px solid color-mix(in srgb,var(--primary) 18%,transparent);padding:24px;overflow:hidden}.evidence-input h2,.evidence-output h2{font-size:28px;margin:0 0 12px}.evidence-input small{display:block;font-size:20px;color:var(--muted);margin-bottom:16px}.tagged-copy{font-size:23px;line-height:1.58;background:var(--surface);border:1px solid color-mix(in srgb,var(--muted) 35%,transparent);padding:20px;height:365px;overflow:hidden}.tagged-segment{--tone:var(--muted)}.tagged-segment b{display:inline-block;border-radius:13px;padding:1px 8px;margin:0 5px;background:var(--tone);color:white;font-size:18px}.transformation-arrow{display:flex;align-items:center;justify-content:center;color:var(--secondary)}.transformation-arrow svg{width:70px;height:70px}.output-copy{min-height:132px;background:var(--primary);color:white;padding:18px;border-radius:4px;font-size:22px;line-height:1.4}.evidence-output h3{font-size:21px;margin:13px 0 8px}.compact-evidence-table .table-head,.compact-evidence-table .table-cell{font-size:19px;min-height:43px;padding:8px 10px}.compact-evidence-table{border-radius:6px}
.artifact-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:24px;height:500px}.artifact-grid.count-1{grid-template-columns:1fr}.artifact-grid.count-2{grid-template-columns:repeat(2,1fr)}.artifact-group{position:relative;background:var(--surface);border:2px solid color-mix(in srgb,var(--muted) 22%,transparent);padding:24px;overflow:hidden}.artifact-group h2{font-size:27px;margin:0 0 8px}.artifact-group>p{font-size:20px;line-height:1.35;color:var(--muted);margin:0}.artifact-stack{position:absolute;left:28px;right:28px;top:112px;bottom:24px}.artifact-preview{position:absolute;width:78%;height:82%;left:11%;top:8%;margin:0;border:2px solid color-mix(in srgb,var(--muted) 42%,transparent);border-radius:2px;background:white;overflow:visible}.artifact-preview img{display:block;width:100%;height:100%;background:white}.artifact-preview figcaption{position:absolute;left:0;bottom:-34px;font-size:19px;font-weight:700;color:var(--text)}.artifact-preview.preview-1{transform:translate(-8%,-3%) rotate(-2deg)}.artifact-preview.preview-2{transform:translate(7%,3%) rotate(2deg)}.artifact-preview.preview-3{transform:translate(0,8%)}.artifact-preview.preview-4{transform:translate(10%,-1%)}
.effort-label{position:absolute;right:0;top:-48px;font-size:25px;font-weight:800}.gantt-grid{display:grid;grid-template-columns:235px 390px repeat(var(--axis-count),minmax(0,1fr));grid-template-rows:60px repeat(var(--task-count),52px);position:relative;border:1px solid color-mix(in srgb,var(--muted) 28%,transparent);background:var(--surface)}.gantt-head{justify-content:flex-start;background:var(--primary);color:white}.gantt-axis{font-size:20px;padding:5px}.gantt-axis span{font-size:16px}.gantt-category,.gantt-task{display:flex;align-items:center;padding:5px 10px;border-right:1px solid color-mix(in srgb,var(--muted) 28%,transparent);border-bottom:1px solid color-mix(in srgb,var(--muted) 28%,transparent);font-size:19px;background:var(--surface);z-index:2}.gantt-task{flex-direction:column;align-items:flex-start;justify-content:center}.gantt-task strong{font-size:20px}.gantt-task span{font-size:16px;color:var(--muted)}.gantt-bar{--tone:var(--secondary);align-self:center;height:22px;margin:0 3px;background:var(--tone);border-radius:2px;z-index:4}.gantt-marker{--tone:var(--muted);position:relative;background:color-mix(in srgb,var(--tone) 12%,transparent);border-left:2px solid var(--tone);border-right:2px solid var(--tone);z-index:3}.gantt-marker span{position:absolute;top:5px;left:5px;writing-mode:vertical-rl;font-size:15px;color:var(--tone);font-weight:700}
.diagram-stage{position:relative;width:${DIAGRAM_WIDTH}px;height:${DIAGRAM_HEIGHT}px;margin:0 auto}.diagram-connectors{position:absolute;inset:0;width:100%;height:100%;overflow:visible}.diagram-connectors line{stroke:var(--secondary);stroke-width:2.4;fill:none}.diagram-connectors marker path{fill:var(--secondary)}.diagram-node{--tone:var(--secondary);position:absolute;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:12px 16px;background:color-mix(in srgb,var(--tone) 10%,var(--surface));border:2px solid color-mix(in srgb,var(--tone) 72%,var(--surface));border-radius:8px;text-align:center;color:var(--text);z-index:2;overflow:hidden}.diagram-node.emphasized{background:var(--tone);border-color:var(--tone);color:white}.diagram-node strong{font-size:24px;line-height:1.2}.diagram-node p{font-size:24px;line-height:1.2;margin:5px 0 0}.diagram-edge-label{position:absolute;display:flex;align-items:center;justify-content:center;padding:2px 8px;background:var(--bg);color:var(--secondary);font-size:24px;line-height:1.15;font-weight:700;text-align:center;z-index:3}
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
