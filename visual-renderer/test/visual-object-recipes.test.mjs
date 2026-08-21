import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import test from "node:test";
import JSZip from "jszip";

const executeFile = promisify(execFile);
const rendererPath = fileURLToPath(new URL("../index.mjs", import.meta.url));
const rendererDirectory = dirname(rendererPath);

const recipeAssets = [
  {
    asset_id: "0123456789abcdef0123456789abcdea",
    fingerprint: "a".repeat(64),
    brief: {
      slideNumber: 1,
      visualPurpose: "direction",
      archetype: "arrow",
      style: "quietCorporate",
      emphasis: "subtle",
      orientation: "right",
      placementRole: "contentConnector",
      paletteRole: "accent",
      recipe: "directionalCue",
    },
  },
  {
    asset_id: "1123456789abcdef0123456789abcdea",
    fingerprint: "b".repeat(64),
    brief: {
      slideNumber: 2,
      visualPurpose: "growth",
      archetype: "arrow",
      style: "technical",
      emphasis: "standard",
      orientation: "up",
      placementRole: "chartAnnotation",
      paletteRole: "positive",
      label: "Acceleration",
      recipe: "growthPath",
    },
  },
  {
    asset_id: "2123456789abcdef0123456789abcdea",
    fingerprint: "c".repeat(64),
    brief: {
      slideNumber: 3,
      visualPurpose: "emphasis",
      archetype: "frame",
      style: "editorial",
      emphasis: "strong",
      orientation: "right",
      placementRole: "focusFrame",
      paletteRole: "accent",
      label: "Primary signal",
      recipe: "focusCorners",
    },
  },
  {
    asset_id: "3123456789abcdef0123456789abcdea",
    fingerprint: "d".repeat(64),
    brief: {
      slideNumber: 4,
      visualPurpose: "annotation",
      archetype: "callout",
      style: "roundedFriendly",
      emphasis: "subtle",
      orientation: "right",
      placementRole: "chartAnnotation",
      paletteRole: "primary",
      label: "Decision point",
      recipe: "annotationPin",
      anchorCategoryOrdinal: 2,
      anchorSeriesOrdinal: 1,
    },
  },
  {
    asset_id: "4123456789abcdef0123456789abcdea",
    fingerprint: "e".repeat(64),
    brief: {
      slideNumber: 5,
      visualPurpose: "annotation",
      archetype: "ribbon",
      style: "editorial",
      emphasis: "subtle",
      orientation: "right",
      placementRole: "sectionDivider",
      paletteRole: "secondary",
      label: "Implication",
      recipe: "sectionRule",
    },
  },
  {
    asset_id: "5123456789abcdef0123456789abcdea",
    fingerprint: "f".repeat(64),
    brief: {
      slideNumber: 6,
      visualPurpose: "cycle",
      archetype: "ring",
      style: "technical",
      emphasis: "subtle",
      orientation: "clockwise",
      placementRole: "backgroundMotif",
      paletteRole: "secondary",
      recipe: "cycleCue",
    },
  },
];

test("curated visual object recipes render as bounded editable native compositions", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-visual-object-recipes-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "visual-object-recipes.pptx");
  const chart = {
    kind: "Line",
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [{ name: "Adoption", values: [18, 31, 52, 78] }],
    showLegend: false,
  };

  try {
    await writeFile(
      specificationPath,
      JSON.stringify({
        title: "Curated visual object recipes",
        rendererContract: "visual-v5",
        theme: { preset: "minimal" },
        design: { style: "executive", density: "balanced", motif: "none" },
        visual_object_assets: recipeAssets,
        slides: [
          {
            kind: "Comparison",
            title: "A clear handoff",
            panels: [
              { title: "Observe", bullets: ["Collect evidence", "Name the constraint"] },
              { title: "Act", bullets: ["Select one response", "Set the owner"] },
            ],
            visualObjects: [{ assetId: recipeAssets[0].asset_id }],
          },
          {
            kind: "Chart",
            title: "Momentum compounds",
            chart,
            takeaway: "Adoption accelerates after Q2.",
            visualObjects: [{ assetId: recipeAssets[1].asset_id }],
          },
          {
            kind: "Metrics",
            title: "One signal deserves focus",
            variant: "spotlight",
            metrics: [
              { value: "78%", label: "Adoption", detail: "Primary signal", tone: "accent" },
              { value: "12", label: "Teams", detail: "Participating" },
              { value: "4.6", label: "Rating", detail: "Out of five" },
            ],
            visualObjects: [{ assetId: recipeAssets[2].asset_id }],
          },
          {
            kind: "Chart",
            title: "Annotate the decision, not the decoration",
            chart,
            visualObjects: [{ assetId: recipeAssets[3].asset_id }],
          },
          {
            kind: "Bullets",
            title: "The implication is separated quietly",
            bullets: ["Keep the evidence in the body", "Reserve the accent for the conclusion"],
            takeaway: "A restrained rule creates hierarchy without another card.",
            visualObjects: [{ assetId: recipeAssets[4].asset_id }],
          },
          {
            kind: "Process",
            title: "The loop is reinforced once",
            variant: "loop",
            steps: [
              { title: "Observe", description: "Capture evidence" },
              { title: "Learn", description: "Explain the signal" },
              { title: "Adapt", description: "Change the next action" },
            ],
            visualObjects: [{ assetId: recipeAssets[5].asset_id }],
          },
        ],
      }),
      "utf8",
    );
    await executeFile(process.execPath, [rendererPath, specificationPath, outputPath], {
      cwd: rendererDirectory,
      timeout: 30_000,
    });

    const archive = await JSZip.loadAsync(await readFile(outputPath));
    const slides = await Promise.all(
      Array.from({ length: 6 }, (_, index) =>
        archive.file(`ppt/slides/slide${index + 1}.xml`)?.async("string") ?? ""),
    );
    const allSlides = slides.join("\n");

    assert.doesNotMatch(allSlides, /<p:pic>/u);
    assert.match(slides[0], /<a:tailEnd type="triangle"\/>/u);
    assert.doesNotMatch(slides[0], /<a:t>Moves to<\/a:t>/u);
    assert.match(slides[1], /<a:tailEnd type="triangle"\/>/u);
    assert.ok((slides[1].match(/prst="ellipse"/gu) ?? []).length >= 2);
    assert.ok((slides[2].match(/prst="line"/gu) ?? []).length >= 8);
    assert.match(slides[2], /<a:t>Primary signal<\/a:t>/u);
    assert.match(slides[3], /prst="roundRect"/u);
    assert.match(slides[3], /<a:t>Decision point<\/a:t>/u);
    const annotationTarget = slides[3].match(
      /name="Visual Object annotationPin target"[\s\S]*?<a:off x="(\d+)" y="(\d+)"/u,
    );
    assert.ok(annotationTarget);
    assert.ok(Number(annotationTarget[1]) < 6.5 * 914400);
    const annotationRun = (slides[3].match(/<a:r>.*?<\/a:r>/gsu) ?? [])
      .find((run) => run.includes("<a:t>Decision point</a:t>")) ?? "";
    assert.match(annotationRun, /<a:srgbClr val="17213A"\/>/u);
    assert.ok((slides[4].match(/prst="line"/gu) ?? []).length >= 2);
    assert.match(slides[4], /<a:t>Implication<\/a:t>/u);
    assert.match(slides[5], /prst="circularArrow"/u);
    assert.match(slides[5], /prst="ellipse"/u);
    const annotationChart = await archive.file("ppt/charts/chart2.xml")?.async("string") ?? "";
    assert.match(annotationChart, /<c:min val="0"\/>/u);
    assert.match(annotationChart, /<c:max val="80"\/>/u);
    if (process.env.PPTX_MCP_VISUAL_EVIDENCE_DIR) {
      await mkdir(process.env.PPTX_MCP_VISUAL_EVIDENCE_DIR, { recursive: true });
      await copyFile(
        outputPath,
        join(process.env.PPTX_MCP_VISUAL_EVIDENCE_DIR, "visual-object-recipes.pptx"),
      );
    }
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
