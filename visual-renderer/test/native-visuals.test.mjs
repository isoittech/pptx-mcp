import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import test from "node:test";
import JSZip from "jszip";

const executeFile = promisify(execFile);
const rendererPath = fileURLToPath(new URL("../index.mjs", import.meta.url));
const rendererDirectory = dirname(rendererPath);

test("semantic diagrams, real variants, and prepared objects remain editable native shapes", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-native-visuals-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "native-visuals.pptx");
  const steps = ["Discover", "Decide", "Build", "Adopt"].map((title, index) => ({
    title,
    description: `Outcome ${index + 1}`,
    label: `0${index + 1}`,
  }));
  const visualObjectAssets = [
    {
      asset_id: "0123456789abcdef0123456789abcdef",
      fingerprint: "a".repeat(64),
      brief: {
        slideNumber: 1,
        visualPurpose: "grouping",
        archetype: "bracket",
        style: "quietCorporate",
        emphasis: "subtle",
        orientation: "down",
        placementRole: "contentConnector",
        paletteRole: "muted",
        label: "Escalation boundary",
      },
    },
    {
      asset_id: "1123456789abcdef0123456789abcdef",
      fingerprint: "b".repeat(64),
      brief: {
        slideNumber: 5,
        visualPurpose: "annotation",
        archetype: "callout",
        style: "quietCorporate",
        emphasis: "standard",
        orientation: "down",
        placementRole: "contentConnector",
        paletteRole: "secondary",
        label: "Decision gate",
      },
    },
    {
      asset_id: "2123456789abcdef0123456789abcdef",
      fingerprint: "c".repeat(64),
      brief: {
        slideNumber: 9,
        visualPurpose: "emphasis",
        archetype: "frame",
        style: "quietCorporate",
        emphasis: "strong",
        orientation: "right",
        placementRole: "focusFrame",
        paletteRole: "accent",
        label: "Primary threshold",
      },
    },
  ];
  const specification = {
    title: "Native visual regression",
    rendererContract: "visual-v5",
    design: { style: "executive", density: "balanced", motif: "none" },
    visual_object_assets: visualObjectAssets,
    slides: [
      {
        kind: "NativeDiagram",
        title: "Decision tree",
        diagram: {
          kind: "tree",
          direction: "topToBottom",
          nodes: [
            { id: "root", label: "Need", emphasize: true },
            { id: "a", label: "Option A" },
            { id: "b", label: "Option B" },
            { id: "out", label: "Decision", tone: "positive" },
          ],
          edges: [
            { from: "root", to: "a", label: "low risk" },
            { from: "root", to: "b", label: "high value" },
            { from: "root", to: "out", label: "fallback" },
          ],
        },
        takeaway: "The decision rule remains separate from the three branches.",
        visualObjects: [{ assetId: visualObjectAssets[0].asset_id }],
      },
      { kind: "Process", title: "Learning loop", variant: "loop", steps },
      { kind: "Timeline", title: "Stepped timeline", variant: "stepped", steps },
      { kind: "Funnel", title: "Value pyramid", variant: "pyramid", steps },
      {
        kind: "Roadmap",
        title: "Staircase roadmap",
        variant: "stepped",
        steps,
        visualObjects: [{ assetId: visualObjectAssets[1].asset_id }],
      },
      {
        kind: "NativeDiagram",
        title: "Flywheel",
        diagram: {
          kind: "cycle",
          direction: "clockwise",
          nodes: steps.map((step, index) => ({ id: `c${index}`, label: step.title })),
        },
      },
      {
        kind: "NativeDiagram",
        title: "Concentric priorities",
        diagram: {
          kind: "concentric",
          direction: "leftToRight",
          nodes: [
            { id: "core", label: "Core", emphasize: true },
            { id: "near", label: "Adjacent" },
            { id: "future", label: "Future" },
          ],
        },
      },
      {
        kind: "Metrics",
        title: "Decision thresholds",
        variant: "spotlight",
        metrics: [
          { value: "300", label: "Participants", detail: "Six months", tone: "accent" },
          { value: "15%", label: "Return ratio", detail: "Target", tone: "positive" },
          { value: "4", label: "Themes", detail: "Minimum", tone: "warning" },
        ],
        visualObjects: [{ assetId: visualObjectAssets[2].asset_id }],
      },
      {
        kind: "NativeDiagram",
        title: "Partner network",
        diagram: {
          kind: "network",
          direction: "leftToRight",
          nodes: [
            { id: "hub", label: "Platform", emphasize: true },
            { id: "p1", label: "Partner 1" },
            { id: "p2", label: "Partner 2" },
            { id: "p3", label: "Partner 3" },
          ],
        },
      },
    ],
  };

  try {
    await writeFile(specificationPath, JSON.stringify(specification), "utf8");
    await executeFile(process.execPath, [rendererPath, specificationPath, outputPath], {
      cwd: rendererDirectory,
      timeout: 30_000,
    });
    const archive = await JSZip.loadAsync(await readFile(outputPath));
    const slideFiles = Object.keys(archive.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name));
    assert.equal(slideFiles.length, specification.slides.length);
    const allSlides = (await Promise.all(slideFiles.map((name) => archive.file(name)?.async("string")))).join("\n");
    assert.match(allSlides, /Decision tree/u);
    assert.match(allSlides, /Learning loop/u);
    assert.match(allSlides, /Value pyramid/u);
    assert.match(allSlides, /Staircase roadmap/u);
    assert.match(allSlides, /Partner network/u);
    assert.match(allSlides, /Escalation boundary/u);
    assert.match(allSlides, /Decision gate/u);
    assert.match(allSlides, /Primary threshold/u);
    const roadmapSlide = await archive.file("ppt/slides/slide5.xml")?.async("string");
    assert.match(roadmapSlide, /prst="wedgeRoundRectCallout"[\s\S]*?<a:t>Decision gate<\/a:t>/u);
    const decisionSlide = await archive.file("ppt/slides/slide1.xml")?.async("string");
    const shapeForText = (text) => decisionSlide
      ?.split("<p:sp>")
      .find((fragment) => fragment.includes(`<a:t>${text}</a:t>`));
    const shapeY = (text) => {
      const shape = shapeForText(text);
      const match = shape?.match(/<a:off x="\d+" y="(\d+)"\/>/u);
      assert.ok(match, `Expected positioned shape containing ${text}`);
      return Number(match[1]);
    };
    assert.ok(
      shapeY("Escalation boundary") < shapeY("Option A"),
      "The bracket label must stay above the lower diagram nodes instead of overlapping them.",
    );
    for (const text of ["Need", "Option A", "low risk", "high value", "fallback"]) {
      const shape = shapeForText(text);
      const sizes = [...(shape?.matchAll(/\bsz="(\d+)"/gu) ?? [])].map((match) => Number(match[1]));
      assert.ok(sizes.length > 0 && Math.min(...sizes) >= 1_400, `${text} must remain at least 14pt`);
    }
    assert.match(allSlides, /<a:tailEnd type="triangle"\/>/u);
    assert.doesNotMatch(allSlides, /<p:pic>/u);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
