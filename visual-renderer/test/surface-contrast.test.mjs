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

const findTextRun = (xml, text) => {
  const runs = xml.match(/<a:r>.*?<\/a:r>/gsu) ?? [];
  return runs.find((run) => run.includes(`<a:t>${text}</a:t>`)) ?? "";
};

const findRunColor = (xml, text) =>
  findTextRun(xml, text).match(/<a:srgbClr val="([0-9A-F]{6})"/u)?.[1] ?? "";

const relativeLuminance = (color) => [0, 2, 4]
  .map((offset) => Number.parseInt(color.slice(offset, offset + 2), 16) / 255)
  .map((value) => value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4)
  .reduce((sum, value, index) => sum + value * [0.2126, 0.7152, 0.0722][index], 0);

const contrastRatio = (first, second) => {
  const luminances = [relativeLuminance(first), relativeLuminance(second)].sort((a, b) => b - a);
  return (luminances[0] + 0.05) / (luminances[1] + 0.05);
};

test("Cards use readable on-surface text when only a dark surface color is supplied", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-surface-contrast-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "surface-cards.pptx");

  try {
    await writeFile(
      specificationPath,
      JSON.stringify({
        title: "Surface contrast",
        rendererContract: "visual-v5",
        theme: { preset: "minimal", surfaceColor: "#111111" },
        design: { style: "executive", density: "balanced", motif: "none" },
        slides: [{
          kind: "Cards",
          title: "Readable cards",
          cards: [
            { title: "Card Alpha", description: "First surface description", icon: "insight" },
            { title: "Card Beta", description: "Second surface description", icon: "target" },
            { title: "Card Gamma", description: "Third surface description", icon: "growth" },
          ],
        }],
      }),
      "utf8",
    );
    await executeFile(process.execPath, [rendererPath, specificationPath, outputPath], {
      cwd: rendererDirectory,
      timeout: 30_000,
    });

    const archive = await JSZip.loadAsync(await readFile(outputPath));
    const slideXml = await archive.file("ppt/slides/slide1.xml")?.async("string") ?? "";
    assert.match(slideXml, /val="111111"/u);
    assert.match(findTextRun(slideXml, "Card Alpha"), /val="FFFFFF"/u);
    assert.ok(contrastRatio(findRunColor(slideXml, "First surface description"), "111111") >= 3);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});

test("all surface-backed business layouts use on-surface text roles", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-surface-layouts-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "surface-layouts.pptx");
  const chart = {
    kind: "Line",
    categories: ["North", "South"],
    series: [{ name: "Demand", values: [10, 14] }],
    showLegend: true,
  };

  try {
    await writeFile(
      specificationPath,
      JSON.stringify({
        title: "Surface layout contrast",
        rendererContract: "visual-v5",
        theme: { preset: "minimal", surfaceColor: "#111111" },
        design: { style: "executive", density: "balanced", motif: "none" },
        slides: [
          { kind: "Agenda", title: "Agenda", bullets: ["Agenda surface item", "Second agenda item"] },
          { kind: "Bullets", title: "Bullets", bullets: ["Bullet surface item", "Second bullet item"] },
          {
            kind: "Metrics",
            title: "Metrics",
            metrics: [
              { value: "42", label: "Metric surface label", detail: "Metric muted detail" },
              { value: "7", label: "Second metric" },
            ],
          },
          {
            kind: "Comparison",
            title: "Comparison",
            panels: [
              { title: "Comparison surface title", bullets: ["Comparison surface bullet"] },
              { title: "Alternative", bullets: ["Alternative detail"] },
            ],
          },
          {
            kind: "StructuredBrief",
            title: "Brief",
            sections: [
              { heading: "Brief surface heading", body: "Brief surface body" },
              { heading: "Second heading", body: "Second body" },
            ],
          },
          {
            kind: "Process",
            title: "Process",
            steps: [
              { title: "Process surface title", description: "Process muted detail" },
              { title: "Build", description: "Build detail" },
              { title: "Run", description: "Run detail" },
            ],
          },
          {
            kind: "Roadmap",
            title: "Roadmap",
            steps: [
              { title: "Roadmap surface title", description: "Roadmap muted detail" },
              { title: "Next", description: "Next detail" },
              { title: "Later", description: "Later detail" },
            ],
          },
          { kind: "Chart", title: "Chart", chart },
          {
            kind: "Dashboard",
            title: "Dashboard",
            metrics: [
              { value: "9", label: "Dashboard surface label", detail: "Dashboard muted detail" },
              { value: "8", label: "Second dashboard metric" },
            ],
            chart,
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
    const expectations = [
      [1, "Agenda surface item", "FFFFFF"],
      [2, "Bullet surface item", "FFFFFF"],
      [3, "Metric surface label", "FFFFFF"],
      [3, "Metric muted detail", null],
      [4, "Comparison surface title", "FFFFFF"],
      [4, "Comparison surface bullet", "FFFFFF"],
      [5, "Brief surface heading", "FFFFFF"],
      [5, "Brief surface body", "FFFFFF"],
      [6, "Process surface title", "FFFFFF"],
      [6, "Process muted detail", null],
      [7, "Roadmap surface title", "FFFFFF"],
      [7, "Roadmap muted detail", null],
      [9, "Dashboard surface label", "FFFFFF"],
      [9, "Dashboard muted detail", null],
    ];
    for (const [slideNumber, text, color] of expectations) {
      const xml = await archive.file(`ppt/slides/slide${slideNumber}.xml`)?.async("string") ?? "";
      const runColor = findRunColor(xml, text);
      if (color) {
        assert.equal(runColor, color, `${text} must use ${color}`);
      } else {
        assert.ok(contrastRatio(runColor, "111111") >= 3, `${text} must remain readable on the surface`);
      }
    }
    const chartXml = await archive.file("ppt/charts/chart1.xml")?.async("string") ?? "";
    const dashboardChartXml = await archive.file("ppt/charts/chart2.xml")?.async("string") ?? "";
    assert.match(chartXml, /val="6B6F76"/u);
    assert.match(dashboardChartXml, /val="6B6F76"/u);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});

test("bright primary secondary and accent fills use dynamic dark foreground roles", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-bright-role-contrast-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "bright-roles.pptx");

  try {
    await writeFile(
      specificationPath,
      JSON.stringify({
        title: "Bright role contrast",
        rendererContract: "visual-v5",
        theme: {
          preset: "minimal",
          primaryColor: "#FFF000",
          secondaryColor: "#FFF000",
          accentColor: "#FFF000",
        },
        design: { style: "executive", density: "balanced", motif: "none" },
        slides: [
          {
            kind: "Bullets",
            title: "Bright secondary",
            eyebrow: "Secondary foreground",
            bullets: ["First", "Second"],
          },
          {
            kind: "Comparison",
            title: "Bright comparison",
            panels: [
              { title: "First", highlight: "Bright comparison highlight", bullets: ["A"] },
              { title: "Second", bullets: ["B"] },
            ],
          },
          {
            kind: "Closing",
            title: "Bright primary title",
            takeaway: "Bright accent takeaway",
          },
          {
            kind: "Section",
            title: "Bright background section",
            eyebrow: "Bright section eyebrow",
          },
          {
            kind: "Metrics",
            title: "Bright background metric",
            takeaway: "Bright background takeaway",
            metrics: [
              { value: "10", label: "First metric" },
              { value: "20", label: "Second metric" },
            ],
          },
          {
            kind: "Statement",
            title: "Bright statement",
            eyebrow: "Bright statement eyebrow",
            body: "The accent role remains readable on the primary background.",
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
    const firstXml = await archive.file("ppt/slides/slide1.xml")?.async("string") ?? "";
    const secondXml = await archive.file("ppt/slides/slide2.xml")?.async("string") ?? "";
    const thirdXml = await archive.file("ppt/slides/slide3.xml")?.async("string") ?? "";
    const fourthXml = await archive.file("ppt/slides/slide4.xml")?.async("string") ?? "";
    const fifthXml = await archive.file("ppt/slides/slide5.xml")?.async("string") ?? "";
    const sixthXml = await archive.file("ppt/slides/slide6.xml")?.async("string") ?? "";
    assert.match(findTextRun(firstXml, "SECONDARY FOREGROUND"), /val="17213A"/u);
    assert.match(findTextRun(secondXml, "Bright comparison highlight"), /val="17213A"/u);
    assert.match(findTextRun(thirdXml, "Bright primary title"), /val="17213A"/u);
    assert.match(findTextRun(thirdXml, "Bright accent takeaway"), /val="17213A"/u);
    const backgroundTextColors = [
      findRunColor(fourthXml, "BRIGHT SECTION EYEBROW"),
      findRunColor(fifthXml, "Bright background takeaway"),
    ];
    for (const color of backgroundTextColors) {
      assert.notEqual(color, "FFF000");
      assert.ok(contrastRatio(color, "F7F7F5") >= 4.5);
    }
    const primaryTextColor = findRunColor(sixthXml, "BRIGHT STATEMENT EYEBROW");
    assert.notEqual(primaryTextColor, "FFF000");
    assert.ok(contrastRatio(primaryTextColor, "FFF000") >= 4.5);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
