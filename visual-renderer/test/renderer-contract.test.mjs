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

const renderSection = async (workingDirectory, contract, outputName) => {
  const specificationPath = join(workingDirectory, `${outputName}.json`);
  const outputPath = join(workingDirectory, `${outputName}.pptx`);
  await writeFile(
    specificationPath,
    JSON.stringify({
      title: "Contract test",
      rendererContract: contract,
      design: { style: "executive", density: "balanced", motif: "none" },
      slides: [
        {
          kind: "Section",
          title: "A stable section",
          subtitle: "Existing lineages retain their original composition",
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
  return archive.file("ppt/slides/slide1.xml")?.async("string");
};

test("visual-v4 lineages omit visual-v5 style decorations", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-renderer-contract-"));

  try {
    const legacyXml = (await renderSection(workingDirectory, "visual-v4", "legacy")) ?? "";
    const modernXml = (await renderSection(workingDirectory, "visual-v5", "modern")) ?? "";
    const shapeCount = (xml) => xml.match(/<p:sp>/gu)?.length ?? 0;

    assert.match(legacyXml, />A stable section</u);
    assert.match(modernXml, />A stable section</u);
    assert.equal(shapeCount(modernXml), shapeCount(legacyXml) + 1);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});

test("visual-v4 retains legacy foreground and semantic-tone colors", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-renderer-v4-colors-"));
  const specificationPath = join(workingDirectory, "legacy-colors.json");
  const outputPath = join(workingDirectory, "legacy-colors.pptx");
  const findTextRun = (xml, text) =>
    (xml.match(/<a:r>.*?<\/a:r>/gsu) ?? [])
      .find((run) => run.includes(`<a:t>${text}</a:t>`)) ?? "";

  try {
    await writeFile(
      specificationPath,
      JSON.stringify({
        title: "Legacy color contract",
        rendererContract: "visual-v4",
        theme: {
          preset: "minimal",
          primaryColor: "#FFB000",
          secondaryColor: "#FFF000",
          accentColor: "#FFF000",
          surfaceColor: "#111111",
          positiveColor: "#ABCDEF",
        },
        design: { style: "executive", density: "balanced", motif: "none" },
        slides: [
          {
            kind: "Comparison",
            title: "Comparison",
            panels: [
              { title: "First", highlight: "Legacy white highlight", bullets: ["A"] },
              { title: "Second", bullets: ["B"] },
            ],
          },
          {
            kind: "Process",
            title: "Process",
            steps: [
              { title: "Plan", label: "PHASE" },
              { title: "Build", label: "02" },
              { title: "Run", label: "03" },
            ],
          },
          {
            kind: "Closing",
            title: "Legacy primary foreground",
            takeaway: "Legacy accent pill text",
          },
          {
            kind: "Metrics",
            title: "Metrics",
            takeaway: "Legacy secondary takeaway",
            metrics: [
              { value: "42", label: "Positive", tone: "positive" },
              { value: "7", label: "Neutral", tone: "neutral" },
            ],
          },
          {
            kind: "Scorecard",
            title: "Scorecard",
            scorecard: {
              options: [{ title: "A" }, { title: "B" }],
              criteria: [
                {
                  criterion: "Legacy criterion",
                  cells: [
                    { rating: "Good", tone: "positive" },
                    { rating: "Fair", tone: "warning" },
                  ],
                },
                {
                  criterion: "Second",
                  cells: [
                    { rating: "Fair", tone: "warning" },
                    { rating: "Good", tone: "positive" },
                  ],
                },
              ],
            },
          },
          {
            kind: "Funnel",
            title: "Funnel",
            steps: [
              { title: "Alpha", label: "01" },
              { title: "Beta", label: "02" },
              { title: "Gamma", label: "03" },
            ],
          },
          {
            kind: "Chart",
            title: "Chart",
            chart: {
              kind: "bar",
              categories: ["First", "Second"],
              series: [{ name: "Legacy series", values: [10, 20] }],
              showLegend: true,
            },
          },
          {
            kind: "Dashboard",
            title: "Dashboard",
            metrics: [
              { value: "10", label: "First" },
              { value: "20", label: "Second" },
            ],
            chart: {
              kind: "line",
              categories: ["First", "Second"],
              series: [{ name: "Legacy dashboard series", values: [10, 20] }],
              showLegend: true,
            },
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
    const slideXml = async (number) =>
      await archive.file(`ppt/slides/slide${number}.xml`)?.async("string") ?? "";
    assert.match(findTextRun(await slideXml(1), "Legacy white highlight"), /val="FFFFFF"/u);
    assert.match(findTextRun(await slideXml(2), "PHASE"), /val="FFFFFF"/u);
    assert.match(findTextRun(await slideXml(3), "Legacy primary foreground"), /val="FFFFFF"/u);
    assert.match(findTextRun(await slideXml(3), "Legacy accent pill text"), /val="FFB000"/u);
    const metricsXml = await slideXml(4);
    assert.match(findTextRun(metricsXml, "42"), /val="159D73"/u);
    assert.match(findTextRun(metricsXml, "Legacy secondary takeaway"), /val="FFF000"/u);
    const scorecardXml = await slideXml(5);
    assert.match(findTextRun(scorecardXml, "Legacy criterion"), /val="202124"/u);
    assert.match(findTextRun(scorecardXml, "● Good"), /val="159D73"/u);
    assert.match(scorecardXml, /val="EEF2F7"/u);
    assert.match(scorecardXml, /val="F7F9FC"/u);
    assert.match(scorecardXml, /val="D9E0E8"/u);
    assert.match(findTextRun(await slideXml(6), "01  Alpha"), /val="202124"/u);
    const chartXml = await archive.file("ppt/charts/chart1.xml")?.async("string") ?? "";
    const dashboardChartXml = await archive.file("ppt/charts/chart2.xml")?.async("string") ?? "";
    for (const xml of [chartXml, dashboardChartXml]) {
      assert.match(xml, /val="D9DFE8"/u);
      assert.doesNotMatch(xml, /val="D9E0E8"/u);
      assert.doesNotMatch(xml, /<c:legend>.*?val="6B6F76".*?<\/c:legend>/su);
    }
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
