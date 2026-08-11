import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import test from "node:test";
import JSZip from "jszip";

const executeFile = promisify(execFile);
const rendererPath = fileURLToPath(new URL("../index.mjs", import.meta.url));
const rendererDirectory = dirname(rendererPath);

test("DataTable renders as an editable native PowerPoint table", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-data-table-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "data-table.pptx");
  const specification = {
    title: "Operations report",
    language: "en-US",
    rendererContract: "visual-v5",
    theme: {
      fontFace: "Legacy Test Font",
      headingFontFace: "Approved Heading Test",
      bodyFontFace: "Approved Body Test",
      backgroundColor: "#0B0F14",
      surfaceColor: "#101820",
      textColor: "#FFFFFF",
      mutedTextColor: "#B8C1CC",
      positiveColor: "#00FF88",
      warningColor: "#FFAA00",
      criticalColor: "#FF3366",
      dataSeriesColors: ["#123456", "#ABCDEF"],
    },
    design: { style: "technical", density: "airy", motif: "none" },
    slides: [
      {
        kind: "DataTable",
        title: "Capacity by region",
        density: "detailed",
        recipeId: "report-table-detailed",
        takeaway: "Increase west-region capacity first",
        dataTable: {
          firstColumnIsHeader: true,
          columns: [
            { header: "Region", align: "left", widthWeight: 1.4 },
            { header: "Utilization", align: "right", widthWeight: 1 },
            { header: "Action", align: "left", widthWeight: 2.2 },
          ],
          rows: [
            {
              cells: [
                { text: "East", tone: "neutral" },
                { text: "82%", tone: "positive", emphasize: true },
                { text: "Maintain", tone: "neutral" },
              ],
            },
            {
              cells: [
                { text: "West", tone: "neutral" },
                { text: "96%", tone: "warning", emphasize: true },
                { text: "Accelerate expansion", tone: "neutral" },
              ],
            },
            {
              cells: [
                { text: "Central", tone: "neutral" },
                { text: "99%", tone: "critical", emphasize: true },
                { text: "Escalate now", tone: "neutral" },
              ],
            },
          ],
        },
      },
      {
        kind: "Chart",
        title: "Capacity trend",
        chart: {
          kind: "Line",
          categories: ["Q1", "Q2"],
          series: [
            { name: "East", values: [72, 82] },
            { name: "West", values: [91, 96] },
          ],
          showLegend: true,
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

    const output = await stat(outputPath);
    const packageBytes = await readFile(outputPath);
    const signature = packageBytes.subarray(0, 2).toString("ascii");
    const archive = await JSZip.loadAsync(packageBytes);
    const slideXml = await archive.file("ppt/slides/slide1.xml")?.async("string");
    const chartXml = await archive.file("ppt/charts/chart1.xml")?.async("string");
    assert.ok(output.size > 10_000);
    assert.equal(signature, "PK");
    assert.match(slideXml ?? "", /<a:tbl>/u);
    assert.match(slideXml ?? "", />Utilization</u);
    assert.match(slideXml ?? "", />96%</u);
    assert.match(slideXml ?? "", /typeface="Approved Heading Test"/u);
    assert.match(slideXml ?? "", /typeface="Approved Body Test"/u);
    assert.doesNotMatch(slideXml ?? "", /typeface="Legacy Test Font"/u);
    assert.match(slideXml ?? "", /val="101820"/u);
    assert.match(slideXml ?? "", /val="00FF88"/u);
    assert.match(slideXml ?? "", /val="FFAA00"/u);
    assert.match(slideXml ?? "", /val="FF3366"/u);
    assert.match(chartXml ?? "", /val="123456"/u);
    assert.match(chartXml ?? "", /val="ABCDEF"/u);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
