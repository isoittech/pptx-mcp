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

test("nodes motif emits non-negative Open XML extents for an upward connector", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-node-motif-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "node-motif.pptx");

  try {
    await writeFile(
      specificationPath,
      JSON.stringify({
        title: "Daily IT news",
        rendererContract: "visual-v5",
        design: { style: "technical", density: "detailed", motif: "nodes" },
        slides: [
          {
            kind: "Title",
            title: "Daily IT news",
            subtitle: "Japan and global",
            density: "balanced",
          },
          {
            kind: "Agenda",
            title: "Report flow",
            density: "balanced",
            bullets: ["Summary", "Global", "Japan", "Metrics", "Actions", "Sources"],
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
    const slideXml = await archive.file("ppt/slides/slide2.xml")?.async("string") ?? "";
    const extents = [...slideXml.matchAll(/<a:ext cx="(-?\d+)" cy="(-?\d+)"/gu)];
    assert.ok(extents.length > 0);
    assert.ok(extents.every((match) => Number(match[1]) >= 0 && Number(match[2]) >= 0));
    assert.match(slideXml, /<a:xfrm flipV="1"><a:off x="10716768" y="201168"\/><a:ext cx="566928" cy="457200"\/>/u);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
