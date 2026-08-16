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

test("speaker notes are stored outside the visible slide canvas", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-speaker-notes-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "speaker-notes.pptx");
  const purpose = "Approve the pilot & its success criteria.";
  const talkScript = "First explain the context.\nThen ask for the decision <today>.";

  try {
    await writeFile(
      specificationPath,
      JSON.stringify({
        title: "Speaker notes",
        rendererContract: "visual-v5",
        slides: [
          {
            kind: "Title",
            title: "Visible title only",
            speakerNotes: { purpose, talkScript },
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
    const slideXml = await archive.file("ppt/slides/slide1.xml")?.async("string") ?? "";
    const notesXml = await archive.file("ppt/notesSlides/notesSlide1.xml")?.async("string") ?? "";

    assert.match(slideXml, /Visible title only/u);
    assert.doesNotMatch(slideXml, /Approve the pilot|First explain/u);
    assert.match(notesXml, /このスライドの狙い/u);
    assert.match(notesXml, /Approve the pilot &amp; its success criteria\./u);
    assert.match(notesXml, /トークスクリプト/u);
    assert.match(notesXml, /Then ask for the decision &lt;today&gt;\./u);
    assert.ok(archive.file("ppt/notesMasters/notesMaster1.xml"));
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
