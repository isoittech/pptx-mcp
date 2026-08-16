import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import test from "node:test";
import sharp from "sharp";

const executeFile = promisify(execFile);
const sanitizerPath = fileURLToPath(new URL("../sanitize-image.mjs", import.meta.url));

test("image sanitizer rotates, bounds, converts to PNG, and strips metadata", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-image-sanitizer-"));
  const inputPath = join(workingDirectory, "input.jpg");
  const outputPath = join(workingDirectory, "output.png");
  try {
    await sharp({
      create: { width: 1200, height: 600, channels: 3, background: "#0088cc" },
    })
      .jpeg()
      .withMetadata({ orientation: 6, exif: { IFD0: { Artist: "must-not-survive" } } })
      .toFile(inputPath);

    const { stdout } = await executeFile(process.execPath, [
      sanitizerPath,
      inputPath,
      outputPath,
      "20000000",
      "800",
      "12582912",
    ], { cwd: dirname(sanitizerPath), timeout: 30_000 });
    const result = JSON.parse(stdout);
    const metadata = await sharp(outputPath).metadata();
    const bytes = await readFile(outputPath);

    assert.equal(metadata.format, "png");
    assert.ok(result.width <= 800 && result.height <= 800);
    assert.equal(result.bytes, bytes.length);
    assert.equal(metadata.exif, undefined);
    assert.equal(metadata.icc, undefined);
    assert.equal(metadata.xmp, undefined);
    assert.doesNotMatch(bytes.toString("latin1"), /must-not-survive/u);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});

test("image sanitizer rejects undecodable input", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-image-sanitizer-invalid-"));
  const inputPath = join(workingDirectory, "input.png");
  try {
    await writeFile(inputPath, "not an image", "utf8");
    await assert.rejects(
      executeFile(process.execPath, [
        sanitizerPath,
        inputPath,
        join(workingDirectory, "output.png"),
        "20000000",
        "800",
        "12582912",
      ], { cwd: dirname(sanitizerPath), timeout: 30_000 }),
    );
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});

test("image sanitizer rejects input whose decoded pixels exceed the configured cap", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-image-sanitizer-pixels-"));
  const inputPath = join(workingDirectory, "input.png");
  try {
    await sharp({
      create: { width: 100, height: 100, channels: 3, background: "#ffffff" },
    }).png().toFile(inputPath);
    await assert.rejects(
      executeFile(process.execPath, [
        sanitizerPath,
        inputPath,
        join(workingDirectory, "output.png"),
        "5000",
        "800",
        "12582912",
      ], { cwd: dirname(sanitizerPath), timeout: 30_000 }),
    );
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
