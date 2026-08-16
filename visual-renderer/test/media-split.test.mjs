import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import test from "node:test";
import JSZip from "jszip";
import sharp from "sharp";

const executeFile = promisify(execFile);
const rendererPath = fileURLToPath(new URL("../index.mjs", import.meta.url));
const rendererDirectory = dirname(rendererPath);

test("Media split embeds a verified PNG with alt text and focal crop", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-media-split-"));
  const assetRoot = join(workingDirectory, "image-assets");
  const assetId = "0123456789abcdef0123456789abcdef";
  const assetDirectory = join(assetRoot, assetId);
  const imagePath = join(assetDirectory, "asset.png");
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "media.pptx");

  try {
    await mkdir(assetDirectory, { recursive: true });
    await sharp({
      create: {
        width: 800,
        height: 400,
        channels: 3,
        background: { r: 20, g: 130, b: 210 },
      },
    }).png().toFile(imagePath);
    const image = await readFile(imagePath);
    const specification = {
      title: "Verified image deck",
      rendererContract: "visual-v5",
      design: { style: "executive", density: "balanced", motif: "none" },
      imageAssets: {
        [assetId]: {
          width: 800,
          height: 400,
          altText: "Blue product overview illustration",
          sha256: createHash("sha256").update(image).digest("hex"),
        },
      },
      slides: [{
        kind: "Media",
        title: "Product overview",
        body: "A verified image is embedded rather than linked.",
        bullets: ["No external relationship", "Accessible alt text"],
        variant: "split",
        media: {
          assetId,
          cropIntent: "focalLeft",
          textPosition: "left",
          caption: "User-provided approved image",
        },
        attribution: "SOURCE-001",
      }],
    };
    await writeFile(specificationPath, JSON.stringify(specification), "utf8");
    await executeFile(process.execPath, [rendererPath, specificationPath, outputPath], {
      cwd: rendererDirectory,
      env: { ...process.env, PPTX_MCP_IMAGE_ASSET_ROOT: assetRoot },
      timeout: 30_000,
    });

    const archive = await JSZip.loadAsync(await readFile(outputPath));
    const slideXml = await archive.file("ppt/slides/slide1.xml")?.async("string");
    const relsXml = await archive.file("ppt/slides/_rels/slide1.xml.rels")?.async("string");
    assert.match(slideXml ?? "", /descr="Blue product overview illustration"/u);
    assert.match(slideXml ?? "", /<a:srcRect l="0"/u);
    assert.match(slideXml ?? "", />SOURCE-001</u);
    assert.match(relsXml ?? "", /relationships\/image/u);
    assert.doesNotMatch(relsXml ?? "", /TargetMode="External"/u);
    assert.ok(Object.keys(archive.files).some((name) => name.startsWith("ppt/media/image-")));
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});

test("Media renderer rejects an asset whose bytes do not match the server hash", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-media-integrity-"));
  const assetRoot = join(workingDirectory, "image-assets");
  const assetId = "fedcba9876543210fedcba9876543210";
  const assetDirectory = join(assetRoot, assetId);
  const specificationPath = join(workingDirectory, "visual-deck.json");
  try {
    await mkdir(assetDirectory, { recursive: true });
    await sharp({
      create: { width: 32, height: 32, channels: 3, background: "#ff0000" },
    }).png().toFile(join(assetDirectory, "asset.png"));
    await writeFile(specificationPath, JSON.stringify({
      title: "Rejected image",
      rendererContract: "visual-v5",
      imageAssets: {
        [assetId]: {
          width: 32,
          height: 32,
          altText: "Red square",
          sha256: "0".repeat(64),
        },
      },
      slides: [{
        kind: "Media",
        title: "Rejected",
        variant: "split",
        media: { assetId },
      }],
    }), "utf8");

    await assert.rejects(
      executeFile(process.execPath, [rendererPath, specificationPath, join(workingDirectory, "out.pptx")], {
        cwd: rendererDirectory,
        env: { ...process.env, PPTX_MCP_IMAGE_ASSET_ROOT: assetRoot },
        timeout: 30_000,
      }),
      /integrity/u,
    );
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});

test("Media renderer rejects a non-PNG asset even when its server hash matches", async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-media-format-"));
  const assetRoot = join(workingDirectory, "image-assets");
  const assetId = "1234567890abcdef1234567890abcdef";
  const assetDirectory = join(assetRoot, assetId);
  const specificationPath = join(workingDirectory, "visual-deck.json");
  try {
    await mkdir(assetDirectory, { recursive: true });
    const bytes = Buffer.from("icns-untrusted-image-data", "utf8");
    await writeFile(join(assetDirectory, "asset.png"), bytes);
    await writeFile(specificationPath, JSON.stringify({
      title: "Rejected format",
      rendererContract: "visual-v5",
      imageAssets: {
        [assetId]: {
          width: 32,
          height: 32,
          altText: "Invalid image",
          sha256: createHash("sha256").update(bytes).digest("hex"),
        },
      },
      slides: [{
        kind: "Media",
        title: "Rejected",
        variant: "split",
        media: { assetId },
      }],
    }), "utf8");

    await assert.rejects(
      executeFile(process.execPath, [rendererPath, specificationPath, join(workingDirectory, "out.pptx")], {
        cwd: rendererDirectory,
        env: { ...process.env, PPTX_MCP_IMAGE_ASSET_ROOT: assetRoot },
        timeout: 30_000,
      }),
      /normalized PNG/u,
    );
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
