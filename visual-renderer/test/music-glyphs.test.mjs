import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  addBravuraGlyph,
  loadBravuraFont,
  MusicGlyph,
  timeSignatureGlyph,
} from "../music-glyphs.mjs";

test("Bravura treble clef becomes editable PowerPoint custom geometry", async () => {
  const font = await loadBravuraFont();
  const calls = [];
  const slide = {
    addShape(shapeType, options) {
      calls.push({ shapeType, options });
    },
  };

  const bounds = addBravuraGlyph(
    slide,
    { ShapeType: { custGeom: "custGeom" } },
    font,
    MusicGlyph.gClef,
    { originX: 1, originY: 2, scale: 0.0005, color: "123456" },
  );

  assert.equal(calls.length, 1);
  assert.equal(calls[0].shapeType, "custGeom");
  assert.equal(calls[0].options.fill.color, "123456");
  assert.equal(calls[0].options.line.type, "none");
  assert.match(calls[0].options.objectName, /^Bravura-E050-/);
  assert.ok(calls[0].options.points.some((point) => point.curve?.type === "cubic"));
  assert.ok(calls[0].options.points.some((point) => point.close === true));
  assert.ok(bounds.width > 0.3 && bounds.height > 0.8);
});

test("SMuFL time-signature digits use the stable Bravura range", () => {
  assert.equal(timeSignatureGlyph("0"), 0xE080);
  assert.equal(timeSignatureGlyph("4"), 0xE084);
  assert.equal(timeSignatureGlyph("9"), 0xE089);
  assert.throws(() => timeSignatureGlyph("x"), /Unsupported time-signature digit/);
});

test("Every exposed music symbol has non-empty native geometry", async () => {
  const font = await loadBravuraFont();
  const calls = [];
  const slide = { addShape: (shapeType, options) => calls.push({ shapeType, options }) };
  const pptx = { ShapeType: { custGeom: "custGeom" } };

  for (const codePoint of Object.values(MusicGlyph)) {
    const bounds = addBravuraGlyph(slide, pptx, font, codePoint, {
      centerX: 1,
      centerY: 1,
      scale: 0.0005,
      color: "202124",
    });
    assert.ok(bounds.width > 0);
    assert.ok(bounds.height > 0);
  }
  assert.equal(calls.length, Object.values(MusicGlyph).length);
  assert.ok(calls.every((call) => call.shapeType === "custGeom"));
});

test("Visual renderer only embeds server-verified image data with alt text", async () => {
  const source = await readFile(new URL("../index.mjs", import.meta.url), "utf8");
  assert.equal(source.match(/\.addImage\s*\(/gu)?.length, 2);
  assert.match(source, /data:\s*asset\.data/u);
  assert.match(source, /altText:\s*asset\.altText/u);
  assert.match(source, /renderApprovedReactIcon\(normalizedIcon/u);
  assert.match(source, /data:image\/svg\+xml;base64/u);
  assert.doesNotMatch(source, /addImage\s*\(\s*\{[^}]*\bpath\s*:/su);
});

test("Bundled Bravura asset matches the reviewed 1.392 release", async () => {
  const bytes = await readFile(new URL("../assets/bravura/Bravura.otf", import.meta.url));
  const digest = createHash("sha256").update(bytes).digest("hex");
  assert.equal(digest, "dca2d90c88437a701b1c2e71fa54e76f9fa41d7deee935d74dc871ea66ecfdd2");

  const license = await readFile(new URL("../assets/bravura/LICENSE.txt", import.meta.url), "utf8");
  assert.match(license, /SIL OPEN FONT LICENSE Version 1\.1/);
  assert.match(license, /Reserved Font Name \"Bravura\"/);
});
