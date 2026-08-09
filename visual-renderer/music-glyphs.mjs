import { readFile } from "node:fs/promises";
import opentype from "opentype.js";

const BRAVURA_PATH = new URL("./assets/bravura/Bravura.otf", import.meta.url);

export const MusicGlyph = Object.freeze({
  gClef: 0xE050,
  timeSig0: 0xE080,
  noteheadWhole: 0xE0A2,
  noteheadHalf: 0xE0A3,
  noteheadBlack: 0xE0A4,
  augmentationDot: 0xE1E7,
  flag8thUp: 0xE240,
  flag8thDown: 0xE241,
  flag16thUp: 0xE242,
  flag16thDown: 0xE243,
  accidentalFlat: 0xE260,
  accidentalNatural: 0xE261,
  accidentalSharp: 0xE262,
  restWhole: 0xE4E3,
  restHalf: 0xE4E4,
  restQuarter: 0xE4E5,
  rest8th: 0xE4E6,
  rest16th: 0xE4E7,
});

let glyphSequence = 0;

export async function loadBravuraFont() {
  const bytes = await readFile(BRAVURA_PATH);
  const arrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const font = opentype.parse(arrayBuffer);
  const familyName = font.names.windows?.fontFamily?.en ?? font.names.macintosh?.fontFamily?.en;
  if (familyName !== "Bravura") {
    throw new Error("The bundled SMuFL font is not Bravura.");
  }
  return font;
}

export function timeSignatureGlyph(digit) {
  if (!/^[0-9]$/.test(String(digit))) {
    throw new Error(`Unsupported time-signature digit: ${digit}`);
  }
  return MusicGlyph.timeSig0 + Number(digit);
}

export function addBravuraGlyph(slide, pptx, font, codePoint, options) {
  const scale = Number(options.scale);
  if (!Number.isFinite(scale) || scale <= 0) {
    throw new Error(`Invalid Bravura glyph scale: ${options.scale}`);
  }

  const geometry = glyphGeometry(font, codePoint, scale);
  let originX = Number(options.originX);
  let originY = Number(options.originY);
  if (Number.isFinite(options.centerX) && Number.isFinite(options.centerY)) {
    originX = Number(options.centerX) - geometry.centerX;
    originY = Number(options.centerY) - geometry.centerY;
  }
  if (!Number.isFinite(originX) || !Number.isFinite(originY)) {
    throw new Error("A Bravura glyph requires either originX/originY or centerX/centerY.");
  }

  const color = String(options.color ?? "202124").replace(/^#/, "");
  slide.addShape(pptx.ShapeType.custGeom, {
    x: originX + geometry.offsetX,
    y: originY + geometry.offsetY,
    w: geometry.width,
    h: geometry.height,
    points: geometry.points,
    fill: { color },
    line: { type: "none" },
    objectName: `Bravura-${codePoint.toString(16).toUpperCase()}-${++glyphSequence}`,
  });
  return {
    x: originX + geometry.offsetX,
    y: originY + geometry.offsetY,
    width: geometry.width,
    height: geometry.height,
  };
}

function glyphGeometry(font, codePoint, scale) {
  const glyph = font.charToGlyph(String.fromCodePoint(codePoint));
  if (!glyph || glyph.index === 0) {
    throw new Error(`Bravura does not contain SMuFL glyph U+${codePoint.toString(16).toUpperCase()}.`);
  }

  const path = glyph.getPath(0, 0, font.unitsPerEm);
  const bounds = path.getBoundingBox();
  const width = (bounds.x2 - bounds.x1) * scale;
  const height = (bounds.y2 - bounds.y1) * scale;
  if (!(width > 0) || !(height > 0)) {
    throw new Error(`Bravura glyph U+${codePoint.toString(16).toUpperCase()} has empty geometry.`);
  }

  const points = [];
  let contourOpen = false;
  for (const command of path.commands) {
    if (command.type === "M" && contourOpen) {
      points.push({ close: true });
      contourOpen = false;
    }
    points.push(commandToPoint(command, bounds, scale));
    contourOpen = command.type !== "Z";
  }
  if (contourOpen) points.push({ close: true });
  return {
    offsetX: round(bounds.x1 * scale),
    offsetY: round(bounds.y1 * scale),
    centerX: round((bounds.x1 + bounds.x2) * scale / 2),
    centerY: round((bounds.y1 + bounds.y2) * scale / 2),
    width: round(width),
    height: round(height),
    points,
  };
}

function commandToPoint(command, bounds, scale) {
  const x = (value) => round((value - bounds.x1) * scale);
  const y = (value) => round((value - bounds.y1) * scale);
  switch (command.type) {
    case "M":
      return { x: x(command.x), y: y(command.y), moveTo: true };
    case "L":
      return { x: x(command.x), y: y(command.y) };
    case "C":
      return {
        x: x(command.x),
        y: y(command.y),
        curve: {
          type: "cubic",
          x1: x(command.x1),
          y1: y(command.y1),
          x2: x(command.x2),
          y2: y(command.y2),
        },
      };
    case "Q":
      return {
        x: x(command.x),
        y: y(command.y),
        curve: {
          type: "quadratic",
          x1: x(command.x1),
          y1: y(command.y1),
        },
      };
    case "Z":
      return { close: true };
    default:
      throw new Error(`Unsupported OpenType path command: ${command.type}`);
  }
}

function round(value) {
  return Number(value.toFixed(6));
}
