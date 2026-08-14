import { stat } from "node:fs/promises";
import sharp from "sharp";

const [inputPath, outputPath, maxPixelsText, maxDimensionText, maxBytesText] = process.argv.slice(2);
if (!inputPath || !outputPath) {
  throw new Error(
    "Usage: node sanitize-image.mjs <input> <output.png> <maxPixels> <maxDimension> <maxBytes>",
  );
}

const maxPixels = positiveInteger(maxPixelsText, "maxPixels");
const maxDimension = positiveInteger(maxDimensionText, "maxDimension");
const maxBytes = positiveInteger(maxBytesText, "maxBytes");
const input = sharp(inputPath, {
  animated: false,
  failOn: "error",
  limitInputPixels: maxPixels,
  sequentialRead: true,
});
const metadata = await input.metadata();
if (!metadata.width || !metadata.height || !["jpeg", "png"].includes(metadata.format)) {
  throw new Error("Only decodable single-frame JPEG and PNG images are accepted.");
}
if ((metadata.pages ?? 1) !== 1 || metadata.width * metadata.height > maxPixels) {
  throw new Error("The image dimensions or frame count exceed the configured limits.");
}

const result = await input
  .rotate()
  .resize({
    width: maxDimension,
    height: maxDimension,
    fit: "inside",
    withoutEnlargement: true,
  })
  .toColourspace("srgb")
  .png({ compressionLevel: 9, palette: false, progressive: false })
  .toFile(outputPath);
const outputStat = await stat(outputPath);
if (outputStat.size <= 0 || outputStat.size > maxBytes) {
  throw new Error("The normalized PNG exceeds the configured output size limit.");
}

process.stdout.write(
  JSON.stringify({ width: result.width, height: result.height, bytes: outputStat.size }),
);

function positiveInteger(value, name) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed <= 0) {
    throw new Error(`${name} must be a positive integer.`);
  }
  return parsed;
}
