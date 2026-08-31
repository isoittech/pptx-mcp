import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";

const require = createRequire(import.meta.url);

test("the unused PptxGenJS image-size dependency is replaced by a fail-closed shim", async () => {
  const imageSizePackage = require("image-size/package.json");
  const commonJsImageSize = require("image-size");
  const esmImageSize = await import("image-size");

  assert.equal(imageSizePackage.name, "@pptx-mcp/image-size-disabled");
  assert.deepEqual(commonJsImageSize.types, []);
  assert.throws(
    () => commonJsImageSize(Buffer.from("icns\0\0\0\0", "binary")),
    /image-size is disabled in PowerPoint MCP/u,
  );
  assert.throws(
    () => commonJsImageSize("/untrusted/input.heic"),
    /image-size is disabled in PowerPoint MCP/u,
  );
  assert.throws(
    () => esmImageSize.default(Uint8Array.from([0xff, 0x0a, 0x00, 0x00])),
    /image-size is disabled in PowerPoint MCP/u,
  );
});
