# Visual renderer third-party notices

- `dom-to-pptx` 2.1.1: MIT License
- `react` / `react-dom` 19.1.1: MIT License
- `react-icons` 5.7.0: MIT License
- Lucide icon set exposed through `react-icons/lu`: ISC License
- `pptxgenjs` 4.0.1: MIT License

The renderer exposes only a server-owned Lucide allowlist. Other icon packs bundled by the `react-icons` package are not selected or presented as approved assets.

PptxGenJS 4.0.1 declares `image-size` even though its distributed runtime does not call it. PowerPoint MCP resolves that dependency to the local `@pptx-mcp/image-size-disabled` fail-closed compatibility shim, which contains no image parsers and rejects every invocation. Image dimensions continue to come only from the existing server-verified PNG asset metadata.
