import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import test from "node:test";
import JSZip from "jszip";
import {
  buildDomDeckHtml,
  canRenderDeckWithDom,
  domSupportedSlideKinds,
} from "../dom-renderer.mjs";
import { approvedReactIcons, renderApprovedReactIcon } from "../react-icons.mjs";

const executeFile = promisify(execFile);
const rendererPath = fileURLToPath(new URL("../index.mjs", import.meta.url));
const rendererDirectory = dirname(rendererPath);

test("the DOM renderer supports the bounded business-layout allowlist", () => {
  assert.ok(domSupportedSlideKinds.has("cards"));
  assert.ok(domSupportedSlideKinds.has("media"));
  assert.ok(domSupportedSlideKinds.has("coverage_map"));
  assert.ok(domSupportedSlideKinds.has("transformation_evidence"));
  assert.ok(domSupportedSlideKinds.has("artifact_showcase"));
  assert.ok(domSupportedSlideKinds.has("gantt_schedule"));
  assert.equal(canRenderDeckWithDom({ slides: [{ kind: "Cards" }, { kind: "Metrics" }] }), true);
  assert.equal(canRenderDeckWithDom({ slides: [{ kind: "Chart" }] }), false);
  assert.equal(canRenderDeckWithDom({ slides: [{ kind: "MusicScore" }] }), false);
});

test("react-icons are rendered only through the approved Lucide allowlist", () => {
  for (const iconName of Object.keys(approvedReactIcons)) {
    const markup = renderApprovedReactIcon(iconName, { color: "#123456" });
    assert.match(markup, /^<svg/u);
    assert.match(markup, /color="#123456"/u);
    assert.doesNotMatch(markup, /<script|javascript:|file:/u);
  }
  assert.throws(
    () => renderApprovedReactIcon("not-approved"),
    /Unsupported react-icons identifier/u,
  );
});

test("server-generated slide HTML escapes content and carries reflection metadata", () => {
  const html = buildDomDeckHtml({
    title: "安全な資料",
    language: "ja-JP",
    templateChrome: true,
    rendererContract: "visual-v6-dom",
    theme: { preset: "minimal" },
    slides: [{
      kind: "Cards",
      title: "<script src=https://invalid.example></script>",
      cards: [{
        title: "判断 & 実行",
        description: "ローカルで安全に描画",
        icon: "decision",
      }],
      speakerNotes: {
        purpose: "判断基準を共有する",
        talkScript: "このページでは判断基準を説明します。",
      },
    }],
  });

  assert.match(html, /class="slide template-chrome"/u);
  assert.match(html, /<template data-pptx-notes>このスライドの狙い\n判断基準を共有する/u);
  assert.match(html, /&lt;script src=https:\/\/invalid\.example&gt;/u);
  assert.doesNotMatch(html, /<script src=https:\/\/invalid\.example>/u);
  assert.match(html, /<svg/u);
  assert.match(html, /\.footer>span:last-child\{[^}]*width:88px;[^}]*white-space:nowrap/u);
  assert.match(html, /\.takeaway p\{flex:1;min-width:0;/u);
  assert.doesNotMatch(html, /class="footer"/u);
  assert.match(html, /box-shadow:none!important;text-shadow:none!important/u);
});

test("the configured default template cover uses a matte white-text overlay without decorative numbering", () => {
  const html = buildDomDeckHtml({
    title: "既定表紙",
    templateChrome: true,
    defaultTemplateCoverOverlay: true,
    rendererContract: "visual-v6-dom",
    theme: { preset: "minimal" },
    slides: [{
      kind: "Section",
      title: "社内研修資料",
      subtitle: "副題",
      body: "写真背景の上でも読める導入文",
    }],
  });

  assert.match(html, /class="slide template-chrome default-template-cover"/u);
  assert.match(html, /\.default-template-cover header h1\{[^}]*color:#fff/u);
  assert.match(html, /\.default-template-cover \.section-number\{display:none\}/u);
  assert.match(html, /class="section-number">01</u);
  assert.match(html, /box-shadow:none!important;text-shadow:none!important/u);
});

test("a user-supplied template does not receive the default-cover overlay", () => {
  const html = buildDomDeckHtml({
    title: "ユーザー指定表紙",
    templateChrome: true,
    rendererContract: "visual-v6-dom",
    theme: { preset: "minimal" },
    slides: [{ kind: "Section", title: "ユーザー指定表紙", body: "本文" }],
  });

  assert.match(html, /class="slide template-chrome"/u);
  assert.doesNotMatch(html, /class="slide template-chrome default-template-cover"/u);
});

test("quality components render semantic, matte DOM without accepting raw HTML", () => {
  const html = buildDomDeckHtml({
    title: "品質部品カタログ",
    rendererContract: "visual-v6-dom",
    theme: { preset: "minimal" },
    slides: [
      {
        kind: "CoverageMap",
        title: "領域と時点の対応",
        coverageMap: {
          columns: [
            { id: "plan", label: "企画" },
            { id: "build", label: "実行" },
            { id: "operate", label: "運用" },
          ],
          groups: [{
            id: "security",
            label: "安全性",
            rows: [{ id: "review", label: "レビュー" }],
          }],
          bars: [{ id: "review-span", rowId: "review", label: "継続確認", startColumn: 1, endColumn: 3 }],
          footerChips: [{ label: "標準確認", tone: "accent" }],
        },
      },
      {
        kind: "TransformationEvidence",
        title: "入力と検証根拠",
        transformationEvidence: {
          inputHeading: "入力",
          inputSegments: [
            { text: "村田", tag: "人物A", tone: "warning" },
            { text: " <script>alert(1)</script>" },
          ],
          outputHeading: "変換後",
          outputText: "[人物A] に置換",
          evidenceTable: {
            columns: [{ header: "タグ" }, { header: "値" }],
            rows: [{ cells: [{ text: "人物A" }, { text: "村田" }] }],
          },
        },
      },
      {
        kind: "ArtifactShowcase",
        title: "成果物",
        artifactShowcase: {
          groups: [{
            title: "報告書",
            artifacts: [{ assetId: "img_sample", label: "最終版" }],
          }],
        },
      },
      {
        kind: "GanttSchedule",
        title: "実施計画",
        ganttSchedule: {
          columns: [
            { id: "w1", label: "W1", groupLabel: "1月" },
            { id: "w2", label: "W2", groupLabel: "1月" },
            { id: "w3", label: "W3", groupLabel: "1月" },
            { id: "w4", label: "W4", groupLabel: "1月" },
          ],
          tasks: [
            { id: "design", category: "設計", title: "要件合意", startColumn: 1, endColumn: 2 },
            { id: "build", category: "実装", title: "試作", startColumn: 2, endColumn: 4, tone: "positive" },
          ],
        },
      },
    ],
  }, {
    img_sample: {
      data: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
      altText: "報告書のプレビュー",
    },
  });

  assert.match(html, /class="coverage-grid"/u);
  assert.match(html, /class="transformation-grid"/u);
  assert.match(html, /class="artifact-grid count-1"/u);
  assert.match(html, /class="gantt-grid"/u);
  assert.match(html, /&lt;script&gt;alert\(1\)&lt;\/script&gt;/u);
  assert.doesNotMatch(html, /<script>alert\(1\)<\/script>/u);
  assert.doesNotMatch(html, /box-shadow:(?!none)/u);
  assert.match(
    html,
    /\.coverage-axis span,\.coverage-group span,\.gantt-axis span,\.gantt-task span,\.gantt-marker span,\.tagged-segment b\{font-size:19px\}/u,
  );
});

test("dom-to-pptx exports editable text, react-icons SVG, and speaker notes", {
  skip: process.env.PPTX_MCP_RUN_DOM_INTEGRATION !== "1",
  timeout: 90_000,
}, async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-dom-renderer-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "dom-deck.pptx");
  try {
    await writeFile(specificationPath, JSON.stringify({
      title: "DOM renderer integration",
      rendererContract: "visual-v6-dom",
      theme: { preset: "minimal" },
      slides: [
        {
          kind: "Title",
          title: "HTMLから編集可能なPowerPointへ",
          subtitle: "dom-to-pptx 2.1.1",
          body: "サーバー管理のHTMLとCSSだけを変換します。",
          speakerNotes: {
            purpose: "新しい生成経路を紹介する。",
            talkScript: "HTMLとCSSで構成したスライドを編集可能なPowerPointへ変換します。",
          },
        },
        {
          kind: "Cards",
          title: "承認済みアイコンを使う",
          cards: [
            { title: "判断", description: "意味に合うアイコン", icon: "decision" },
            { title: "安全", description: "サーバー側allowlist", icon: "shield" },
            { title: "実行", description: "PPTXへベクター出力", icon: "automation" },
          ],
          speakerNotes: {
            purpose: "react-iconsの安全な利用境界を伝える。",
            talkScript: "モデルはアイコンIDだけを選び、サーバーが承認済みLucideアイコンへ解決します。",
          },
        },
        {
          kind: "Process",
          title: "全画像確認までを標準化する",
          steps: [
            { label: "STEP 1", title: "生成", description: "編集可能な要素へ変換します。" },
            { label: "STEP 2", title: "合成", description: "会社テンプレートへ合成します。" },
            { label: "STEP 3", title: "確認", description: "全ページを画像で確認します。" },
          ],
          takeaway: "承認後は4工程を標準手順として固定",
          speakerNotes: {
            purpose: "品質確認フローを共有する。",
            talkScript: "生成から全ページ画像確認までを標準手順として説明します。",
          },
        },
        {
          kind: "Bullets",
          title: "標準の箇条書きを使う",
          bullets: ["第一の論点", "第二の論点"],
          speakerNotes: {
            purpose: "箇条書きの編集可能性を確認する。",
            talkScript: "箇条書きはPowerPoint標準の段落として出力します。",
          },
        },
      ],
    }), "utf8");

    const execution = await executeFile(process.execPath, [rendererPath, specificationPath, outputPath], {
      cwd: rendererDirectory,
      timeout: 85_000,
    });
    assert.match(execution.stdout, /PPTX_MCP_RENDERER=dom-to-pptx@2\.1\.1\+react-icons@5\.7\.0/u);

    const archive = await JSZip.loadAsync(await readFile(outputPath));
    const slide1 = await archive.file("ppt/slides/slide1.xml")?.async("string") ?? "";
    const slide2 = await archive.file("ppt/slides/slide2.xml")?.async("string") ?? "";
    const slide3 = await archive.file("ppt/slides/slide3.xml")?.async("string") ?? "";
    const slide4 = await archive.file("ppt/slides/slide4.xml")?.async("string") ?? "";
    const notes = await archive.file("ppt/notesSlides/notesSlide1.xml")?.async("string") ?? "";
    assert.match(slide1, /HTMLから編集可能なPowerPointへ/u);
    assert.match(slide2, /承認済みアイコンを使う/u);
    assert.match(notes, /新しい生成経路を紹介する/u);
    assert.ok((slide2.match(/<p:sp>/gu) ?? []).length >= 3);
    const takeawayShape = (slide3.match(/<p:sp>.*?<\/p:sp>/gu) ?? [])
      .find((shape) => shape.includes("承認後は4工程を標準手順として固定"));
    assert.ok(takeawayShape, "the Process takeaway must remain editable text");
    const takeawayExtent = takeawayShape.match(/<a:ext cx="(\d+)" cy="\d+"/u);
    assert.ok(Number(takeawayExtent?.[1]) >= 6_000_000, "the takeaway text box must use the remaining flex width");
    assert.match(slide4, /<a:buChar char="•"\/>/u);
    const bulletShape = (slide4.match(/<p:sp>.*?<\/p:sp>/gu) ?? [])
      .find((shape) => shape.includes("第一の論点"));
    assert.ok(bulletShape, "the DOM bullet must remain editable text");
    const bulletSizes = [...bulletShape.matchAll(/\bsz="(\d+)"/gu)].map((match) => Number(match[1]));
    assert.ok(bulletSizes.length > 0 && Math.min(...bulletSizes) >= 1_400);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
