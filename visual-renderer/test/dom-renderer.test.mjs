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
  canRenderSlideWithDom,
  domSupportedSlideKinds,
  layoutDirectedDiagram,
  normalizeDomListTextInsets,
  sanitizeAuthoredHtmlFragment,
  validateAndScopeAuthoredCss,
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
  assert.equal(canRenderSlideWithDom({ kind: "NativeDiagram", diagram: { kind: "Tree" } }), true);
  assert.equal(canRenderSlideWithDom({ kind: "NativeDiagram", diagram: { kind: "Flow" } }), true);
  assert.equal(canRenderSlideWithDom({ kind: "NativeDiagram", diagram: { kind: "Cycle" } }), false);
  assert.equal(canRenderDeckWithDom({ slides: [{ kind: "NativeDiagram", diagram: { kind: "Tree" } }] }), true);
});

test("DOM list inset normalization keeps list left padding on the left edge", async () => {
  const archive = new JSZip();
  archive.file("ppt/slides/slide1.xml", `<p:sld><p:cSld><p:spTree>
    <p:sp><p:txBody><a:bodyPr lIns="0" tIns="198120" rIns="0" bIns="0"/><a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>箇条書き</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:txBody><a:bodyPr lIns="111" tIns="222" rIns="0" bIns="0"/><a:p><a:r><a:t>通常本文</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld></p:sld>`);
  const input = await archive.generateAsync({ type: "nodebuffer" });

  const output = await normalizeDomListTextInsets(input);
  const normalizedArchive = await JSZip.loadAsync(output);
  const slide = await normalizedArchive.file("ppt/slides/slide1.xml")?.async("string") ?? "";
  const listShape = (slide.match(/<p:sp>[\s\S]*?<\/p:sp>/gu) ?? [])
    .find((shape) => shape.includes("箇条書き"));
  const bodyShape = (slide.match(/<p:sp>[\s\S]*?<\/p:sp>/gu) ?? [])
    .find((shape) => shape.includes("通常本文"));

  assert.match(listShape ?? "", /<a:bodyPr[^>]*\blIns="198120"[^>]*\btIns="0"/u);
  assert.match(bodyShape ?? "", /<a:bodyPr[^>]*\blIns="111"[^>]*\btIns="222"/u);
});

test("the default company body style separates a 30pt title from a 16pt claim bullet", () => {
  const html = buildDomDeckHtml({
    title: "会社テンプレート",
    templateChrome: true,
    defaultTemplateBodyStyle: true,
    rendererContract: "visual-v6-dom",
    theme: { preset: "minimal", secondaryColor: "005B96" },
    slides: [
      { kind: "Title", title: "表紙" },
      {
        kind: "NativeDiagram",
        title: "NDAは「契約」であり結論は条文次第",
        subtitle: "共通して問題になるのは目的外利用と第三者開示の2条項",
        diagram: {
          kind: "Tree",
          direction: "TopToBottom",
          nodes: [
            { id: "root", label: "前提" },
            { id: "left", label: "目的外利用" },
            { id: "right", label: "第三者開示" },
          ],
          edges: [
            { from: "root", to: "left" },
            { from: "root", to: "right" },
          ],
        },
      },
    ],
  });

  assert.match(html, /class="slide template-chrome default-template-body has-header-claim"/u);
  assert.match(html, /<h1>NDAは「契約」であり結論は条文次第<\/h1><ul class="header-claim"><li>共通して問題になるのは目的外利用と第三者開示の2条項<\/li><\/ul>/u);
  assert.match(html, /\.slide\.default-template-body header h1\{font-size:50px;color:var\(--secondary\);font-weight:800\}/u);
  assert.match(html, /\.header-claim\{[^}]*font-size:27px/u);
});

test("directed-diagram layout preserves visible whitespace between independent nodes", () => {
  const diagram = {
    kind: "Tree",
    direction: "LeftToRight",
    nodes: [
      { id: "start", label: "分かれ目" },
      { id: "a", label: "管理の態様" },
      { id: "b", label: "守るべき価値" },
      { id: "c", label: "AIツールの機能" },
      { id: "end", label: "共通リスク" },
    ],
    edges: [
      { from: "start", to: "a" },
      { from: "start", to: "b" },
      { from: "start", to: "c" },
      { from: "a", to: "end" },
      { from: "b", to: "end" },
      { from: "c", to: "end" },
    ],
  };
  const { positions } = layoutDirectedDiagram(diagram);
  const middleNodes = [positions.get("a"), positions.get("b"), positions.get("c")];
  for (let index = 1; index < middleNodes.length; index += 1) {
    const previous = middleNodes[index - 1];
    const current = middleNodes[index];
    assert.ok(current.y - (previous.y + previous.height) >= 32);
  }
  assert.ok(positions.get("a").x - (positions.get("start").x + positions.get("start").width) >= 32);
  assert.ok(positions.get("end").x - (positions.get("a").x + positions.get("a").width) >= 32);
});

test("model-authored HTML/CSS remains the slide design while the server scopes and resolves safe components", () => {
  const html = buildDomDeckHtml({
    title: "AI HTMLデザイン",
    rendererContract: "visual-v7-author-html",
    templateChrome: true,
    theme: { preset: "minimal", secondaryColor: "005B96" },
    slides: [{
      kind: "NativeDiagram",
      title: "NDAは「契約」であり結論は条文次第",
      authoredHtml: {
        html: `<div class="page"><h1>NDAは「契約」であり結論は条文次第</h1><ul class="claim"><li>共通して問題になるのは目的外利用と第三者開示の2条項</li></ul><div class="flow"><article class="node emphasis">前提</article><div class="arrow" data-pptx-icon="arrow" aria-label="次へ"></div><article class="node">2つの論点</article></div></div>`,
        css: `.slide{background:#ff0000;padding:68px 78px}.slide h1{margin:0;color:var(--secondary);font-size:50px;line-height:1.1}.slide .claim{color:var(--secondary);font-size:27px}.slide .flow{display:flex;align-items:center;justify-content:center;gap:64px;margin-top:80px}.slide .node{width:300px;min-height:120px;border:2px solid var(--secondary);padding:24px;font-size:24px}.slide .emphasis{background:var(--secondary);color:white}.slide .arrow{width:60px;height:60px;color:var(--secondary)}`,
      },
      speakerNotes: { purpose: "論点を分ける", talkScript: "二つの論点を説明します。" },
    }],
  });

  assert.match(html, /data-author-slide="1"/u);
  assert.match(html, /NDAは「契約」であり結論は条文次第/u);
  assert.match(html, /\.slide\[data-author-slide="1"\] h1\{margin:0;color:var\(--secondary\);font-size:50px/u);
  assert.match(html, /class="arrow pptx-icon"/u);
  assert.match(html, /<svg/u);
  assert.doesNotMatch(html, /data-pptx-icon/u);
  assert.doesNotMatch(html, /class="accent-rail"/u);
  assert.match(html, /\.slide\[data-author-slide="1"\]\{background:transparent!important\}/u);
});

test("model-authored HTML accepts semantic quotation blocks", () => {
  const html = sanitizeAuthoredHtmlFragment(
    `<section><h2>契約条文</h2><blockquote class="quote">開示された機密情報は本検討の範囲内で使用する。</blockquote></section>`,
  );

  assert.match(html, /<blockquote class="quote">/u);
});

test("model-authored HTML accepts safe table column sizing markup", () => {
  const html = sanitizeAuthoredHtmlFragment(
    `<table><colgroup><col style="width:25%"><col style="width:75%"></colgroup><tbody><tr><td>項目</td><td>説明</td></tr></tbody></table>`,
  );

  assert.match(html, /<colgroup>/u);
  assert.match(html, /<col style="width:25%">/u);
});

test("model-authored HTML/CSS rejects executable markup, unmanaged resources, and unscoped rules", () => {
  assert.throws(
    () => sanitizeAuthoredHtmlFragment(`<div onclick="alert(1)">危険</div>`),
    /executable markup/u,
  );
  assert.throws(
    () => sanitizeAuthoredHtmlFragment(`<img src="https://invalid.example/a.png" alt="危険">`),
    /remote resources|resource attribute/u,
  );
  assert.throws(
    () => sanitizeAuthoredHtmlFragment(`<svg><path d="M0 0L1 1"/></svg>`),
    /executable markup|tag is not allowed/u,
  );
  assert.throws(
    () => validateAndScopeAuthoredCss(`body{background:red}`, 0),
    /scoped below \.slide/u,
  );
  assert.throws(
    () => validateAndScopeAuthoredCss(`.slide{background-image:url(https://invalid.example/a.png)}`, 0),
    /remote resources|disallowed CSS/u,
  );
  assert.throws(
    () => validateAndScopeAuthoredCss(`.slide .card{box-shadow:0 2px 8px #000}`, 0),
    /disallowed CSS/u,
  );
});

test("model-authored HTML enforces the default body title and claim contract", () => {
  const options = {
    requireDefaultBodyContract: true,
    expectedBodyTitle: "NDAは契約であり、結論は条文次第",
    expectedBodyClaim: "共通論点は目的外利用と第三者開示",
  };
  const valid = sanitizeAuthoredHtmlFragment(
    `<div><h2 data-pptx-role="BODY-TITLE">NDAは契約であり、結論は条文次第</h2><ul data-pptx-role="BODY-CLAIM"><li>共通論点は目的外利用と第三者開示</li></ul></div>`,
    [],
    {},
    options,
  );

  assert.match(valid, /data-pptx-role="body-title"/u);
  assert.match(valid, /data-pptx-role="body-claim"/u);
  assert.throws(
    () => sanitizeAuthoredHtmlFragment(`<div><h2>題名</h2><p>主張</p></div>`, [], {}, options),
    /body-title/u,
  );
  assert.throws(
    () => sanitizeAuthoredHtmlFragment(
      `<div><h2 data-pptx-role="body-title">別の題名</h2><ul data-pptx-role="body-claim"><li>共通論点は目的外利用と第三者開示</li></ul></div>`,
      [],
      {},
      options,
    ),
    /exactly match/u,
  );
});

test("model-authored CSS enforces the 24px floor with a bounded 20px source exception", () => {
  assert.throws(
    () => validateAndScopeAuthoredCss(`.slide .body{font-size:23px}`, 0),
    /at least 24px/u,
  );
  assert.throws(
    () => validateAndScopeAuthoredCss(`.slide .body{font-size:1.5rem}`, 0),
    /literal px/u,
  );
  assert.throws(
    () => validateAndScopeAuthoredCss(`.slide .source{font-size:20px}`, 0),
    /at least 24px/u,
  );
  assert.throws(
    () => validateAndScopeAuthoredCss(`.slide [data-pptx-role="source-meta"]{font-size:19px}`, 0),
    /at least 20px/u,
  );
  assert.throws(
    () => validateAndScopeAuthoredCss(`.slide [data-pptx-role="source-meta"] span{font-size:20px}`, 0),
    /at least 24px/u,
  );
  assert.match(
    validateAndScopeAuthoredCss(`.slide [data-pptx-role="source-meta"]{font-size:20px}`, 2),
    /data-author-slide="3"/u,
  );
  assert.doesNotThrow(() => sanitizeAuthoredHtmlFragment(
    `<small data-pptx-role="source-meta" style="font-size:20px">出典</small>`,
  ));
  assert.throws(
    () => validateAndScopeAuthoredCss(`.slide [data-pptx-role="body-title"]{font-size:34px}`, 0),
    /body-title must be exactly 50px/u,
  );
  assert.throws(
    () => validateAndScopeAuthoredCss(`.slide [data-pptx-role="body-claim"],.slide [data-pptx-role="body-claim"] li{font-size:24px}`, 0),
    /body-claim must be exactly 27px/u,
  );
  assert.throws(
    () => sanitizeAuthoredHtmlFragment(
      `<h2 data-pptx-role="body-title" style="font-size:40px">題名</h2>`,
    ),
    /body-title must be exactly 50px/u,
  );
  assert.doesNotThrow(() => validateAndScopeAuthoredCss(
    `.slide [data-pptx-role="body-title"]{font-size:50px}.slide [data-pptx-role="body-claim"],.slide [data-pptx-role="body-claim"] li{font-size:27px}`,
    0,
  ));
  assert.throws(
    () => validateAndScopeAuthoredCss(
      `.slide [data-pptx-role="body-title"]{font-size:50px}.slide p{font-size:24px}`,
      1,
      { requireDefaultBodyContract: true },
    ),
    /must explicitly size body-title at 50px and body-claim/u,
  );
  assert.throws(
    () => validateAndScopeAuthoredCss(
      `.slide [data-pptx-role="body-title"],.slide .other{font-size:50px}`,
      1,
    ),
    /may not be grouped/u,
  );
});

test("default-template compliance styles are appended after the AI-authored design", () => {
  const html = buildDomDeckHtml({
    title: "既定テンプレート",
    rendererContract: "visual-v7-author-html",
    templateChrome: true,
    defaultTemplateBodyStyle: true,
    theme: { preset: "minimal", secondaryColor: "005B96" },
    slides: [
      {
        kind: "Title",
        title: "表紙",
        authoredHtml: {
          html: `<div class="cover"><h1>表紙</h1></div>`,
          css: `.slide{padding:80px}.slide h1{font-size:54px}`,
        },
      },
      {
        kind: "StructuredBrief",
        title: "NDAは契約であり、結論は条文次第",
        subtitle: "共通論点は目的外利用と第三者開示",
        authoredHtml: {
          html: `<div><h2 data-pptx-role="body-title">NDAは契約であり、結論は条文次第</h2><ul data-pptx-role="body-claim"><li>共通論点は目的外利用と第三者開示</li></ul><p>本文</p></div>`,
          css: `.slide{padding:60px}.slide [data-pptx-role="body-title"]{font-size:50px}.slide [data-pptx-role="body-claim"]{font-size:27px}.slide p{font-size:24px}`,
        },
      },
    ],
  });

  const authoredTitleRule = html.indexOf(`.slide[data-author-slide="2"] [data-pptx-role="body-title"]{font-size:50px}`);
  const complianceTitleRule = html.lastIndexOf(`.slide[data-author-slide="2"] [data-pptx-role="body-title"]{font-size:50px!important;line-height:1.1!important;margin:0!important`);
  const complianceClaimRule = html.lastIndexOf(`.slide[data-author-slide="2"] [data-pptx-role="body-claim"]{font-size:27px!important;margin:8px 0 0!important;padding-left:0!important`);
  assert.ok(authoredTitleRule >= 0);
  assert.ok(complianceTitleRule > authoredTitleRule);
  assert.ok(complianceClaimRule > authoredTitleRule);
});

test("model-authored default cover is limited to a simple title and subtitle overlay", () => {
  const specification = {
    title: "既定表紙",
    rendererContract: "visual-v7-author-html",
    templateChrome: true,
    defaultTemplateCoverOverlay: true,
    theme: { preset: "minimal" },
    slides: [{
      kind: "Title",
      title: "生成AIに秘密情報を入力してよいか",
      subtitle: "NDAの二つの論点から考える",
      authoredHtml: {
        html: `<div class="cover"><h1 data-pptx-role="cover-title">生成AIに秘密情報を入力してよいか</h1><p data-pptx-role="cover-subtitle">NDAの二つの論点から考える</p></div>`,
        css: `.slide .cover{padding:260px 110px;background:#fff}.slide [data-pptx-role="cover-title"]{font-size:54px}.slide [data-pptx-role="cover-subtitle"]{font-size:27px}`,
        assetIds: [],
      },
    }],
  };

  const html = buildDomDeckHtml(specification);
  assert.match(html, /data-pptx-role="cover-title"/u);
  assert.match(html, /data-pptx-role="cover-subtitle"/u);
  assert.match(html, /div,[^\n]+section,[^\n]+article,[^\n]+header,[^\n]+main,[^\n]+aside\{background:transparent!important;border:0!important\}/u);
  assert.match(html, /data-pptx-role="cover-title"\]\{font-size:54px!important/u);
  assert.throws(
    () => buildDomDeckHtml({
      ...specification,
      slides: [{
        ...specification.slides[0],
        authoredHtml: {
          ...specification.slides[0].authoredHtml,
          html: `${specification.slides[0].authoredHtml.html}<p>対象：法務担当者</p>`,
        },
      }],
    }),
    /may contain only its title and subtitle/u,
  );
});

test("default-template role sizes accept unique classes on the protected HTML elements", () => {
  assert.doesNotThrow(() => buildDomDeckHtml({
    title: "既定テンプレート",
    rendererContract: "visual-v7-author-html",
    templateChrome: true,
    defaultTemplateBodyStyle: true,
    theme: { preset: "minimal", secondaryColor: "005B96" },
    slides: [
      {
        kind: "Title",
        title: "表紙",
        authoredHtml: {
          html: `<div class="cover"><h1>表紙</h1><p class="cover-source" data-pptx-role="source-meta">出典</p></div>`,
          css: `.slide .cover{font-size:24px}.slide .cover-source{font-size:20px}`,
          assetIds: [],
        },
      },
      {
        kind: "Bullets",
        title: "NDAは契約であり、結論は条文次第",
        subtitle: "共通論点は目的外利用と第三者開示",
        authoredHtml: {
          html: `<div><h2 class="page-title" data-pptx-role="body-title">NDAは契約であり、結論は条文次第</h2><ul class="page-claim" data-pptx-role="body-claim"><li>共通論点は目的外利用と第三者開示</li></ul></div>`,
          css: `.slide .page-title{font-size:50px}.slide .page-claim li{font-size:27px}`,
          assetIds: [],
        },
      },
    ],
  }));
});

test("protected font exceptions do not follow a class reused by an unrelated element", () => {
  assert.throws(() => buildDomDeckHtml({
    title: "再利用クラス",
    rendererContract: "visual-v7-author-html",
    slides: [
      {
        kind: "Title",
        title: "表紙",
        authoredHtml: {
          html: `<div><p class="small" data-pptx-role="source-meta">出典</p><p class="small">本文</p></div>`,
          css: `.slide .small{font-size:20px}`,
          assetIds: [],
        },
      },
    ],
  }), /at least 24px/u);
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
      templateChrome: true,
      defaultTemplateBodyStyle: true,
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
          subtitle: "意味に合う承認済みアイコンだけを使用する",
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
        {
          kind: "NativeDiagram",
          title: "NDAは「契約」であり結論は条文次第",
          subtitle: "共通して問題になるのは目的外利用と第三者開示の2条項",
          diagram: {
            kind: "Tree",
            direction: "TopToBottom",
            nodes: [
              { id: "premise", label: "前提：NDAは契約", emphasize: true },
              { id: "common", label: "共通の型がある" },
              { id: "purpose", label: "目的外利用の禁止" },
              { id: "disclosure", label: "第三者開示の禁止" },
              { id: "format-a", label: "共通フォーマット第2条3項" },
              { id: "format-b", label: "共通フォーマット第2条4項" },
            ],
            edges: [
              { from: "premise", to: "common" },
              { from: "common", to: "purpose" },
              { from: "common", to: "disclosure" },
              { from: "purpose", to: "format-a" },
              { from: "disclosure", to: "format-b" },
            ],
          },
          speakerNotes: {
            purpose: "DOMで編集可能なツリーを確認する。",
            talkScript: "タイトルと主張を分け、各ノードの間隔を保ったツリーを説明します。",
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
    const slide5 = await archive.file("ppt/slides/slide5.xml")?.async("string") ?? "";
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
    assert.match(slide5, /NDAは「契約」であり結論は条文次第/u);
    assert.match(slide5, /共通して問題になるのは目的外利用と第三者開示の2条項/u);
    assert.match(slide5, /<a:buChar char="•"\/>/u);
    const titleShape = (slide5.match(/<p:sp>.*?<\/p:sp>/gu) ?? [])
      .find((shape) => shape.includes("NDAは「契約」であり結論は条文次第"));
    const claimShape = (slide5.match(/<p:sp>.*?<\/p:sp>/gu) ?? [])
      .find((shape) => shape.includes("共通して問題になるのは目的外利用と第三者開示の2条項"));
    assert.ok(titleShape, "the company body title must remain editable text");
    assert.ok(claimShape, "the company body claim must remain an editable bullet");
    const titleSizes = [...titleShape.matchAll(/\bsz="(\d+)"/gu)].map((match) => Number(match[1]));
    const claimSizes = [...claimShape.matchAll(/\bsz="(\d+)"/gu)].map((match) => Number(match[1]));
    assert.ok(titleSizes.length > 0 && Math.max(...titleSizes) >= 3_000);
    assert.ok(claimSizes.length > 0 && Math.min(...claimSizes) >= 1_600);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});

test("model-authored HTML/CSS is converted without replacing the AI composition", {
  skip: process.env.PPTX_MCP_RUN_DOM_INTEGRATION !== "1",
  timeout: 90_000,
}, async () => {
  const workingDirectory = await mkdtemp(join(tmpdir(), "pptx-mcp-authored-html-"));
  const specificationPath = join(workingDirectory, "visual-deck.json");
  const outputPath = join(workingDirectory, "authored-html.pptx");
  try {
    await writeFile(specificationPath, JSON.stringify({
      title: "AI HTML integration",
      rendererContract: "visual-v7-author-html",
      templateChrome: true,
      defaultTemplateBodyStyle: true,
      theme: { preset: "minimal", secondaryColor: "005B96", accentColor: "0072BC" },
      slides: [
        {
          kind: "NativeDiagram",
          title: "NDAは「契約」であり結論は条文次第",
          authoredHtml: {
            html: `<div class="page"><h1>NDAは「契約」であり結論は条文次第</h1><ul class="claim"><li>共通して問題になるのは目的外利用と第三者開示の2条項</li></ul><div class="diagram"><div class="premise">前提：NDAは契約</div><div class="branch"><div class="node">目的外利用の禁止</div><div class="node">第三者開示の禁止</div></div><div class="icon" data-pptx-icon="decision" aria-label="判断"></div></div></div>`,
            css: `.slide{background:transparent;padding:68px 78px}.slide h1{margin:0;color:var(--secondary);font-size:50px;line-height:1.1}.slide .claim{margin:12px 0 0;padding-left:30px;color:var(--secondary);font-size:27px}.slide .diagram{position:relative;margin:55px auto 0;width:1120px;height:500px}.slide .premise{margin:0 auto;width:360px;padding:24px;background:var(--secondary);color:white;border-radius:10px;text-align:center;font-size:24px;font-weight:700}.slide .branch{display:grid;grid-template-columns:1fr 1fr;gap:120px;margin-top:96px}.slide .node{min-height:130px;padding:30px;border:2px solid var(--secondary);border-radius:10px;background:#EAF3F8;text-align:center;font-size:24px;font-weight:700}.slide .icon{position:absolute;right:8px;bottom:8px;width:56px;height:56px;color:var(--accent)}`,
            assetIds: [],
          },
          speakerNotes: {
            purpose: "NDAの共通論点を分けて示す。",
            talkScript: "NDAは契約であり、結論は条文によって変わります。共通論点を二つに分けて説明します。",
          },
        },
        {
          kind: "Comparison",
          title: "本文タイトル",
          subtitle: "本文を補足する主張",
          authoredHtml: {
            html: `<div class="body"><h2 data-pptx-role="body-title">本文タイトル</h2><ul class="body-claim" data-pptx-role="body-claim"><li>本文を補足する主張</li></ul><div class="content">本文ブロック</div></div>`,
            css: `.slide .body{padding:44px 100px}.slide [data-pptx-role="body-title"]{font-size:50px}.slide .body-claim{font-size:27px;padding-left:30px}.slide .body-claim li{font-size:27px}.slide .content{font-size:24px;margin-top:28px}`,
            assetIds: [],
          },
          speakerNotes: {
            purpose: "本文見出しの変換差を確認する。",
            talkScript: "タイトルと主張の間隔、および箇条書きの上余白を確認します。",
          },
        },
      ],
    }), "utf8");

    const execution = await executeFile(process.execPath, [rendererPath, specificationPath, outputPath], {
      cwd: rendererDirectory,
      timeout: 85_000,
    });
    assert.match(execution.stdout, /PPTX_MCP_RENDERER=dom-to-pptx@2\.1\.1\+react-icons@5\.7\.0/u);

    const generatedHtml = await readFile(join(workingDirectory, "visual-deck.html"), "utf8");
    assert.match(generatedHtml, /class="diagram"/u);
    assert.match(generatedHtml, /grid-template-columns:1fr 1fr/u);
    assert.doesNotMatch(generatedHtml, /class="accent-rail"/u);
    const archive = await JSZip.loadAsync(await readFile(outputPath));
    const slide = await archive.file("ppt/slides/slide1.xml")?.async("string") ?? "";
    const bodySlide = await archive.file("ppt/slides/slide2.xml")?.async("string") ?? "";
    const notes = await archive.file("ppt/notesSlides/notesSlide1.xml")?.async("string") ?? "";
    assert.match(slide, /NDAは「契約」であり結論は条文次第/u);
    assert.match(slide, /目的外利用の禁止/u);
    assert.match(slide, /第三者開示の禁止/u);
    assert.match(slide, /<a:buChar char="•"\/>/u);
    const authoredClaimShape = (slide.match(/<p:sp>.*?<\/p:sp>/gu) ?? [])
      .find((shape) => shape.includes("共通して問題になるのは目的外利用と第三者開示の2条項"));
    assert.ok(authoredClaimShape, "the ordinary authored list must remain editable text");
    assert.match(authoredClaimShape, /<a:bodyPr[^>]*\blIns="[1-9]\d*"[^>]*\btIns="0"/u);
    const claimShape = (bodySlide.match(/<p:sp>.*?<\/p:sp>/gu) ?? [])
      .find((shape) => shape.includes("本文を補足する主張"));
    assert.ok(claimShape, "the protected body claim must remain an editable bullet");
    assert.match(claimShape, /<a:bodyPr[^>]*\btIns="0"/u);
    assert.match(notes, /NDAの共通論点を分けて示す/u);
  } finally {
    await rm(workingDirectory, { recursive: true, force: true });
  }
});
