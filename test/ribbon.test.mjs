/**
 * The ribbon must keep one constant height on every tab: 1/16 of the viewport.
 * Groups stay on a single row and scroll horizontally instead of wrapping, so
 * switching from File to Home never moves the document.
 */
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import path from "node:path";
import test from "node:test";

const here = path.dirname(fileURLToPath(import.meta.url));
const clientPath = path.join(here, "..", "client.js");
const fixturePath = path.join(here, "fixture.json");

function loadReact() {
  for (const root of [
    new URL("../../apps/web/package.json", import.meta.url),
    new URL("../../package.json", import.meta.url),
  ]) {
    try {
      const requireFromRoot = createRequire(root);
      return {
        React: requireFromRoot("react"),
        server: requireFromRoot("react-dom/server"),
      };
    } catch {
      // Try the next resolution root.
    }
  }
  throw new Error("react is not installed in this workspace");
}

const { React, server } = loadReact();
let captured = null;
globalThis.window = {
  __ModuleLoader__: {
    load(definition) {
      captured = definition;
    },
  },
};
new Function("window", readFileSync(clientPath, "utf8"))(globalThis.window);
assert.ok(captured, "the loader captures the module definition");
const bundled = captured.factory((id) => {
  if (id === "react") return React;
  throw new Error(`client.js may only require react (asked for ${id})`);
});
const model = JSON.parse(readFileSync(fixturePath, "utf8"));
const pageMap = Object.fromEntries(
  model.sections
    .flatMap((section) => section.blocks)
    .map((block, index) => [block.id, index < 3 ? 1 : 2]),
);

function render(options) {
  return server.renderToStaticMarkup(
    __test.renderForTest(
      model,
      Object.assign(
        { blockPages: pageMap, pageCount: 2, tab: "home" },
        options || {},
      ),
    ),
  );
}
const __test = bundled.__test;

test("ribbon keeps the agreed 1/16 compromise so nothing clips when short", () => {
  const markup = render();
  // 1/16 of the viewport on a normal window, never shorter than 56px: a 50px
  // ribbon (805px viewport) clipped captions, labels mode and the gallery.
  assert.match(
    markup,
    /\.dx-ribbon\{[^}]*height:max\(56px,calc\(100vh \/ 16\)\)/,
  );
  assert.match(markup, /\.dx-ribbon\{[^}]*min-height:56px/);
  assert.match(markup, /\.dx-ribbon\{[^}]*overflow-x:auto/);
  assert.match(markup, /\.dx-ribbon\{[^}]*overflow-y:hidden/);
  // The old 120px cap and 36px floor are gone; they are what moved the height
  // away from the constant rule.
  assert.doesNotMatch(markup, /\.dx-ribbon\{[^}]*max-height:120px/);
  assert.doesNotMatch(markup, /\.dx-ribbon\{[^}]*min-height:36px/);
  // The height must not be driven by the active tab's content.
  assert.doesNotMatch(markup, /\.dx-ribbon\{[^}]*min-height:7[0-9]px/);
});

test("captions and group contents cannot clip inside the ribbon", () => {
  const markup = render();
  // A caption is one 9px line box: no vertical ellipsis clipping.
  assert.match(markup, /\.dx-group-label\{[^}]*line-height:1[;}]/);
  assert.match(markup, /\.dx-group-label\{[^}]*white-space:nowrap/);
  // Group items are centred, so a 30px button never hugs the caption.
  assert.match(markup, /\.dx-group-items\{[^}]*align-items:center/);
  // Labels mode and the style gallery are bounded by the available height.
  assert.match(
    markup,
    /\.dx-show-labels \.dx-tool\{[^}]*height:calc\(100% - 2px\)/,
  );
  assert.match(
    markup,
    /\.dx-show-labels \.dx-style-card\{[^}]*max-height:100%/,
  );
  assert.match(markup, /\.dx-style-card\{[^}]*max-height:100%/);
  assert.match(markup, /\.dx-gallery\{[^}]*max-height:100%/);
  assert.match(markup, /\.dx-stepper\{[^}]*max-height:100%/);
  assert.match(markup, /\.dx-select\{[^}]*max-height:100%/);
});

test("the File compact row and the backstage header keep their own height", () => {
  const markup = render({ tab: "file" });
  assert.match(markup, /\.dx-file-compact\{[^}]*height:100%/);
  assert.match(markup, /\.dx-file-compact\{[^}]*overflow-y:hidden/);
  assert.match(markup, /\.dx-file-compact \.dx-tool\{[^}]*width:30px/);
  assert.match(markup, /\.dx-backstage-head\{[^}]*flex:0 0 auto/);
  assert.match(markup, /\.dx-backstage-head\{[^}]*min-height:52px/);
});

test("the palette follows the real DSH theme tokens in both modes", () => {
  const markup = render();
  for (const token of [
    "--dsw-alias-label-primary",
    "--dsw-alias-label-secondary",
    "--dsw-alias-border-l1",
    "--dsw-alias-bg-layer-1",
    "--dsw-alias-bg-base",
    "--dsw-alias-brand-primary",
    "--dsw-alias-state-error-primary",
    "--dsw-alias-state-success-primary",
    "--dsw-alias-state-warn-primary",
    "--dsw-specific-sidebar-fill",
  ])
    assert.ok(markup.includes(token), `the stylesheet must use ${token}`);
  // The legacy --dsh-* names stay as fallbacks inside the token chain.
  assert.match(
    markup,
    /--dx-text:var\(--dsw-alias-label-primary,var\(--dsh-text,/,
  );
  assert.match(
    markup,
    /--dx-muted:var\(--dsw-alias-label-secondary,var\(--dsh-muted,/,
  );
  assert.match(
    markup,
    /--dx-border:var\(--dsw-alias-border-l1,var\(--dsh-border,/,
  );
  assert.match(markup, /--dx-surface:var\(--dsw-alias-bg-layer-1,/);
  assert.match(markup, /--dx-bg:var\(--dsw-alias-bg-base,/);
  assert.match(
    markup,
    /--dx-accent:var\(--dsw-alias-brand-primary,var\(--dsh-accent,/,
  );
  assert.match(markup, /--dx-danger:var\(--dsw-alias-state-error-primary,/);
  assert.match(markup, /--dx-success:var\(--dsw-alias-state-success-primary,/);
  assert.match(markup, /--dx-warn:var\(--dsw-alias-state-warn-primary,/);
  // Text on the accent bar comes from the inverted label token, so the bar is
  // readable whichever way the brand colour points.
  assert.match(
    markup,
    /--dx-on-accent:var\(--dsw-alias-label-primary-inverted,/,
  );
  // The colour tokens are NOT redefined by prefers-color-scheme: the tokens
  // already switch. Only the paper sheet keeps a media query.
  const media = markup.slice(
    markup.indexOf("@media (prefers-color-scheme:dark)"),
    markup.indexOf("@media (prefers-color-scheme:dark)") + 260,
  );
  assert.ok(media.includes("--dx-paper-border"));
  for (const token of ["--dx-text", "--dx-muted", "--dx-accent", "--dx-bg"])
    assert.ok(
      !media.includes(`${token}:`),
      `${token} must not be redefined by prefers-color-scheme`,
    );
});

test("tool icons are 20px with the stronger 1.8 stroke and currentColor", () => {
  const markup = render({ tab: "home" });
  assert.match(
    markup,
    /aria-label="Undo"[^>]*><svg[^>]*width="20"[^>]*height="20"[^>]*stroke="currentColor"[^>]*stroke-width="1\.8"/,
  );
  // No 15/16px tool icons and no 1.7 stroke remain in the ribbon.
  assert.doesNotMatch(markup, /aria-label="Undo"[^>]*><svg[^>]*width="1[56]"/);
  assert.doesNotMatch(
    markup,
    /aria-label="Undo"[^>]*><svg[^>]*stroke-width="1\.7"/,
  );
  // The glyph buttons grow with the pictograms.
  assert.match(markup, /\.dx-glyph\{[^}]*width:20px;height:20px/);
});

test("tool buttons are 30px and no fixed 24px tool size remains", () => {
  const markup = render();
  assert.match(markup, /\.dx-tool\{[^}]*width:30px;height:30px/);
  assert.doesNotMatch(markup, /\.dx-tool\{[^}]*width:24px/);
  assert.doesNotMatch(markup, /\.dx-tool\{[^}]*height:24px/);
  assert.doesNotMatch(markup, /\.dx-tool-small\{[^}]*24px/);
  assert.doesNotMatch(markup, /\.dx-tool-tiny\{[^}]*24px/);
  assert.doesNotMatch(markup, /\.dx-file-compact \.dx-tool\{[^}]*28px/);
  // Labels mode gives the compact variants more room, not less.
  assert.match(
    markup,
    /\.dx-show-labels \.dx-tool\.dx-tool-small[\s\S]{0,160}?min-width:32px/,
  );
  // Disabled tools stay readable.
  assert.match(markup, /\.dx-tool\[disabled\]\{opacity:\.55/);
  // Hover/active/pressed use the brand colour and a visible fill.
  assert.match(
    markup,
    /\.dx-tool:hover:not\(\[disabled\]\)\{[^}]*color:var\(--dx-accent\)/,
  );
  assert.match(
    markup,
    /\.dx-tool\[aria-pressed=true\]\{[^}]*color:var\(--dx-accent\)/,
  );
});

test("ribbon groups stay on one row so tabs cannot change the height", () => {
  const markup = render();
  assert.match(markup, /\.dx-group-items\{[^}]*flex-wrap:nowrap/);
  assert.match(markup, /\.dx-group-items\{[^}]*overflow-x:auto/);
  assert.match(markup, /\.dx-group\{[^}]*height:100%/);
  // No per-tab minimum heights remain.
  assert.doesNotMatch(markup, /\.dx-group\{[^}]*min-height:66px/);
});

test("every tab renders inside the same ribbon container", () => {
  for (const tab of [
    "file",
    "home",
    "insert",
    "design",
    "layout",
    "references",
    "view",
  ]) {
    const markup = render({ tab });
    assert.equal(
      (markup.match(/class="dx-ribbon"/g) || []).length,
      1,
      `${tab} keeps exactly one ribbon element`,
    );
    assert.match(markup, /class="dx-ribbon"/);
  }
});

test("the File tab keeps the same ribbon height and opens its details as an overlay", () => {
  const markup = render({ tab: "file" });
  assert.equal((markup.match(/class="dx-ribbon"/g) || []).length, 1);
  assert.match(markup, /class="dx-file-compact"/);
  assert.match(markup, /class="dx-backstage"/);
  assert.match(markup, /role="dialog"/);
  assert.match(markup, /Open documents/);
  assert.match(markup, /Choose folder…/);
  // The overlay is positioned absolutely, so it cannot add layout height.
  assert.match(markup, /\.dx-backstage\{[^}]*position:absolute/);
});

test("dialogs from the File backstage stack above it", () => {
  const markup = render({ tab: "file" });
  const overlay = markup.match(/\.dx-overlay\{[^}]*z-index:(\d+)/);
  const backstage = markup.match(/\.dx-backstage\{[^}]*z-index:(\d+)/);
  assert.ok(overlay, "the dialog overlay has a z-index");
  assert.ok(backstage, "the backstage has a z-index");
  assert.ok(
    Number(overlay[1]) > Number(backstage[1]),
    `dialog overlay (${overlay[1]}) must stack above the backstage (${backstage[1]})`,
  );
});

test("Escape closes the File backstage even when focus is outside it", () => {
  const source = readFileSync(clientPath, "utf8");
  // A window-level capture listener is required: a handler on the overlay div
  // never sees keydown while focus is on the File tab button.
  assert.match(
    source,
    /window\.addEventListener\("keydown",\s*onKeyDown,\s*true\)/,
  );
  assert.match(source, /getElementById\("dx-tab-file"\)/);
  assert.match(source, /closeBackstage/);
  // Dialogs own Escape first.
  assert.match(source, /querySelector\("\.dx-dialog"\)/);
});

test("icon-only mode omits labels but keeps tooltips, and the toggle restores them", () => {
  const compact = render({ showLabels: false });
  assert.ok(
    !compact.includes('class="dx-tool-label"'),
    "icon-only mode has no label spans",
  );
  assert.match(compact, /title="Bold \(Ctrl\+B\)[^"]*"/);
  assert.match(compact, /aria-label="Bold"/);

  const labelled = render({ showLabels: true });
  assert.ok(
    labelled.includes('class="dx-tool-label"'),
    "labels mode renders label spans",
  );
  assert.match(
    labelled,
    /\.dx-show-labels \.dx-tool\{[^}]*height:calc\(100% - 2px\)/,
  );
});

/**
 * A hard line break (DOCX `<w:br/>`, Word's Shift+Enter) used to render as an
 * empty inline atom with a permanent ↵ marker, so the paragraph stayed on one
 * line. It must be a real `br` inside the single-character atom wrapper, with
 * the marker drawn only under View → Formatting marks.
 */
function modelWithLineBreak() {
  const withBreak = JSON.parse(JSON.stringify(model));
  const paragraph = withBreak.sections
    .flatMap((section) => section.blocks)
    .find((block) => block.id === "p2");
  paragraph.inlines.splice(2, 0, { type: "break", kind: "line", props: {} });
  return withBreak;
}

function renderModel(renderedModel, options) {
  return server.renderToStaticMarkup(
    __test.renderForTest(
      renderedModel,
      Object.assign(
        { blockPages: pageMap, pageCount: 2, tab: "home" },
        options || {},
      ),
    ),
  );
}

test("a line break renders a real br inside one atom character", () => {
  const markup = renderModel(modelWithLineBreak());
  assert.match(
    markup,
    /class="dx-atom dx-atom-break"[^>]*data-dx-atom="\n"[^>]*data-dx-inline="2"[^>]*><br\/>/,
  );
  // The break stays in the inline flow; an inline-block box would swallow it.
  assert.match(markup, /\.dx-atom-break\{display:inline;/);
  // Exactly one character in the offset space, so caret mapping is unchanged.
  assert.equal((markup.match(/data-dx-atom="\n"/g) || []).length, 1);
  // The ↵ marker is CSS-only and only inside the formatting-marks scope.
  assert.ok(
    markup.includes('.dx-marks .dx-atom-break::before{content:"\\21B5"'),
  );
  assert.ok(
    !markup.includes(".dx-atom-break::after"),
    "the always-on marker must be gone",
  );
});

test("the line-break marker appears only with formatting marks on", () => {
  const off = renderModel(modelWithLineBreak());
  assert.ok(!off.includes('class="dx-editor dx-marks"'));
  assert.ok(!off.includes('class="dx-editor dx-marks"'));
  const on = renderModel(modelWithLineBreak(), { showMarks: true });
  assert.match(on, /class="dx-editor dx-marks"/);
  assert.ok(
    on.includes('.dx-marks .dx-atom-break::before{content:"\\21B5"'),
    "the marker rule is in the stylesheet and scoped to .dx-marks",
  );
});

test("Shift+Enter inserts a line break and keeps plain Enter splitting", () => {
  const source = readFileSync(clientPath, "utf8");
  assert.match(
    source,
    /if \(event\.key === "Enter"\) \{\s*event\.preventDefault\(\);\s*\/\/[^\n]*\n\s*if \(event\.shiftKey\) void insertLineBreakAtCaret\(\);\s*else void splitParagraph\(\);/,
  );
  assert.match(source, /const insertLineBreakAtCaret = React\.useCallback/);
  // The Host op and the caret that sits immediately after the new break.
  assert.match(
    source,
    /"insertPageBreak",\s*\{ position: \{ blockId, offset \}, kind: "line" \},\s*\{ reload: true, caret: target \}/,
  );
  assert.match(
    source,
    /offset: offset \+ 1[\s\S]{0,80}?end: \{ blockId, offset: offset \+ 1 \}/,
  );
});

test("opening by an advanced path reveals the document and prepares the session", () => {
  const source = readFileSync(clientPath, "utf8");
  // The advanced-path action must use the same flow as the folder browser:
  // /open-path, workspace switch, session preparation, then the Home tab.
  assert.match(source, /const openDocumentByPath = React\.useCallback/);
  assert.match(source, /runOpenDocumentPathFlow\(trimmed,/);
  assert.match(source, /advancedOpen: false,\s*tab: "home"/);
  // It must not fall back to the plain legacy open that left the backstage up.
  assert.doesNotMatch(
    source,
    /const result = await apiClient\.open\(trimmed\)/,
  );
});
