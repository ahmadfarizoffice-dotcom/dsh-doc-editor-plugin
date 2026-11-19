/**
 * The ribbon is 1/10 of the viewport by default and user-resizable: dragging
 * the bottom separator (or the keyboard) picks any height between 64px and
 * half the viewport, and the choice is remembered for this browser. Groups
 * stay on a single row and scroll horizontally instead of wrapping, so
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
const source = readFileSync(clientPath, "utf8");

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
new Function("window", source)(globalThis.window);
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

test("the ribbon defaults to 1/10 of the viewport and is user-resizable", () => {
  const markup = render();
  // Default: 1/10 of the viewport, clamped to a band that always fits content.
  assert.match(
    markup,
    /\.dx-ribbon\{[^}]*height:clamp\(72px,calc\(100vh \/ 10\),240px\)/,
  );
  // The old 1/16 rule is gone, and so is the 56px floor it needed.
  assert.equal(source.includes("100vh / 16"), false);
  assert.doesNotMatch(markup, /max\(56px,calc\(100vh/);
  assert.match(markup, /\.dx-ribbon\{[^}]*min-height:64px/);
  assert.match(markup, /\.dx-ribbon\{[^}]*max-height:50vh/);
  assert.match(markup, /\.dx-ribbon\{[^}]*overflow-x:auto/);
  assert.match(markup, /\.dx-ribbon\{[^}]*overflow-y:hidden/);
  // The bottom 6px strip is reserved for the handle so nothing hides under it.
  assert.match(markup, /\.dx-ribbon\{[^}]*padding:0 6px 6px/);
  // The old caps that moved the height away from the rule are gone.
  assert.doesNotMatch(markup, /\.dx-ribbon\{[^}]*max-height:120px/);
  assert.doesNotMatch(markup, /\.dx-ribbon\{[^}]*min-height:36px/);
  // The height must not be driven by the active tab's content.
  assert.doesNotMatch(markup, /\.dx-ribbon\{[^}]*min-height:7[0-9]px/);
});

test("the ribbon exposes a keyboard-accessible drag separator", () => {
  const markup = render();
  assert.match(markup, /class="dx-ribbon-handle"/);
  assert.match(markup, /class="dx-ribbon-grip"/);
  assert.match(markup, /role="separator"/);
  assert.match(markup, /aria-orientation="horizontal"/);
  assert.match(markup, /aria-label="Ribbon height"/);
  assert.match(markup, /aria-valuemin="64"/);
  assert.match(markup, /aria-valuenow="\d+"/);
  assert.match(markup, /aria-valuemax="\d+"/);
  assert.match(markup, /aria-valuetext="\d+ pixels"/);
  assert.match(
    markup,
    /title="Drag to resize the ribbon; double-click to reset"/,
  );
  // The strip is focusable and row-resize, 6px tall.
  assert.match(markup, /\.dx-ribbon-handle\{[^}]*cursor:row-resize/);
  assert.match(markup, /\.dx-ribbon-handle\{[^}]*height:6px/);
  assert.match(
    markup,
    /\.dx-ribbon-handle:focus-visible\{[^}]*outline:2px solid/,
  );
  // A double click and Enter/Space reset to the default.
  const handleSource = source.slice(
    source.indexOf("function RibbonResizer"),
    source.indexOf("function ribbonStyle"),
  );
  assert.match(handleSource, /onDoubleClick: \(\) => apply\(null, true\)/);
  assert.match(handleSource, /event\.key === "Enter"/);
  assert.match(handleSource, /event\.key === "ArrowUp"/);
  assert.match(handleSource, /event\.key === "ArrowDown"/);
  assert.match(handleSource, /event\.key === "Home"/);
  assert.match(handleSource, /event\.key === "End"/);
  assert.match(handleSource, /RIBBON_KEYBOARD_STEP/);
  assert.match(handleSource, /setPointerCapture/);
  assert.match(handleSource, /releasePointerCapture/);
  // The File backstage follows the measured ribbon bottom, not a constant.
  assert.match(
    markup,
    /\.dx-backstage\{[^}]*top:var\(--dx-ribbon-bottom,72px\)/,
  );
  assert.equal(markup.includes("top:72px;bottom:0"), false);
  assert.match(source, /new ResizeObserver\(measure\)/);
  assert.match(source, /"--dx-ribbon-bottom"/);
});

test("ribbon height helpers clamp and default to the 1/10 rule", () => {
  const { defaultRibbonHeight, clampRibbonHeight, ribbonMaxHeight } = __test;
  // 900px viewport -> 90px, 1080 -> 108, tiny -> the 72px floor, huge -> 240.
  assert.equal(defaultRibbonHeight(900), 90);
  assert.equal(defaultRibbonHeight(1080), 108);
  assert.equal(defaultRibbonHeight(500), 72);
  assert.equal(defaultRibbonHeight(4000), 240);
  assert.equal(defaultRibbonHeight(0), 72);
  // Dragging clamps to 64px … 50vh.
  assert.equal(ribbonMaxHeight(900), 450);
  assert.equal(ribbonMaxHeight(0), 240);
  assert.equal(clampRibbonHeight(40, 900), 64);
  assert.equal(clampRibbonHeight(900, 900), 450);
  assert.equal(clampRibbonHeight(140, 900), 140);
  assert.equal(clampRibbonHeight(139.6, 900), 140);
  assert.equal(clampRibbonHeight(5000, 0), 240);
  assert.equal(clampRibbonHeight(null, 900), null);
  assert.equal(clampRibbonHeight(Number.NaN, 900), null);
});

test("the chosen ribbon height round-trips through localStorage", () => {
  const storage = new Map();
  globalThis.localStorage = {
    getItem: (key) => (storage.has(key) ? storage.get(key) : null),
    setItem: (key, value) => storage.set(key, value),
    removeItem: (key) => storage.delete(key),
  };
  try {
    assert.equal(__test.RIBBON_HEIGHT_KEY, "dsh-doc:ribbon-height");
    assert.equal(__test.readStoredRibbonHeight(), null, "unset by default");
    __test.storeRibbonHeight(140.4);
    assert.equal(storage.get("dsh-doc:ribbon-height"), "140");
    assert.equal(__test.readStoredRibbonHeight(), 140);
    // A reset removes the key, so the CSS 1/10 default applies again.
    __test.storeRibbonHeight(null);
    assert.equal(storage.has("dsh-doc:ribbon-height"), false);
    // A corrupt value is ignored instead of throwing.
    storage.set("dsh-doc:ribbon-height", "tall");
    assert.equal(__test.readStoredRibbonHeight(), null);
    // A fresh panel state restores the stored height on mount.
    storage.set("dsh-doc:ribbon-height", "220");
    assert.equal(__test.initialState().ribbonHeight, 220);
    const markup = render({ ribbonHeight: 220 });
    assert.match(
      markup,
      /class="dx-ribbon" id="dx-ribbon"[^>]*style="height:220px"/,
    );
  } finally {
    delete globalThis.localStorage;
  }
});

test("the chooser remembers the folder that was actually opened", () => {
  const storage = new Map();
  globalThis.localStorage = {
    getItem: (key) => (storage.has(key) ? storage.get(key) : null),
    setItem: (key, value) => storage.set(key, value),
    removeItem: (key) => storage.delete(key),
  };
  try {
    assert.equal(__test.LAST_FOLDER_KEY, "dsh-doc:last-folder");
    assert.equal(__test.readStoredFolder(), "");
    __test.storeFolder("C:\\Users\\me\\Documents\\Reports");
    assert.equal(
      __test.readStoredFolder(),
      "C:\\Users\\me\\Documents\\Reports",
    );
    assert.equal(
      __test.folderPathOf("C:\\Users\\me\\Documents\\Reports\\Q3.docx"),
      "C:\\Users\\me\\Documents\\Reports",
    );
    assert.equal(
      __test.folderPathOf("/home/me/Documents/Q3.docx"),
      "/home/me/Documents",
    );
    assert.equal(__test.folderPathOf("Q3.docx"), "");
  } finally {
    delete globalThis.localStorage;
  }
});

test("captions and group contents cannot clip inside the ribbon", () => {
  const markup = render();
  // A caption is one 9px line whose line box holds the full glyph height
  // (line-height 1.2), so descenders are never cut by the 64px floor.
  assert.match(markup, /\.dx-group-label\{[^}]*line-height:1\.2/);
  assert.match(markup, /\.dx-group-label\{[^}]*white-space:nowrap/);
  // Labels mode uses the same roomy line box for its per-tool captions.
  assert.match(
    markup,
    /\.dx-show-labels \.dx-tool \.dx-tool-label\{[^}]*line-height:1\.2/,
  );
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
  assert.match(markup, /\.dx-file-compact \.dx-tool\{[^}]*width:32px/);
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

test("tool icons are 22px with the 1.9 stroke and currentColor", () => {
  const markup = render({ tab: "home" });
  assert.match(
    markup,
    /aria-label="Undo"[^>]*><svg[^>]*width="22"[^>]*height="22"[^>]*stroke="currentColor"[^>]*stroke-width="1\.9"/,
  );
  // No 15/16/20px tool icons and no 1.7/1.8 stroke remain on the ribbon tools.
  assert.doesNotMatch(markup, /aria-label="Undo"[^>]*><svg[^>]*width="1[56]"/);
  assert.doesNotMatch(markup, /aria-label="Undo"[^>]*><svg[^>]*width="20"/);
  assert.doesNotMatch(
    markup,
    /aria-label="Undo"[^>]*><svg[^>]*stroke-width="1\.[78]"/,
  );
  // Every icon stays inside the 24-unit grid with round caps and joins.
  assert.match(
    markup,
    /aria-label="Undo"[^>]*><svg[^>]*viewBox="0 0 24 24"[^>]*stroke-linecap="round"[^>]*stroke-linejoin="round"/,
  );
});

function buttonMarkup(markup, label) {
  const needle = `aria-label="${label}"`;
  const at = markup.indexOf(needle);
  assert.notEqual(at, -1, `the ribbon should render the "${label}" tool`);
  const start = markup.lastIndexOf("<button", at);
  const end = markup.indexOf("</button>", at);
  return markup.slice(start, end + "</button>".length);
}

test("the Font group draws B/I/U/S/sub/sup as SVG paths, not text glyphs", () => {
  const markup = render({ tab: "home" });
  const labels = [
    "Bold",
    "Italic",
    "Underline",
    "Strikethrough",
    "Subscript",
    "Superscript",
  ];
  for (const label of labels) {
    const button = buttonMarkup(markup, label);
    // A real pictogram: 22px SVG with at least one path, no letter text node.
    assert.match(button, /<svg[^>]*width="22"[^>]*height="22"/, label);
    assert.match(button, /stroke-width="1\.9"/, label);
    assert.match(button, /<path d="M/, label);
    assert.doesNotMatch(button, /dx-glyph/, label);
    assert.doesNotMatch(button, />\s*[BIUS]\s*</, label);
    assert.doesNotMatch(button, />x[²₂]</, label);
  }
  // The glyph component and its stylesheet rules are gone entirely.
  assert.equal(source.includes("dx-glyph"), false);
  assert.equal(source.includes("glyph:"), false);
  // Labels mode still keeps its captions and tooltips.
  const labelled = render({ tab: "home", showLabels: true });
  assert.ok(labelled.includes(">Bold<"));
  assert.match(buttonMarkup(labelled, "Bold"), /title="Bold \(Ctrl\+B\)/);
});

test("tool buttons are 32px and grow to a 34px minimum with labels", () => {
  const markup = render();
  assert.match(markup, /\.dx-tool\{[^}]*width:32px;height:32px/);
  assert.doesNotMatch(markup, /\.dx-tool\{[^}]*width:30px/);
  assert.doesNotMatch(markup, /\.dx-tool\{[^}]*height:30px/);
  assert.doesNotMatch(markup, /\.dx-tool\{[^}]*width:24px/);
  assert.doesNotMatch(markup, /\.dx-tool\{[^}]*height:24px/);
  assert.doesNotMatch(markup, /\.dx-tool-small\{[^}]*30px/);
  assert.doesNotMatch(markup, /\.dx-tool-tiny\{[^}]*30px/);
  assert.doesNotMatch(markup, /\.dx-file-compact \.dx-tool\{[^}]*30px/);
  // Labels mode gives the compact variants more room, not less.
  assert.match(
    markup,
    /\.dx-show-labels \.dx-tool\.dx-tool-small[\s\S]{0,160}?min-width:34px/,
  );
  assert.match(markup, /\.dx-show-labels \.dx-tool\{[^}]*min-width:34px/);
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

test("Ctrl+O opens the in-panel file chooser from the panel root", () => {
  // The root keydown handler owns the shortcut, not a window listener: the
  // panel only takes Ctrl+O while it is mounted and focused.
  assert.match(source, /const handleKeyDown = React\.useCallback/);
  const at = source.indexOf('event.key === "o" || event.key === "O"');
  assert.notEqual(at, -1, "Ctrl+O is handled");
  const block = source.slice(at - 80, at + 500);
  assert.match(block, /!typing/);
  assert.match(block, /event\.preventDefault\(\)/);
  assert.match(block, /openFileChooserRef\.current\(\)/);
  // A text field or the contenteditable document owns the keystroke, and an
  // open dialog returns before the shortcut is considered.
  assert.match(
    source,
    /const typing =\s*target &&[\s\S]{0,320}?contenteditable=true/,
  );
  assert.match(source, /if \(inDialog\) \{/);
  // The chooser opens the existing /browse dialog at the remembered folder.
  assert.match(source, /const openFileChooser = React\.useCallback/);
  const chooser = source.slice(
    source.indexOf("const openFileChooser = React.useCallback"),
    source.indexOf("openFileChooserRef.current = openFileChooser"),
  );
  assert.match(chooser, /readStoredFolder\(\)/);
  assert.match(chooser, /openBrowseDialog\(target\)/);
  // Every entry point calls the same action.
  assert.match(source, /case "openFileChooser":\s*await openFileChooser\(\);/);
  assert.match(source, /h\("input", \{\s*key: "file-input"/);
  // The OS file dialog is a secondary option: if click() is unavailable the
  // panel opens the chooser instead of dead-ending on a notice.
  const openDocument = source.slice(
    source.indexOf('case "openDocument":'),
    source.indexOf('case "toggleAdvanced":'),
  );
  assert.match(openDocument, /fileInputRef\.current\.click\(\)/);
  assert.match(openDocument, /await openFileChooser\(\)/);
  assert.equal(
    source.includes("This Host cannot open the native file dialog; use"),
    false,
  );
  // Opening a document from the chooser still goes through /open-path and the
  // shared finish flow (Home tab + workspace/chat chip).
  const openFromBrowse = source.slice(
    source.indexOf("const openFromBrowse = React.useCallback"),
    source.indexOf("const openCurrentFolder = React.useCallback"),
  );
  assert.match(openFromBrowse, /runOpenDocumentPathFlow\(path/);
  assert.match(openFromBrowse, /storeFolder\(folder\)/);
  assert.match(openFromBrowse, /finishOpenedDocument\(result\)/);
});

test("the File tab offers Open file… next to the device and folder pickers", () => {
  const markup = render({ tab: "file" });
  assert.ok(markup.includes(">Open file…<"));
  assert.ok(markup.includes(">Open from this device…<"));
  assert.ok(markup.includes(">Choose folder…<"));
  // The compact ribbon row and the backstage Document column both expose it.
  const compactRow = markup.slice(
    markup.indexOf('class="dx-file-compact"'),
    markup.indexOf("</div>", markup.indexOf('class="dx-file-compact"')),
  );
  assert.ok(compactRow.includes('aria-label="Open file…"'));
  const backstage = markup.slice(markup.indexOf('class="dx-backstage"'));
  assert.ok(backstage.includes('aria-label="Open file…"'));
  // The empty state's open affordance reaches the same chooser.
  const empty = render({ noDocument: true });
  assert.ok(empty.includes(">Open file…<"));
  assert.ok(empty.includes(">Open by path<"));
});

test("keyboard resize follows the drag direction (down grows, up shrinks)", () => {
  const source = readFileSync(clientPath, "utf8");
  assert.match(
    source,
    /event\.key === "ArrowDown"\) next = value \+ RIBBON_KEYBOARD_STEP/,
    "ArrowDown grows the ribbon like dragging the bottom edge down",
  );
  assert.match(
    source,
    /event\.key === "ArrowUp"\) next = value - RIBBON_KEYBOARD_STEP/,
    "ArrowUp shrinks the ribbon",
  );
});
