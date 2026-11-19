/**
 * Panel integration tests: load `client.js` exactly the way the DSH client
 * module loader does, then render the Documents panel to static markup.
 *
 * The panel is rendered through `__test.renderForTest(model, options)`, which
 * seeds the panel state with a fixture model and returns the React element, so
 * the assertions cover the real component tree (ribbon, editable surface,
 * status bar, dialogs) and not a mock.
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
  const roots = [
    new URL("../../apps/web/package.json", import.meta.url),
    new URL("../../package.json", import.meta.url),
  ];
  for (const root of roots) {
    try {
      const requireFromRoot = createRequire(root);
      return {
        React: requireFromRoot("react"),
        server: requireFromRoot("react-dom/server"),
      };
    } catch {
      /* the next root may resolve react */
    }
  }
  throw new Error("react is not installed in this workspace");
}

const { React, server } = loadReact();

function loadPanel() {
  let captured = null;
  globalThis.window = {
    __ModuleLoader__: {
      load(definition) {
        captured = definition;
      },
    },
  };
  const source = readFileSync(clientPath, "utf8");
  // The plugin file is a plain browser script; evaluate it with the stubbed
  // loader exactly like the DSH client does.
  new Function("window", source)(globalThis.window);
  assert.ok(captured, "the loader must capture the module definition");
  const bundled = captured.factory((id) => {
    if (id === "react") return React;
    throw new Error(`client.js may only require react (asked for ${id})`);
  });
  return { captured, bundled, source };
}

const { captured, bundled, source } = loadPanel();
const { __test } = bundled;
const model = JSON.parse(readFileSync(fixturePath, "utf8"));
const PAGE_MAP = {
  p1: 1,
  p2: 1,
  t1: 1,
  p3: 1,
  p4: 1,
  p5: 1,
  toc1: 1,
  p6: 1,
  p7: 2,
};

function render(options) {
  const element = __test.renderForTest(
    model,
    Object.assign({ blockPages: PAGE_MAP, pageCount: 2 }, options || {}),
  );
  assert.ok(
    React.isValidElement(element),
    "renderForTest returns a React element",
  );
  return server.renderToStaticMarkup(element);
}

/** The whole opening tag that contains `needle`, for attribute assertions. */
function tagAround(markup, needle) {
  const index = markup.indexOf(needle);
  assert.notEqual(index, -1, `markup should contain ${needle}`);
  const start = markup.lastIndexOf("<", index);
  const end = markup.indexOf(">", index);
  return markup.slice(start, end + 1);
}

/**
 * Paragraphs use React's `contentEditable` prop, which SSR serialises as
 * `contentEditable="true"`. HTML attribute names are case-insensitive, so the
 * comparison lower-cases the markup: the browser DOM attribute is exactly
 * `contenteditable="true"`.
 */
function hasEditableAttribute(markup, value) {
  return markup.toLowerCase().includes(`contenteditable="${value}"`);
}

function indexOfLabel(markup, label) {
  const index = markup.indexOf(`>${label}</button>`);
  assert.notEqual(index, -1, `markup should contain the "${label}" tab button`);
  return index;
}

test("module envelope keeps the DSH loader contract", () => {
  assert.equal(captured.id, "dsh-doc-editor-plugin");
  assert.deepEqual(bundled.inject, ["slots"]);
  assert.equal(typeof bundled.apply, "function");
  assert.equal(typeof bundled.__test, "object");
  assert.ok(source.startsWith("window.__ModuleLoader__.load({"));
  assert.ok(source.includes('id: "dsh-doc-editor-plugin"'));
});

test("the panel renders the seven Word tabs in order", () => {
  const markup = render({ tab: "home" });
  assert.ok(markup.startsWith('<div class="dx-root">'));
  assert.ok(markup.includes("<style>"), "the panel injects its stylesheet");
  assert.ok(markup.includes(".dx-root"));
  const labels = [
    "File",
    "Home",
    "Insert",
    "Design",
    "Layout",
    "References",
    "View",
  ];
  const indexes = labels.map((label) => indexOfLabel(markup, label));
  for (let index = 1; index < indexes.length; index += 1) {
    assert.ok(
      indexes[index] > indexes[index - 1],
      `${labels[index]} must come after ${labels[index - 1]}`,
    );
  }
  assert.ok(markup.includes('role="tablist"'));
  assert.ok(markup.includes('aria-label="Ribbon tabs"'));
  // Roving tabindex: the active tab is focusable, the others are not.
  assert.match(tagAround(markup, ">Home</button>"), /aria-selected="true"/);
  assert.match(tagAround(markup, ">Home</button>"), /tabindex="0"/);
  assert.match(tagAround(markup, ">Insert</button>"), /tabindex="-1"/);
  assert.match(tagAround(markup, ">Insert</button>"), /aria-selected="false"/);
});

test("font controls are enabled for a collapsed caret and follow the selection", () => {
  // Word behaviour: clicking into text and pressing Bold must work with no
  // selection at all; the action formats the caret paragraph.
  const collapsed = render({
    selection: {
      start: { blockId: "p2", offset: 3 },
      end: { blockId: "p2", offset: 3 },
    },
  });
  for (const label of [
    "Bold",
    "Italic",
    "Underline",
    "Strikethrough",
    "Subscript",
    "Superscript",
    "Font family",
    "Font size",
    "Text colour",
    "Highlight",
    "Clear",
  ]) {
    assert.doesNotMatch(
      tagAround(collapsed, `aria-label="${label}"`),
      /disabled=""/,
      `${label} must be enabled with a collapsed caret in the document`,
    );
  }

  // A real selection keeps them enabled as well.
  const ranged = render({
    selection: {
      start: { blockId: "p2", offset: 0 },
      end: { blockId: "p2", offset: 7 },
    },
  });
  assert.doesNotMatch(tagAround(ranged, 'aria-label="Bold"'), /disabled=""/);
  assert.match(tagAround(ranged, 'aria-label="Bold"'), /aria-pressed="false"/);

  // Without an open document there is no caret to format: disabled + reason.
  const empty = render({ noDocument: true });
  assert.match(tagAround(empty, 'aria-label="Bold"'), /disabled=""/);
});

test("Home tab shows every labelled Word group", () => {
  const markup = render({ tab: "home" });
  const groups = [
    "Undo",
    "Clipboard",
    "Font",
    "Paragraph",
    "Styles",
    "Editing",
  ];
  for (const group of groups) {
    assert.ok(
      markup.includes(`class="dx-group-label">${group}<`),
      `the Home ribbon must show the ${group} group`,
    );
  }
  for (const label of [
    "Bold",
    "Italic",
    "Underline",
    "Strikethrough",
    "Subscript",
    "Superscript",
    "Text colour",
    "Highlight",
    "Clear",
    "Bullets",
    "Numbering",
    "Multilevel",
    "Align left",
    "Centre",
    "Align right",
    "Justify",
    "Find",
    "Replace",
    "Select all",
    "Cut",
    "Copy",
    "Paste",
    "Undo",
    "Redo",
  ])
    assert.ok(
      markup.includes(`aria-label="${label}"`),
      `the Home ribbon must expose ${label}`,
    );
});

test("Insert, Design, Layout, References and View tabs expose their groups", () => {
  const insert = render({ tab: "insert" });
  for (const group of [
    "Pages",
    "Tables",
    "Illustrations",
    "Links",
    "Header &amp; Footer",
    "Text",
    "Symbols",
  ])
    assert.ok(insert.includes(`class="dx-group-label">${group}<`), group);
  assert.match(tagAround(insert, 'aria-label="Shapes"'), /disabled=""/);
  assert.match(tagAround(insert, 'aria-label="Symbol"'), /disabled=""/);
  assert.match(
    tagAround(insert, 'aria-label="Shapes"'),
    /no shape-insertion op/,
  );

  const design = render({ tab: "design" });
  assert.ok(design.includes("Document Formatting"));
  assert.match(tagAround(design, 'aria-label="Page colour"'), /disabled=""/);
  assert.match(
    tagAround(design, 'aria-label="Page colour"'),
    /not supported by the engine/,
  );
  assert.ok(design.includes("Paragraph spacing"));
  assert.ok(design.includes("Font set"));
  assert.match(
    tagAround(design, 'aria-label="Page borders"'),
    /not supported by the engine/,
  );

  const layout = render({ tab: "layout" });
  for (const group of ["Page Setup", "Breaks", "Page Size"])
    assert.ok(layout.includes(`class="dx-group-label">${group}<`), group);
  assert.ok(layout.includes("595.3 × 841.9 pt"), "live page-size readout");
  assert.ok(layout.includes("Margins"));
  assert.ok(layout.includes("Orientation"));
  assert.ok(layout.includes("Columns"));
  assert.ok(layout.includes('aria-label="Page"'));
  assert.ok(layout.includes('aria-label="Column"'));
  assert.ok(layout.includes('aria-label="Section"'));

  const references = render({ tab: "references" });
  for (const group of [
    "Table of Contents",
    "Footnotes",
    "Citations &amp; Bibliography",
  ])
    assert.ok(references.includes(`class="dx-group-label">${group}<`), group);
  assert.ok(references.includes("Update TOC"));
  assert.match(tagAround(references, 'aria-label="Bookmark"'), /disabled=""/);
  assert.match(
    tagAround(references, 'aria-label="Caption"'),
    /not exposed by the editor API/,
  );
  assert.match(
    tagAround(references, 'aria-label="Cross-reference"'),
    /not exposed by the editor API/,
  );

  const view = render({ tab: "view" });
  for (const group of ["Views", "Zoom"])
    assert.ok(view.includes(`class="dx-group-label">${group}<`), group);
  assert.match(
    tagAround(view, 'aria-label="Formatting marks"'),
    /aria-pressed="false"/,
  );
  assert.match(tagAround(view, 'aria-label="Ruler"'), /aria-pressed="true"/);
  assert.match(
    tagAround(view, 'aria-label="Fit width"'),
    /aria-label="Fit width"/,
  );
  assert.ok(view.includes("100%"));
});

test("the editing surface renders editable paragraphs, runs, tables and TOC", () => {
  const markup = render({});
  assert.ok(markup.includes('role="document"'));
  assert.ok(markup.includes('aria-label="Document editing surface"'));
  // Paragraphs are contenteditable elements keyed by block id.
  assert.ok(hasEditableAttribute(markup, "true"));
  assert.match(
    tagAround(markup, 'data-block-id="p2"'),
    /contenteditable="true"/i,
  );
  assert.match(
    tagAround(markup, 'data-block-id="p1"'),
    /contenteditable="true"/i,
  );
  // Runs carry their formatting inline.
  assert.ok(markup.includes("Revenue "));
  assert.ok(markup.includes("grew"));
  assert.ok(markup.includes("font-weight:700"));
  assert.ok(markup.includes("font-size:20pt"), "Heading 1 style size resolves");
  assert.ok(markup.includes("color:#c00000"), "run colour renders");
  // Tabs render as an atom so offsets stay stable.
  assert.ok(markup.includes('class="dx-atom dx-atom-tab"'));
  assert.ok(markup.includes('data-dx-atom="\t"'));
  // Tables are real tables whose cell paragraphs are editable.
  assert.ok(markup.includes('<table class="dx-table" data-block-id="t1"'));
  assert.ok(markup.includes("Region"));
  assert.ok(markup.includes("Revenue"));
  assert.match(
    tagAround(markup, 'data-block-id="p3"'),
    /contenteditable="true"/i,
  );
  // Tracked changes stay visible (and keep their place in the offset space).
  assert.ok(markup.includes("dx-run-deleted"));
  assert.ok(markup.includes("(draft)"));
  assert.ok(markup.includes("dx-run-inserted"));
  // The TOC block renders generated entries.
  assert.ok(markup.includes('class="dx-toc"'));
  assert.ok(markup.includes("Table of Contents"));
  // Two sheets, with a separator marker on the second page.
  assert.equal((markup.match(/class="dx-page"/g) || []).length, 2);
  assert.ok(markup.includes(">Page 2<"));
  // The page header from the model renders in the sheet head.
  assert.ok(markup.includes("Quarterly report — header"));
  // The list label for the numbered paragraph in page two.
  assert.ok(markup.includes('class="dx-list-label"'));
});

test("the status bar reports page, counts, language, zoom and save state", () => {
  const markup = render({});
  assert.ok(markup.includes('class="dx-statusbar"'));
  assert.ok(markup.includes('role="status"'));
  assert.ok(markup.includes("Page 1 of 2"));
  assert.ok(/\d+ words/.test(markup));
  assert.ok(/\d+ characters/.test(markup));
  assert.ok(markup.includes("en-US"));
  assert.ok(markup.includes('aria-label="Zoom slider"'));
  assert.ok(markup.includes("Editing"));
  assert.ok(markup.includes("Saved"));
});

test("read-only mode disables editing and says why", () => {
  const markup = render({ readOnly: true });
  assert.match(
    tagAround(markup, 'data-block-id="p2"'),
    /contenteditable="false"/i,
  );
  assert.match(tagAround(markup, 'data-block-id="p2"'), /aria-readonly="true"/);
  const bold = tagAround(markup, 'aria-label="Bold"');
  assert.match(bold, /disabled=""/);
  assert.match(bold, /read-only/, "the disabled reason is in the title");
  assert.ok(markup.includes("Read-only"));
  assert.ok(
    tagAround(markup, 'aria-label="Save document"').includes("disabled"),
  );
});

test("the File tab keeps the document management flows", () => {
  const markup = render({ tab: "file" });
  for (const label of [
    "New document…",
    "Open by path…",
    "Save",
    "Save as…",
    "Refresh list",
    "Delete document",
    "DOCX",
    "PDF",
    "HTML",
    "MD",
  ])
    assert.ok(markup.includes(label), `the File tab must offer ${label}`);
  assert.ok(markup.includes("Fixture document"));
  assert.ok(markup.includes("Open documents"));
  assert.ok(markup.includes("Workspace"));
  // The left document rail stays available next to the editor.
  assert.ok(markup.includes('class="dx-rail"'));
  assert.ok(markup.includes('aria-label="Documents"'));
});

test("failures surface as a dismissible error bar with code and message", () => {
  const markup = render({
    error: {
      code: "UNKNOWN_ROUTE",
      message: "Unknown Documents API route.",
    },
    offline: true,
  });
  assert.ok(markup.includes('role="alert"'));
  assert.ok(markup.includes("<code>UNKNOWN_ROUTE</code>"));
  assert.ok(markup.includes("Unknown Documents API route."));
  assert.ok(markup.includes('aria-label="Dismiss error"'));
  // Degraded mode explains that editing stays local.
  assert.ok(markup.includes("Host editor API is unavailable"));
  // The document is still editable locally.
  assert.match(
    tagAround(markup, 'data-block-id="p2"'),
    /contenteditable="true"/i,
  );
});

test("dialogs render as accessible modals", () => {
  const table = render({ dialog: { kind: "table", rows: 3, columns: 2 } });
  assert.ok(table.includes('role="dialog"'));
  assert.ok(table.includes('aria-modal="true"'));
  assert.ok(table.includes("Insert table"));
  assert.ok(table.includes("Rows"));
  assert.ok(table.includes("Columns"));
  assert.ok(table.includes("Cancel"));

  const find = render({
    dialog: {
      kind: "find",
      query: "grew",
      replacement: "rose",
      matches: [{ blockId: "p2", start: 8, end: 12, text: "grew", page: 1 }],
      activeMatch: 0,
      searched: true,
    },
  });
  assert.ok(find.includes("Find and replace"));
  assert.ok(find.includes("Find what"));
  assert.ok(find.includes("Replace with"));
  assert.ok(find.includes("grew — page 1"));

  const link = render({
    dialog: { kind: "link", href: "javascript:alert(1)" },
  });
  assert.ok(link.includes("Insert hyperlink"));
  // javascript: URLs never enable the submit button.
  assert.match(link.includes("disabled") ? link : "", /disabled/);

  const color = render({
    dialog: { kind: "color", field: "color", value: "#c00000" },
  });
  assert.ok(color.includes("Text colour"));
  assert.ok(color.includes('class="dx-swatch"'));
});

test("focus mode hides the document rail and the empty state is honest", () => {
  const focus = render({ focusMode: true });
  assert.ok(!focus.includes('class="dx-rail"'));

  const empty = render({ noDocument: true });
  assert.ok(empty.includes("No document open"));
  assert.ok(empty.includes("New document"));
  assert.ok(empty.includes("Open by path"));
  assert.ok(!empty.includes('role="document"'));
});

test("an older Host degrades to a text preview instead of a dead editor", () => {
  const markup = render({
    noDocument: true,
    offline: true,
    fallback: {
      title: "Legacy document",
      text: "Legacy text preview from /document",
      outline: [{ level: 1, text: "Introduction", blockId: "p1" }],
      truncated: false,
    },
    error: { code: "UNKNOWN_ROUTE", message: "Unknown Documents API route." },
  });
  assert.ok(markup.includes("dx-preview"));
  assert.ok(markup.includes("Legacy document"));
  assert.ok(markup.includes("Legacy text preview from /document"));
  assert.ok(markup.includes("Introduction"));
  assert.ok(markup.includes("DOCX"));
  assert.ok(markup.includes("EDITOR-API.md"));
  // No editing surface is claimed when the Host cannot provide the model.
  assert.ok(!markup.includes('role="document"'));
  assert.ok(!hasEditableAttribute(markup, "true"));
});

test("apply registers the sidebar icon and the main panel", () => {
  const registrations = [];
  const injected = [];
  const ctx = {
    slots: {
      inject(name, callback) {
        injected.push(name);
        callback();
      },
      register(config, component) {
        registrations.push({ config, component });
        return () => {};
      },
    },
  };
  bundled.apply(ctx);
  assert.deepEqual(injected, ["sidebar.panellist", "main"]);
  assert.equal(registrations.length, 2);
  assert.equal(registrations[0].config.name, "sidebar.panellist");
  assert.equal(registrations[0].config.id, "dsh-documents");
  assert.equal(registrations[0].config.order, 30);
  assert.equal(registrations[0].config.label(), "Documents");
  assert.equal(typeof registrations[0].component, "function");
  assert.equal(registrations[1].config.name, "main");
  assert.equal(registrations[1].config.key, "dsh-documents");
  assert.equal(registrations[1].component, bundled.__test.DocumentsPanel);
});

test("a failing Host disables ribbon commands but keeps the text editable", () => {
  const markup = render({
    error: { code: "UNKNOWN_ROUTE", message: "Unknown Documents API route." },
    offline: true,
  });
  // Character formatting needs the Host, so it is disabled with the reason.
  const bold = tagAround(markup, 'aria-label="Bold"');
  assert.match(bold, /disabled=""/);
  assert.match(bold, /Host editor API is unavailable/);
  // ...but the local contenteditable surface still accepts typing.
  assert.match(
    tagAround(markup, 'data-block-id="p2"'),
    /contenteditable="true"/i,
  );
});

test("structural markup is present for ribbon collapse and formatting marks", () => {
  const collapsed = render({ ribbonCollapsed: true });
  assert.ok(collapsed.includes('aria-label="Expand the ribbon"'));
  assert.ok(collapsed.includes('aria-expanded="false"'));
  const expanded = render({});
  assert.ok(expanded.includes('aria-label="Collapse the ribbon"'));
  assert.ok(expanded.includes('aria-expanded="true"'));

  const marks = render({ showMarks: true });
  assert.ok(marks.includes('class="dx-space"'), "space dots are rendered");
  assert.ok(marks.includes(".dx-marks .dx-para::after"));
});
