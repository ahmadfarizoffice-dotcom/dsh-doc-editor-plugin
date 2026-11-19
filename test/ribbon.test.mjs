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

test("ribbon height is exactly one sixteenth of the viewport", () => {
  const markup = render();
  assert.match(markup, /\.dx-ribbon\{[^}]*height:calc\(100vh \/ 16\)/);
  assert.match(markup, /\.dx-ribbon\{[^}]*overflow-x:auto/);
  // The height must not be driven by the active tab's content.
  assert.doesNotMatch(markup, /\.dx-ribbon\{[^}]*min-height:7[0-9]px/);
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
