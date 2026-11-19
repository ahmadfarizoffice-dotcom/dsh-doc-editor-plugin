/**
 * Pure helper tests for the Documents panel. Everything here is exported on
 * `__test` by `client.js`, so these assertions cover the exact functions the
 * panel uses for diffing, offset clamping, model/text mapping and page grouping
 * without a browser or a Host.
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

const requireFromWeb = createRequire(
  new URL("../../apps/web/package.json", import.meta.url),
);
const React = requireFromWeb("react");
const server = requireFromWeb("react-dom/server");

let captured = null;
globalThis.window = {
  __ModuleLoader__: {
    load(definition) {
      captured = definition;
    },
  },
};
const source = readFileSync(clientPath, "utf8");
new Function("window", source)(globalThis.window);
const bundled = captured.factory((id) => {
  if (id === "react") return React;
  throw new Error(`client.js may only require react (asked for ${id})`);
});
const __test = bundled.__test;
const fixture = JSON.parse(readFileSync(fixturePath, "utf8"));

function paragraph(...inlines) {
  return { type: "paragraph", id: "p", style: "Normal", props: {}, inlines };
}

function text(value, props) {
  return { type: "text", text: value, props: props || {} };
}

test("the module is a data-only bundle with no unsafe document sink", () => {
  assert.equal(typeof __test, "object");
  // Document text is rendered through React children only. The forbidden sink
  // must not appear anywhere in the module source.
  const forbidden = ["dangerously" + "SetInnerHTML", "innerHTML", "eval("];
  for (const token of forbidden)
    assert.equal(
      source.includes(token),
      false,
      `client.js must not contain ${token}`,
    );
  // Only React may be required.
  const requires = [...source.matchAll(/require\("([^"]+)"\)/g)].map(
    (match) => match[1],
  );
  assert.deepEqual([...new Set(requires)], ["react"]);
  // Every CSS class must be namespaced with the dx- prefix.
  const classNames = [...source.matchAll(/className:\s*"([^"]+)"/g)].flatMap(
    (match) => match[1].split(/\s+/),
  );
  for (const className of classNames) {
    if (!className || className.includes("${")) continue;
    assert.ok(className.startsWith("dx-"), `class ${className} is not dx-*`);
  }
});

test("diffTextRange returns a minimal replacement for every edit shape", () => {
  assert.equal(__test.diffTextRange("abc", "abc"), null);
  // Insert only, in the middle.
  assert.deepEqual(__test.diffTextRange("abcd", "abXYcd"), {
    start: 2,
    end: 2,
    insert: "XY",
  });
  // Insert at the end.
  assert.deepEqual(__test.diffTextRange("abc", "abcd"), {
    start: 3,
    end: 3,
    insert: "d",
  });
  // Insert at the start.
  assert.deepEqual(__test.diffTextRange("abc", "Xabc"), {
    start: 0,
    end: 0,
    insert: "X",
  });
  // Delete only, in the middle.
  assert.deepEqual(__test.diffTextRange("abcd", "ad"), {
    start: 1,
    end: 3,
    insert: "",
  });
  // Delete everything.
  assert.deepEqual(__test.diffTextRange("abc", ""), {
    start: 0,
    end: 3,
    insert: "",
  });
  // Replace a word.
  assert.deepEqual(__test.diffTextRange("hello world", "hello there"), {
    start: 6,
    end: 11,
    insert: "there",
  });
  // Replace a range with a longer string.
  assert.deepEqual(__test.diffTextRange("ab", "aXYb"), {
    start: 1,
    end: 1,
    insert: "XY",
  });
  // Empty previous text is a pure insert.
  assert.deepEqual(__test.diffTextRange("", "typed"), {
    start: 0,
    end: 0,
    insert: "typed",
  });
  // Repeated characters are handled by the prefix/suffix passes: the minimal
  // edit appends at the end of the matched prefix.
  assert.deepEqual(__test.diffTextRange("aaaa", "aaaaa"), {
    start: 4,
    end: 4,
    insert: "a",
  });
});

test("clampOffset never leaves the paragraph bounds", () => {
  assert.equal(__test.clampOffset(-4, 10), 0);
  assert.equal(__test.clampOffset(0, 10), 0);
  assert.equal(__test.clampOffset(4, 10), 4);
  assert.equal(__test.clampOffset(10, 10), 10);
  assert.equal(__test.clampOffset(99, 10), 10);
  assert.equal(__test.clampOffset(4, 0), 0);
  assert.equal(__test.clampOffset(Number.NaN, 10), 0);
  assert.equal(__test.clampOffset(3.9, 10), 3);
});

test("model to text mapping mirrors the engine's offset space", () => {
  const value = paragraph(
    text("Revenue "),
    text("grew", { bold: true }),
    { type: "tab", props: {} },
    { type: "break", kind: "line", props: {} },
    { type: "image", mediaId: "m1", width: 10, height: 10 },
    { type: "bookmark", name: "b", bookmarkId: "b1", kind: "start" },
  );
  assert.equal(__test.paragraphText(value), "Revenue grew\t\n\uFFFC");
  const spans = __test.flatInlines(value);
  assert.equal(spans.length, 6);
  assert.deepEqual(
    spans.map((span) => [span.start, span.end, span.text]),
    [
      [0, 8, "Revenue "],
      [8, 12, "grew"],
      [12, 13, "\t"],
      [13, 14, "\n"],
      [14, 15, "\uFFFC"],
      [15, 15, ""],
    ],
  );
  // Contract offsets count every text run, including tracked deletions: the
  // Host resolves the same `{blockId, offset}` with the engine's
  // `inlineTextLength`, so the client must agree character for character.
  const deleted = paragraph(
    text("keep "),
    text("gone", { revision: { id: "r1", kind: "delete" } }),
  );
  assert.equal(__test.paragraphText(deleted), "keep gone");
  // Statistics and the outline use the engine's final projection instead.
  assert.equal(__test.finalParagraphText(deleted), "keep ");
  assert.equal(
    __test.computeStats({ sections: [{ blocks: [deleted] }] }).words,
    1,
    "deleted words are not counted in the final statistics",
  );
  // Insert-revision paragraphs still exist for offsets but are empty in final.
  const inserted = {
    type: "paragraph",
    id: "px",
    props: {},
    revision: { id: "r2", kind: "insert" },
    inlines: [text("new")],
  };
  assert.equal(__test.paragraphText(inserted), "new");
  assert.equal(__test.finalParagraphText(inserted), "");
});

test("spanAtOffset and runPropsAtOffset address the right run", () => {
  const value = paragraph(text("Revenue "), text("grew", { bold: true }));
  assert.equal(__test.runPropsAtOffset(value, 0).bold, undefined);
  assert.equal(__test.runPropsAtOffset(value, 9).bold, true);
  // Straddling the end of the paragraph clamps to the last run.
  assert.equal(__test.runPropsAtOffset(value, 999).bold, true);
  const span = __test.spanAtOffset(value, 9);
  assert.equal(span.text, "grew");
});

test("applyTextReplace patches the model in place and keeps formatting", () => {
  // Insert in the middle of the first run.
  const insert = paragraph(text("Revenue "), text("grew", { bold: true }));
  __test.applyTextReplace(insert, 7, 7, "X");
  assert.equal(__test.paragraphText(insert), "RevenueX grew");
  assert.equal(__test.paragraphText(insert).length, 13);

  // Replace the bold word, keeping the prefix formatting intact.
  const replace = paragraph(
    text("Revenue "),
    text("grew", { bold: true }),
    text("!"),
  );
  __test.applyTextReplace(replace, 8, 12, "rose");
  assert.equal(__test.paragraphText(replace), "Revenue rose!");
  const runs = replace.inlines.filter((inline) => inline.type === "text");
  const insertedRun = runs.find((inline) => inline.text === "rose");
  assert.ok(insertedRun, "the replacement is its own run");
  assert.equal(
    insertedRun.props.bold,
    true,
    "the replaced text inherits the run",
  );

  // Delete across two runs.
  const remove = paragraph(
    text("abc"),
    text("def", { italic: true }),
    text("ghi"),
  );
  __test.applyTextReplace(remove, 2, 7, "");
  assert.equal(__test.paragraphText(remove), "abhi");

  // Delete the whole paragraph text keeps one empty run.
  const clear = paragraph(text("abc"), text("def", { bold: true }));
  __test.applyTextReplace(clear, 0, 6, "");
  assert.equal(__test.paragraphText(clear), "");
  assert.equal(clear.inlines.length, 1);
  assert.equal(clear.inlines[0].type, "text");

  // Inserting at the end appends.
  const append = paragraph(text("abc"));
  __test.applyTextReplace(append, 3, 3, "!");
  assert.equal(__test.paragraphText(append), "abc!");

  // Atom inlines (tabs, breaks) inside a replaced range are removed.
  const atoms = paragraph(
    text("a"),
    { type: "tab", props: {} },
    { type: "break", kind: "line", props: {} },
    text("b"),
  );
  __test.applyTextReplace(atoms, 1, 3, "-");
  assert.equal(__test.paragraphText(atoms), "a-b");

  // Adjacent runs with identical formatting merge back together.
  const merge = paragraph(text("a"), text("b"), text("c"));
  __test.applyTextReplace(merge, 1, 1, "X");
  assert.equal(__test.paragraphText(merge), "aXbc");
  assert.equal(
    merge.inlines.filter((inline) => inline.type === "text").length,
    1,
  );
});

test("run and paragraph formatting resolve through the style chain", () => {
  const heading = __test.collectBlocks(fixture).byId.get("p1").block;
  const resolved = __test.resolveRunProps(fixture, heading, heading.inlines[0]);
  assert.equal(resolved.fontFamily, "Calibri");
  assert.equal(resolved.fontSize, 20, "Heading 1 direct style wins");
  assert.equal(resolved.bold, true);

  // Heading 2 is basedOn Heading 1 and overrides the size only.
  const heading2 = {
    type: "paragraph",
    id: "h2",
    style: "Heading2",
    props: {},
    inlines: [text("Sub heading")],
  };
  const resolved2 = __test.resolveRunProps(
    fixture,
    heading2,
    heading2.inlines[0],
  );
  assert.equal(resolved2.fontSize, 16);
  assert.equal(resolved2.bold, true, "basedOn chain supplies bold");

  // Direct run props win over the paragraph style.
  const direct = {
    type: "paragraph",
    id: "d1",
    style: "Heading1",
    props: {},
    inlines: [text("small", { fontSize: 9, italic: true })],
  };
  const resolved3 = __test.resolveRunProps(fixture, direct, direct.inlines[0]);
  assert.equal(resolved3.fontSize, 9);
  assert.equal(resolved3.italic, true);

  const css = __test.runStyleToCss(resolved);
  assert.equal(css.fontSize, "20pt");
  assert.equal(css.fontWeight, 700);

  // Paragraph formatting, including indent, alignment and spacing.
  const props = __test.resolveParagraphProps(fixture, {
    type: "paragraph",
    id: "x",
    props: { alignment: "center", indent: { left: 36 } },
    inlines: [],
  });
  const paraCss = __test.paragraphStyleToCss(props);
  assert.equal(paraCss.textAlign, "center");
  assert.equal(paraCss.marginLeft, "36pt");
});

test("URL and media sanitisation rejects every disallowed scheme", () => {
  assert.equal(
    __test.sanitizeUrl("https://example.com/a?b=c"),
    "https://example.com/a?b=c",
  );
  assert.equal(
    __test.sanitizeUrl("http://example.com/"),
    "http://example.com/",
  );
  assert.equal(
    __test.sanitizeUrl("mailto:user@example.com"),
    "mailto:user@example.com",
  );
  assert.equal(__test.sanitizeUrl("javascript:alert(1)"), null);
  assert.equal(__test.sanitizeUrl("data:text/html,<script>"), null);
  assert.equal(__test.sanitizeUrl("ftp://example.com/"), null);
  assert.equal(__test.sanitizeUrl("file:///etc/passwd"), null);
  assert.equal(__test.sanitizeUrl("//example.com"), null);
  assert.equal(__test.sanitizeUrl("example.com"), null);
  assert.equal(__test.sanitizeUrl(""), null);
  assert.equal(__test.sanitizeUrl(null), null);

  const png = { id: "m1", contentType: "image/png", data: "iVBORw0KGgo=" };
  assert.equal(__test.mediaUrl(png), "data:image/png;base64,iVBORw0KGgo=");
  assert.equal(
    __test.mediaUrl({
      id: "m2",
      contentType: "image/svg+xml",
      data: "PHN2Zz4=",
    }),
    null,
    "SVG is refused",
  );
  assert.equal(
    __test.mediaUrl({
      id: "m3",
      contentType: "image/png",
      data: "not base64!",
    }),
    null,
  );
  assert.equal(__test.mediaUrl(null), null);
});

test("stats, outline and page grouping describe the fixture document", () => {
  const stats = __test.computeStats(fixture);
  assert.ok(stats.words > 0);
  assert.equal(stats.tables, 1);
  assert.equal(stats.paragraphs, 7);
  assert.ok(stats.characters <= stats.charactersWithSpaces);

  const outline = __test.computeOutline(fixture);
  assert.equal(outline.length, 1);
  assert.equal(outline[0].text, "Quarterly report");
  assert.equal(outline[0].level, 1);
  assert.equal(outline[0].blockId, "p1");

  const blocks = __test.documentBodyBlocks(fixture);
  assert.deepEqual(
    blocks.map((block) => block.id),
    ["p1", "p2", "t1", "p5", "toc1", "p6", "p7"],
  );
  const pages = __test.groupBlocksIntoPages(blocks, {
    p1: 1,
    p2: 1,
    t1: 1,
    p5: 1,
    toc1: 1,
    p6: 1,
    p7: 2,
  });
  assert.deepEqual(
    pages.map((page) => [page.page, page.blocks.map((block) => block.id)]),
    [
      [1, ["p1", "p2", "t1", "p5", "toc1", "p6"]],
      [2, ["p7"]],
    ],
  );
  // Without a Host page map every block lands on a single sheet.
  const single = __test.groupBlocksIntoPages(blocks, null);
  assert.equal(single.length, 1);
  assert.equal(single[0].blocks.length, 7);

  // A page break in the model advances the fallback page map.
  const map = __test.defaultPageMap(fixture);
  assert.equal(map.p1, 1);
  assert.equal(map.p7, 2);
});

test("numbered list labels are generated per numId and level", () => {
  const labels = __test.computeListLabels(fixture);
  assert.ok(labels.has("p7"));
  assert.equal(labels.get("p7").text, "\u2022");
  assert.equal(labels.get("p7").level, 0);
  // Paragraphs without numbering get no label.
  assert.equal(labels.has("p2"), false);
});

test("selection helpers normalise and detect collapsed ranges", () => {
  const order = new Map([
    ["p1", 0],
    ["p2", 1],
    ["t1", 2],
    ["p5", 3],
    ["toc1", 4],
    ["p6", 5],
    ["p7", 6],
  ]);
  const reversed = {
    start: { blockId: "p5", offset: 3 },
    end: { blockId: "p2", offset: 1 },
  };
  assert.deepEqual(__test.normalizeSelection(reversed, order), {
    start: { blockId: "p2", offset: 1 },
    end: { blockId: "p5", offset: 3 },
  });
  const sameBlock = {
    start: { blockId: "p2", offset: 9 },
    end: { blockId: "p2", offset: 2 },
  };
  assert.deepEqual(__test.normalizeSelection(sameBlock, order), {
    start: { blockId: "p2", offset: 2 },
    end: { blockId: "p2", offset: 9 },
  });
  assert.equal(
    __test.isCollapsed({
      start: { blockId: "p2", offset: 3 },
      end: { blockId: "p2", offset: 3 },
    }),
    true,
  );
  assert.equal(
    __test.isCollapsed({
      start: { blockId: "p2", offset: 3 },
      end: { blockId: "p2", offset: 4 },
    }),
    false,
  );
  assert.equal(
    __test.isCollapsed({
      start: { blockId: "p1", offset: 3 },
      end: { blockId: "p2", offset: 3 },
    }),
    false,
  );
});

test("split op results resolve the created paragraph id", () => {
  // The Host currently answers insertText("\n") with this shape.
  assert.equal(
    __test.newBlockIdFromOpResult(
      {
        blockIds: ["p1", "p2"],
        start: { blockId: "p1" },
        end: { blockId: "p2" },
      },
      "p1",
    ),
    "p2",
  );
  // Documented spellings are accepted too.
  assert.equal(
    __test.newBlockIdFromOpResult(
      { result: { paragraphIds: ["p2", "p3"] } },
      "p1",
    ),
    "p3",
  );
  assert.equal(
    __test.newBlockIdFromOpResult(
      { result: { createdBlocks: ["p1", "p9"] } },
      "p1",
    ),
    "p9",
  );
  // Unusable or absent results fall back to the refetch path.
  assert.equal(__test.newBlockIdFromOpResult(null, "p1"), null);
  assert.equal(__test.newBlockIdFromOpResult({}, "p1"), null);
  assert.equal(__test.newBlockIdFromOpResult({ blockIds: ["p1"] }, "p1"), null);
  assert.equal(__test.newBlockIdFromOpResult({ result: null }, "p1"), null);
});

test("collapsed-caret formatting targets the caret paragraph", () => {
  const collapsed = {
    start: { blockId: "p2", offset: 3 },
    end: { blockId: "p2", offset: 3 },
  };
  const p2 = __test.collectBlocks(fixture).byId.get("p2").block;
  const p2Length = __test.paragraphText(p2).length;
  const ops = __test.formatTextOpsForSelection(fixture, collapsed, {
    bold: true,
  });
  assert.equal(ops.length, 1, "a collapsed caret produces exactly one op");
  assert.equal(ops[0].op, "formatText");
  assert.deepEqual(ops[0].args.range, {
    start: { blockId: "p2", offset: 0 },
    end: { blockId: "p2", offset: p2Length },
  });
  assert.deepEqual(ops[0].args.props, { bold: true });
  assert.equal(ops[0].args.clear, false);
  assert.equal(__test.blockRangesForSelection(fixture, collapsed).length, 0);

  // A real selection formats exactly the selected range.
  const ranged = __test.formatTextOpsForSelection(
    fixture,
    { start: { blockId: "p2", offset: 2 }, end: { blockId: "p2", offset: 6 } },
    { italic: true },
  );
  assert.equal(ranged.length, 1);
  assert.deepEqual(ranged[0].args.range, {
    start: { blockId: "p2", offset: 2 },
    end: { blockId: "p2", offset: 6 },
  });

  // A selection spanning paragraphs produces one op per touched paragraph.
  const spanning = __test.formatTextOpsForSelection(
    fixture,
    { start: { blockId: "p1", offset: 1 }, end: { blockId: "p2", offset: 2 } },
    { bold: true },
  );
  assert.deepEqual(
    spanning.map((op) => op.args.range.start.blockId),
    ["p1", "p2"],
  );
  assert.equal(spanning[0].args.range.start.offset, 1);
  assert.equal(spanning[1].args.range.end.offset, 2);

  // A caret inside a table cell formats that cell paragraph.
  const cell = __test.formatTextOpsForSelection(
    fixture,
    { start: { blockId: "p3", offset: 0 }, end: { blockId: "p3", offset: 0 } },
    { bold: true },
  );
  assert.equal(cell.length, 1);
  assert.equal(cell[0].args.range.start.blockId, "p3");
  assert.equal(
    cell[0].args.range.end.offset,
    __test.paragraphText(__test.collectBlocks(fixture).byId.get("p3").block)
      .length,
  );

  // Clearing formatting uses the same range with `clear`.
  const cleared = __test.formatTextOpsForSelection(
    fixture,
    collapsed,
    {},
    { clear: true },
  );
  assert.equal(cleared[0].args.clear, true);

  // No selection at all yields no ops, so the panel can report CARET_REQUIRED.
  assert.deepEqual(
    __test.formatTextOpsForSelection(fixture, null, { bold: true }),
    [],
  );
  assert.deepEqual(
    __test.formatTextOpsForSelection(
      fixture,
      {
        start: { blockId: "nope", offset: 0 },
        end: { blockId: "nope", offset: 0 },
      },
      { bold: true },
    ),
    [],
  );
});

test("DOM caret mapping and typing simulation stay in sync with the model", () => {
  function fakeText(data) {
    return { nodeType: 3, data, parentNode: null };
  }
  function fakeElement(attributes, children) {
    const element = {
      nodeType: 1,
      tagName: "SPAN",
      attributes,
      childNodes: [],
      hasAttribute(name) {
        return Object.prototype.hasOwnProperty.call(attributes, name);
      },
      getAttribute(name) {
        return Object.prototype.hasOwnProperty.call(attributes, name)
          ? attributes[name]
          : null;
      },
    };
    for (const child of children) {
      child.parentNode = element;
      element.childNodes.push(child);
    }
    element.textContent = element.childNodes
      .map((child) =>
        child.nodeType === 3
          ? child.data
          : child.hasAttribute("data-dx-atom")
            ? ""
            : child.textContent,
      )
      .join("");
    return element;
  }

  const revenue = fakeElement({ "data-dx-inline": "0" }, [
    fakeText("Revenue "),
  ]);
  const grew = fakeElement({ "data-dx-inline": "1" }, [fakeText("grew")]);
  const tab = fakeElement({ "data-dx-atom": "\t" }, []);
  const world = fakeElement({ "data-dx-inline": "3" }, [fakeText("world")]);
  grew.textContent = "grew";
  world.textContent = "world";
  const block = fakeElement({ "data-block-id": "p2" }, [
    revenue,
    grew,
    tab,
    world,
  ]);

  // Rendered DOM text mirrors the model projection (atoms included).
  assert.equal(__test.domTextOfBlock(block), "Revenue grew\tworld");

  // Caret offsets resolve to DOM points the browser can restore.
  assert.deepEqual(__test.textPointAtOffset(block, 0), {
    node: revenue.childNodes[0],
    offset: 0,
  });
  assert.deepEqual(__test.textPointAtOffset(block, 12), {
    node: grew.childNodes[0],
    offset: 4,
  });
  assert.deepEqual(__test.textPointAtOffset(block, 13), {
    node: block,
    offset: 3,
  });
  assert.deepEqual(__test.textPointAtOffset(block, 18), {
    node: world.childNodes[0],
    offset: 5,
  });

  // Without a DOM document the offset reader degrades instead of throwing.
  assert.equal(__test.offsetInBlockElement(undefined, null, 0), 0);
  assert.equal(__test.offsetInBlockElement({}, {}, 0), 0);
  assert.equal(__test.blockElementFor(null), null);
  assert.equal(__test.blockElementFor({}), null);

  // Simulate one keystroke: DOM text → minimal diff → model patch. This is the
  // exact pipeline `handleInput` uses before the debounced replaceText op.
  const value = paragraph(text("Revenue "), text("grew", { bold: true }));
  const typedDom = fakeElement({ "data-block-id": "p" }, [
    fakeElement({}, [fakeText("Revenue ")]),
    fakeElement({}, [fakeText("grew more")]),
  ]);
  const typed = __test.domTextOfBlock(typedDom);
  assert.equal(typed, "Revenue grew more");
  const edit = __test.diffTextRange(__test.paragraphText(value), typed);
  assert.deepEqual(edit, { start: 12, end: 12, insert: " more" });
  __test.applyTextReplace(value, edit.start, edit.end, edit.insert);
  assert.equal(
    __test.paragraphText(value),
    typed,
    "model catches up with the DOM",
  );
  assert.deepEqual(
    value.inlines.map((inline) => inline.text),
    ["Revenue ", "grew more"],
  );
});

test("document order includes table cells so merges find the right neighbour", () => {
  const { list, byId } = __test.collectBlocks(fixture);
  assert.deepEqual(
    list.map((entry) => entry.block.id),
    ["p1", "p2", "t1", "p3", "p4", "p5", "toc1", "p6", "p7"],
  );
  assert.equal(byId.get("p3").block.type, "paragraph");
  assert.equal(byId.get("t1").block.type, "table");
  // The previous paragraph before the first table cell paragraph is the body
  // paragraph above the table, which is what Backspace-at-start merges into.
  const index = list.findIndex((entry) => entry.block.id === "p3");
  assert.equal(list[index - 1].block.id, "t1");
});

test("ribbon projections read the formatting at the caret", () => {
  const order = new Map([
    ["p1", 0],
    ["p2", 1],
  ]);
  const insideBold = __test.readFormatting(
    fixture,
    { start: { blockId: "p2", offset: 9 }, end: { blockId: "p2", offset: 9 } },
    order,
  );
  assert.equal(insideBold.bold, true);
  assert.equal(insideBold.hasRange, false);

  const plain = __test.readFormatting(
    fixture,
    { start: { blockId: "p2", offset: 1 }, end: { blockId: "p2", offset: 9 } },
    order,
  );
  assert.equal(plain.bold, false);
  assert.equal(plain.hasRange, true);

  const heading = __test.readFormatting(
    fixture,
    { start: { blockId: "p1", offset: 0 }, end: { blockId: "p1", offset: 0 } },
    order,
  );
  assert.equal(heading.styleId, "Heading1");
  assert.equal(heading.fontSize, 20);

  const centered = __test.readFormatting(
    fixture,
    { start: { blockId: "p5", offset: 0 }, end: { blockId: "p5", offset: 0 } },
    order,
  );
  assert.equal(centered.alignment, "center");
});

test("style gallery and font list come from the document model", () => {
  const gallery = __test.styleGalleryFor(fixture);
  assert.deepEqual(
    gallery.map((entry) => entry.label),
    [
      "Normal",
      "No Spacing",
      "Title",
      "Heading 1",
      "Heading 2",
      "Heading 3",
      "Quote",
      "List Paragraph",
    ],
  );
  assert.ok(gallery.every((entry) => entry.available));
  const families = __test.fontFamiliesFor(fixture);
  assert.ok(families.includes("Calibri"));
  assert.ok(families.includes("Calibri Light"));
  assert.ok(families.includes("Times New Roman"));
});

test("seed state renders without a Host and exposes layout labels", () => {
  const state = __test.seedState(fixture, { tab: "layout" });
  assert.equal(state.tab, "layout");
  assert.ok(state.bootstrap);
  assert.equal(state.bootstrap.pageCount, 2);
  const section = fixture.sections[0];
  assert.equal(__test.pageSizeLabel(section), "595.3 × 841.9 pt (portrait)");
  assert.match(__test.marginLabel(section), /^72 \/ 72 \/ 72 \/ 72 pt$/);
  // The seeded element renders through react-dom/server.
  const markup = server.renderToStaticMarkup(
    __test.renderForTest(fixture, { tab: "layout" }),
  );
  assert.ok(markup.includes("595.3 × 841.9 pt"));
});
