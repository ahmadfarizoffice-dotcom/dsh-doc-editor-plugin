/**
 * Interaction tests for the pieces the e2e suites exercise: the shortcut
 * decision table (with explicit MC/DC-style coverage), the capture-listener
 * contract, selection normalisation, and the text-flush duplication guard.
 *
 * The panel is loaded exactly like the DSH client module loader loads it.
 */
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import path from "node:path";
import test from "node:test";

const here = path.dirname(fileURLToPath(import.meta.url));
const clientPath = path.join(here, "..", "client.js");
const source = readFileSync(clientPath, "utf8");

function loadPanel() {
  let captured = null;
  const windowStub = {
    __ModuleLoader__: {
      load(definition) {
        captured = definition;
      },
    },
  };
  new Function("window", source)(windowStub);
  assert.ok(captured, "the loader must capture the module definition");
  const bundled = captured.factory((id) => {
    if (id === "react") return loadReact();
    throw new Error(`client.js may only require react (asked for ${id})`);
  });
  return { bundled, windowStub };
}

function loadReact() {
  const roots = [
    new URL("../../apps/web/package.json", import.meta.url),
    new URL("../../package.json", import.meta.url),
  ];
  for (const root of roots) {
    try {
      return createRequire(root)("react");
    } catch {
      /* try the next root */
    }
  }
  throw new Error("react is not installed in this workspace");
}

function loadReactServer() {
  const roots = [
    new URL("../../apps/web/package.json", import.meta.url),
    new URL("../../package.json", import.meta.url),
  ];
  for (const root of roots) {
    try {
      return createRequire(root)("react-dom/server");
    } catch {
      /* try the next root */
    }
  }
  throw new Error("react-dom is not installed in this workspace");
}

const { bundled, windowStub } = loadPanel();
const { __test } = bundled;
const fixture = JSON.parse(
  readFileSync(path.join(here, "fixture.json"), "utf8"),
);

/* ------------------------------------------------------- shortcut table */

test("the shortcut table covers the required commands on physical codes", () => {
  const ids = __test.SHORTCUT_TABLE.map((entry) => entry.id);
  for (const required of [
    "bold",
    "italic",
    "underline",
    "undo",
    "redo",
    "redoShift",
    "save",
    "find",
    "selectAll",
    "open",
    "link",
    "numberedList",
    "bulletList",
  ]) {
    assert.ok(ids.includes(required), `${required} must be in the table`);
  }
  assert.equal(new Set(ids).size, ids.length, "command ids must be unique");
  for (const entry of __test.SHORTCUT_TABLE) {
    assert.match(entry.code, /^(Key[A-Z]|Digit[0-9])$/, entry.id);
    assert.ok(entry.scope === "editable" || entry.scope === "panel");
    assert.equal(typeof entry.repeat, "boolean");
  }
  // Shift+7/8 are physical codes, not the shifted characters ("&" and "*").
  assert.equal(
    __test.SHORTCUT_TABLE.find((entry) => entry.id === "numberedList").code,
    "Digit7",
  );
});

test("Ctrl and Cmd are equivalent, and Alt disqualifies a combo", () => {
  const base = { code: "KeyB", shiftKey: false, altKey: false };
  assert.equal(
    __test.matchShortcut(
      { ...base, ctrlKey: true },
      {
        region: "editable",
        textField: false,
        dialog: false,
        composing: false,
        repeat: false,
        outsideModal: false,
      },
    ).id,
    "bold",
  );
  assert.equal(
    __test.matchShortcut(
      { ...base, metaKey: true },
      {
        region: "editable",
        textField: false,
        dialog: false,
        composing: false,
        repeat: false,
        outsideModal: false,
      },
    ).id,
    "bold",
  );
  assert.equal(
    __test.matchShortcut(
      { ...base, ctrlKey: true, altKey: true },
      {
        region: "editable",
        textField: false,
        dialog: false,
        composing: false,
        repeat: false,
        outsideModal: false,
      },
    ),
    null,
  );
});

/**
 * MC/DC-style decision table for the panel shortcut predicate.
 *
 * Conditions: composing, outsideModal, textField, dialog, editable region,
 * repeat, modifier, shift, alt. A full truth table is 2^9 = 512 rows; JS cannot
 * measure MC/DC directly, so this builds the table and asserts that each
 * condition independently flips the outcome at least once (the MC/DC requirement
 * of independent influence), reporting the coverage count honestly.
 */
test("shortcut predicate: decision table with independent-influence coverage", () => {
  const conditions = [
    "composing",
    "outsideModal",
    "textField",
    "dialog",
    "panelRegion",
    "repeat",
    "modifier",
    "shift",
    "alt",
  ];
  const baseEnv = {
    region: "editable",
    textField: false,
    dialog: false,
    composing: false,
    repeat: false,
    outsideModal: false,
  };
  const baseEvent = {
    code: "KeyB",
    ctrlKey: true,
    shiftKey: false,
    altKey: false,
  };
  const evaluate = (flags) => {
    const env = {
      ...baseEnv,
      repeat: flags.repeat,
      composing: flags.composing,
      outsideModal: flags.outsideModal,
      textField: flags.textField,
      dialog: flags.dialog,
    };
    if (flags.panelRegion) env.region = "panel";
    const event = {
      ...baseEvent,
      code: flags.shift ? "KeyZ" : "KeyB",
      ctrlKey: flags.modifier,
      shiftKey: flags.shift,
      altKey: flags.alt,
      repeat: flags.repeat,
    };
    // The pair is (command, handled) so that a condition changing *which*
    // command runs also counts as an influence.
    const entry = __test.matchShortcut(event, env);
    return entry ? entry.id : null;
  };
  const all = Object.fromEntries(conditions.map((name) => [name, false]));
  const rows = [];
  for (let mask = 0; mask < 1 << conditions.length; mask += 1) {
    const flags = { ...all };
    conditions.forEach((name, index) => {
      flags[name] = Boolean(mask & (1 << index));
    });
    rows.push({ flags, result: evaluate(flags) });
  }
  assert.equal(rows.length, 512, "the full truth table is exercised");

  const covered = [];
  for (const condition of conditions) {
    let independent = 0;
    for (const row of rows) {
      const flipped = {
        ...row.flags,
        [condition]: !row.flags[condition],
      };
      const other = rows.find(
        (candidate) =>
          conditions.every(
            (name) =>
              name === condition || candidate.flags[name] === row.flags[name],
          ) && candidate.flags[condition] === flipped[condition],
      );
      if (other && other.result !== row.result) independent += 1;
    }
    if (independent > 0) covered.push(condition);
  }
  // Every condition must have at least one independent pair. (Composing,
  // outsideModal, textField, dialog and region veto; repeat vetoes only for
  // commands with repeat:false; modifier/shift/alt select the combo.)
  assert.deepEqual(
    covered.sort(),
    [...conditions].sort(),
    `conditions without independent influence: ${conditions
      .filter((name) => !covered.includes(name))
      .join(", ")}`,
  );
  // Coverage is reported as the number of conditions with an independent pair.
  assert.equal(covered.length, conditions.length);
});

test("an auto-repeat is ignored for toggles but allowed for history", () => {
  const env = {
    region: "editable",
    textField: false,
    dialog: false,
    composing: false,
    repeat: true,
    outsideModal: false,
  };
  assert.equal(
    __test.matchShortcut({ code: "KeyB", ctrlKey: true, repeat: true }, env),
    null,
  );
  assert.equal(
    __test.matchShortcut({ code: "KeyZ", ctrlKey: true, repeat: true }, env).id,
    "undo",
  );
});

/* --------------------------------------------------- capture contract */

test("the capture listener consumes before DSH and unregisters on unmount", () => {
  const listeners = [];
  windowStub.addEventListener = (type, handler, capture) =>
    listeners.push({ type, handler, capture });
  windowStub.removeEventListener = (type, handler, capture) => {
    const index = listeners.findIndex(
      (entry) =>
        entry.type === type &&
        entry.handler === handler &&
        entry.capture === capture,
    );
    if (index >= 0) listeners.splice(index, 1);
  };
  const target = {
    nodeType: 1,
    closest(selector) {
      return /data-block-id/.test(selector) ? { blockId: "p1" } : null;
    },
  };
  const root = { contains: (node) => node === target };
  const calls = [];
  const unbind = __test.bindShortcutCapture({
    getRoot: () => root,
    getState: () => ({ dialog: false, composing: false }),
    onCommand: (entry) => calls.push(entry.id),
  });
  assert.equal(listeners.length, 1);
  assert.equal(listeners[0].type, "keydown");
  assert.equal(
    listeners[0].capture,
    true,
    "the listener must be capture-phase",
  );

  let prevented = 0;
  let stopped = 0;
  let immediate = 0;
  const event = {
    code: "KeyB",
    key: "b",
    ctrlKey: true,
    shiftKey: false,
    altKey: false,
    repeat: false,
    isComposing: false,
    target,
    preventDefault: () => (prevented += 1),
    stopPropagation: () => (stopped += 1),
    stopImmediatePropagation: () => (immediate += 1),
  };
  listeners[0].handler(event);
  assert.deepEqual(calls, ["bold"]);
  assert.equal(prevented, 1);
  assert.equal(stopped, 1);
  assert.equal(immediate, 1);

  // A non-shortcut key is passed through untouched.
  const foreign = { ...event, code: "KeyQ" };
  listeners[0].handler(foreign);
  assert.deepEqual(calls, ["bold"]);

  unbind();
  assert.equal(listeners.length, 0);
  const diagnostics = __test.shortcutDiagnostics();
  assert.equal(diagnostics.bound, false);
  assert.equal(diagnostics.combos.length, __test.SHORTCUT_TABLE.length);
});

test("the shortcut service is optional, conflict-aware and failure-tolerant", () => {
  const registered = [];
  const service = {
    describeBinding: (binding) => ({
      conflicts: binding.code === "KeyB" ? ["sidebar.left.toggle"] : [],
      issue: null,
    }),
    register(command) {
      registered.push(command);
      return () => {};
    },
  };
  const disposers = __test.registerShortcutCommands({ shortcuts: service });
  assert.ok(registered.length > 0, "non-conflicting commands are registered");
  assert.ok(
    !registered.some((command) => command.id.endsWith(".bold")),
    "a conflicting default is skipped instead of throwing",
  );
  for (const command of registered) {
    assert.match(command.id, /^docEditor\./);
    assert.deepEqual(command.regions, ["editable", "page", "dialog"]);
    assert.equal(typeof command.resolve, "function");
    const resolution = command.resolve({ region: "editable" });
    assert.equal(resolution.status, "handled");
  }
  for (const dispose of disposers) dispose();

  const diagnostics = __test.shortcutDiagnostics();
  assert.ok(diagnostics.service.registered > 0);
  assert.ok(
    diagnostics.service.skipped.some((entry) => entry.id === "bold"),
    "the conflict is recorded, not hidden",
  );

  // A throwing or missing service never breaks the panel.
  assert.doesNotThrow(() =>
    __test.registerShortcutCommands({
      shortcuts: {
        register() {
          throw new Error("registry unavailable");
        },
      },
    }),
  );
  assert.doesNotThrow(() => __test.registerShortcutCommands({}));
  assert.doesNotThrow(() =>
    __test.registerShortcutCommands({ shortcuts: null }),
  );
});

test("composing is read from the native event behind a React synthetic event", () => {
  assert.equal(__test.eventIsComposing({ isComposing: true }), true);
  assert.equal(
    __test.eventIsComposing({ nativeEvent: { isComposing: true } }),
    true,
  );
  assert.equal(__test.eventIsComposing({ nativeEvent: {} }), false);
  assert.equal(__test.eventIsComposing(null), false);
});

/* ------------------------------------------------------- flush guard */

test("planTextFlush sends one diff per block and never re-sends in-sync text", () => {
  const host = new Map([
    ["p1", "Alpha"],
    ["p2", "Beta"],
  ]);
  const plan = __test.planTextFlush(host, [
    ["p1", "AlphaX"],
    ["p2", "Beta"],
    ["p3", "unknown"],
  ]);
  assert.equal(plan.ops.length, 1);
  assert.deepEqual(plan.ops[0], {
    blockId: "p1",
    start: 5,
    end: 5,
    insert: "X",
    text: "AlphaX",
  });
  assert.deepEqual(plan.skipped.map((entry) => entry.reason).sort(), [
    "in-sync",
    "unknown-block",
  ]);
  // Applying the planned op reproduces the pending text exactly once.
  assert.equal(
    __test.applyTextDiff(
      "Alpha",
      plan.ops[0].start,
      plan.ops[0].end,
      plan.ops[0].insert,
    ),
    "AlphaX",
  );
});

test("planTextFlush never emits two ops for the same block", () => {
  const host = new Map([["p1", ""]]);
  const plan = __test.planTextFlush(host, [
    ["p1", "a"],
    ["p1", "ab"],
  ]);
  assert.equal(plan.ops.length, 1);
  assert.equal(plan.ops[0].insert, "ab");
});

test("a diff round-trip reproduces the DOM text for random edits", () => {
  let seed = 42;
  const random = () => {
    seed = (seed * 1664525 + 1013904223) >>> 0;
    return seed / 4294967296;
  };
  const alphabet = "abc XYZ0é\n";
  for (let index = 0; index < 300; index += 1) {
    const before = Array.from(
      { length: 12 },
      () => alphabet[Math.floor(random() * alphabet.length)],
    ).join("");
    const after = Array.from(
      { length: 12 },
      () => alphabet[Math.floor(random() * alphabet.length)],
    ).join("");
    const diff = __test.diffTextRange(before, after);
    if (!diff) {
      assert.equal(before, after);
      continue;
    }
    assert.equal(
      __test.applyTextDiff(before, diff.start, diff.end, diff.insert),
      after,
    );
  }
});

/* ---------------------------------------------------- selection helpers */

test("selectionStateOf reports blocks and endpoints for multi-block ranges", () => {
  const state = __test.selectionStateOf(fixture, {
    start: { blockId: "p1", offset: 2 },
    end: { blockId: "p2", offset: 3 },
  });
  assert.ok(state);
  assert.equal(state.collapsed, false);
  assert.deepEqual(state.from, { blockId: "p1", offset: 2 });
  assert.deepEqual(state.to, { blockId: "p2", offset: 3 });
  assert.ok(state.blocks.includes("p1"));
  assert.ok(state.blocks.includes("p2"));
  const first = __test.collectBlocks(fixture).byId.get("p1").block;
  assert.equal(state.text.split("\n")[0], __test.paragraphText(first).slice(2));
});

test("selectionStateOf normalises a reversed range and flags collapses", () => {
  const reversed = __test.selectionStateOf(fixture, {
    start: { blockId: "p2", offset: 3 },
    end: { blockId: "p1", offset: 2 },
  });
  assert.deepEqual(reversed.from, { blockId: "p1", offset: 2 });
  assert.deepEqual(reversed.to, { blockId: "p2", offset: 3 });
  const collapsed = __test.selectionStateOf(fixture, {
    start: { blockId: "p1", offset: 2 },
    end: { blockId: "p1", offset: 2 },
  });
  assert.equal(collapsed.collapsed, true);
  assert.equal(collapsed.text, "");
  assert.equal(__test.selectionStateOf(fixture, null), null);
});

test("wordRangeAt selects a whole word at either edge", () => {
  assert.deepEqual(__test.wordRangeAt("Alpha paragraph.", 2), {
    start: 0,
    end: 5,
  });
  assert.deepEqual(__test.wordRangeAt("Alpha paragraph.", 5), {
    start: 0,
    end: 5,
  });
  assert.deepEqual(__test.wordRangeAt("Alpha paragraph.", 8), {
    start: 6,
    end: 15,
  });
  assert.deepEqual(__test.wordRangeAt("", 0), { start: 0, end: 0 });
});

test("Shift+Arrow focus moves across paragraph boundaries", () => {
  const p1 = __test.collectBlocks(fixture).byId.get("p1").block;
  const length = __test.paragraphText(p1).length;
  const focus = __test.moveSelectionFocus(
    fixture,
    { blockId: "p1", offset: length },
    "ArrowRight",
  );
  assert.equal(focus.blockId, "p2");
  assert.equal(focus.offset, 0);
  const back = __test.moveSelectionFocus(fixture, focus, "ArrowLeft");
  assert.equal(back.blockId, "p1");
  assert.equal(back.offset, length);
  const extension = __test.selectionForKeyboardExtension(
    fixture,
    {
      start: { blockId: "p1", offset: length },
      end: { blockId: "p1", offset: length },
    },
    null,
    "ArrowRight",
  );
  assert.deepEqual(extension.start, { blockId: "p1", offset: length });
  assert.deepEqual(extension.end, { blockId: "p2", offset: 0 });
});

test("sameSelection compares both endpoints", () => {
  const a = {
    start: { blockId: "p1", offset: 1 },
    end: { blockId: "p2", offset: 2 },
  };
  assert.equal(__test.sameSelection(a, { ...a }), true);
  assert.equal(
    __test.sameSelection(a, {
      start: { blockId: "p1", offset: 1 },
      end: { blockId: "p2", offset: 3 },
    }),
    false,
  );
  assert.equal(__test.sameSelection(null, a), false);
});

/* ------------------------------------------- Ctrl+B reservation fallback */

test("both bold chords are in the table and resolve independently", () => {
  const env = {
    region: "editable",
    textField: false,
    dialog: false,
    composing: false,
    repeat: false,
    outsideModal: false,
  };
  assert.equal(
    __test.matchShortcut({ code: "KeyB", ctrlKey: true }, env).id,
    "bold",
  );
  assert.equal(
    __test.matchShortcut({ code: "KeyB", ctrlKey: true, shiftKey: true }, env)
      .id,
    "boldShift",
  );
  const entry = __test.SHORTCUT_TABLE.find((item) => item.id === "boldShift");
  assert.equal(entry.code, "KeyB");
  assert.equal(entry.shift, true);
  assert.equal(entry.action, "bold");
  assert.equal(entry.label, "Ctrl+Shift+B");
});

test("detectBoldReservation sees only a host-owned primary+KeyB", () => {
  const own = {
    id: "docEditor.bold",
    binding: { code: "KeyB", modifiers: ["primary"] },
  };
  const host = {
    id: "sidebar.left.toggle",
    binding: { code: "KeyB", modifiers: ["primary"] },
  };
  const altHost = {
    id: "sidebar.right.toggle",
    binding: { code: "KeyB", modifiers: ["primary", "alt"] },
  };
  assert.equal(
    __test.detectBoldReservation({
      catalog: { getSnapshot: () => [own, host] },
    }),
    true,
  );
  assert.equal(
    __test.detectBoldReservation({
      catalog: { getSnapshot: () => [own, altHost] },
    }),
    false,
    "Ctrl+Alt+B is not the reserved plain chord",
  );
  assert.equal(
    __test.detectBoldReservation({
      catalog: {
        getSnapshot: () => [own, { id: "x", keys: ["Ctrl+Shift+B"] }],
      },
    }),
    false,
  );
  assert.equal(
    __test.detectBoldReservation({
      describeBinding: () => ({ conflicts: ["sidebar.left.toggle"] }),
    }),
    true,
  );
  assert.equal(
    __test.detectBoldReservation({
      describeBinding: () => ({ conflicts: ["docEditor.bold"] }),
    }),
    false,
    "our own command is not a host reservation",
  );
  assert.equal(__test.detectBoldReservation(null), false);
  assert.equal(
    __test.detectBoldReservation({
      catalog: {
        getSnapshot() {
          throw new Error("unreadable");
        },
      },
    }),
    false,
  );
});

test("the bold tooltip and the View tab switch when Ctrl+B is reserved", () => {
  const server = loadReactServer();
  const render = (options) =>
    server.renderToStaticMarkup(
      __test.renderForTest(fixture, {
        blockPages: { p1: 1, p2: 1 },
        pageCount: 1,
        tab: options.tab || "home",
        boldReserved: options.boldReserved,
      }),
    );
  const free = render({ boldReserved: false });
  assert.match(free, /title="Bold \(Ctrl\+B\)[^"]*"/);
  assert.ok(!free.includes(__test.BOLD_RESERVED_NOTE));

  const reserved = render({ boldReserved: true });
  assert.match(reserved, /title="Bold \(Ctrl\+Shift\+B\)[^"]*"/);
  assert.ok(
    !reserved.includes('title="Bold (Ctrl+B)'),
    "the reserved chord must not be advertised as the panel shortcut",
  );

  const view = render({ tab: "view", boldReserved: true });
  assert.ok(
    view.includes(__test.BOLD_RESERVED_NOTE),
    "the View tab must carry the reservation note",
  );
  const viewFree = render({ tab: "view", boldReserved: false });
  assert.ok(!viewFree.includes(__test.BOLD_RESERVED_NOTE));
});
