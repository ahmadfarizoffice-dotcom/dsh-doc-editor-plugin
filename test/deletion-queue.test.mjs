/**
 * Deletion unit tests for the panel's key handling.
 *
 * A Backspace/Delete that arrives while a text flush is pending is *not* left
 * to the browser: the panel queues the intent, resolves the flush, re-reads the
 * model and only then deletes. These tests pin the three pieces that make that
 * deterministic — the one-character range, the batch replay, and the single
 * flight queue that orders the presses — because the e2e suite can only observe
 * the timing-dependent symptom.
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
const source = readFileSync(path.join(here, "..", "client.js"), "utf8");

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
  return bundled;
}

const { __test } = loadPanel();

function paragraph(id, text) {
  return {
    id,
    type: "paragraph",
    style: "Normal",
    props: {},
    inlines: [{ type: "text", text, props: {} }],
  };
}

function documentOf(texts) {
  return {
    id: "doc",
    sections: [
      {
        blocks: Object.keys(texts).map((id) => paragraph(id, texts[id])),
      },
    ],
  };
}

function textOf(model, blockId) {
  return __test.paragraphText(
    __test.collectBlocks(model).byId.get(blockId).block,
  );
}

/** Apply planned ranges to a copy of the model, in press order. */
function applyRanges(text, ranges) {
  let value = text;
  for (const range of ranges)
    value = value.slice(0, range.start) + value.slice(range.end);
  return value;
}

/* ------------------------------------------------------- range helper */

test("deletionRangeAt removes exactly one character in each direction", () => {
  assert.deepEqual(__test.deletionRangeAt("Alpha paragraph.", 16, "backward"), {
    start: 15,
    end: 16,
  });
  assert.deepEqual(__test.deletionRangeAt("Alpha paragraph.", 0, "forward"), {
    start: 0,
    end: 1,
  });
  // Paragraph edges are not text deletions: they belong to the merge path.
  assert.equal(__test.deletionRangeAt("Alpha paragraph.", 0, "backward"), null);
  assert.equal(__test.deletionRangeAt("Alpha paragraph.", 16, "forward"), null);
  assert.equal(__test.deletionRangeAt("", 0, "forward"), null);
  assert.equal(__test.deletionRangeAt("", 0, "backward"), null);
});

test("deletionRangeAt never splits a surrogate pair", () => {
  const text = "a😀b";
  // 😀 is two UTF-16 code units: deleting "backward" from just after it must
  // remove the whole character, and the caret must move by two units.
  assert.deepEqual(__test.deletionRangeAt(text, 3, "backward"), {
    start: 1,
    end: 3,
  });
  assert.deepEqual(__test.deletionRangeAt(text, 1, "forward"), {
    start: 1,
    end: 3,
  });
  // Inside the pair (a half caret) the caret is snapped out of the pair, so a
  // deletion still removes the whole character exactly once.
  assert.deepEqual(__test.deletionRangeAt(text, 2, "backward"), {
    start: 1,
    end: 3,
  });
  assert.deepEqual(__test.deletionRangeAt(text, 2, "forward"), {
    start: 1,
    end: 3,
  });
});

test("deletionRangeAt clamps an out-of-range caret instead of throwing", () => {
  assert.deepEqual(__test.deletionRangeAt("abc", 99, "backward"), {
    start: 2,
    end: 3,
  });
  assert.deepEqual(__test.deletionRangeAt("abc", -4, "forward"), {
    start: 0,
    end: 1,
  });
  assert.equal(__test.deletionRangeAt(null, 1, "backward"), null);
});

/* -------------------------------------------------------- batch replay */

test("planDeletions replays one Backspace per intent, in press order", () => {
  const model = documentOf({ p1: "Alpha paragraph.abc", p2: "Beta" });
  const plan = __test.planDeletions(model, { blockId: "p1", offset: 19 }, [
    { direction: "backward" },
    { direction: "backward" },
    { direction: "backward" },
  ]);
  assert.deepEqual(plan.ranges, [
    { blockId: "p1", start: 18, end: 19, direction: "backward" },
    { blockId: "p1", start: 17, end: 18, direction: "backward" },
    { blockId: "p1", start: 16, end: 17, direction: "backward" },
  ]);
  assert.deepEqual(plan.caret, { blockId: "p1", offset: 16 });
  assert.equal(plan.merge, null);
  assert.equal(
    applyRanges("Alpha paragraph.abc", plan.ranges),
    "Alpha paragraph.",
  );
});

test("planDeletions keeps deleting past the typed text instead of stopping", () => {
  // Seven Backspaces after typing "abc" remove the three typed characters and
  // then four characters of the base text: the count is what the user pressed.
  const model = documentOf({ p1: "Alpha paragraph.abc" });
  const intents = Array.from({ length: 7 }, () => ({ direction: "backward" }));
  const plan = __test.planDeletions(
    model,
    { blockId: "p1", offset: 19 },
    intents,
  );
  assert.equal(plan.ranges.length, 7);
  assert.deepEqual(plan.caret, { blockId: "p1", offset: 12 });
  assert.equal(applyRanges("Alpha paragraph.abc", plan.ranges), "Alpha paragr");
});

test("planDeletions reports a paragraph edge as a merge, not a text range", () => {
  const model = documentOf({ p1: "ab", p2: "cd" });
  const backward = __test.planDeletions(model, { blockId: "p1", offset: 1 }, [
    { direction: "backward" },
    { direction: "backward" },
  ]);
  assert.equal(backward.ranges.length, 1);
  assert.deepEqual(backward.merge, {
    direction: "backward",
    blockId: "p1",
    offset: 0,
  });
  // The press that hits the edge is consumed by the merge, like a native one;
  // only the presses after it are replayed against the merged document.
  assert.deepEqual(backward.remaining, []);
  const forward = __test.planDeletions(model, { blockId: "p1", offset: 1 }, [
    { direction: "forward" },
    { direction: "forward" },
  ]);
  assert.equal(forward.ranges.length, 1);
  assert.deepEqual(forward.merge, {
    direction: "forward",
    blockId: "p1",
    offset: 1,
  });
  assert.deepEqual(forward.remaining, []);
});

test("planDeletions hands the presses after a merge back for replay", () => {
  const model = documentOf({ p1: "ab", p2: "cd" });
  const plan = __test.planDeletions(model, { blockId: "p1", offset: 2 }, [
    { direction: "forward" },
    { direction: "forward" },
    { direction: "forward" },
  ]);
  assert.equal(plan.ranges.length, 0, "the first press is the merge itself");
  assert.ok(plan.merge);
  assert.equal(plan.remaining.length, 2, "the remaining presses are not lost");
});

test("planDeletions seeds from the intent anchor when the live caret is gone", () => {
  // A flush may rebuild the model between the key press and the replay; the
  // intent's block/offset is the anchor, and it is clamped to the new text.
  const model = documentOf({ p1: "Alpha" });
  const plan = __test.planDeletions(model, null, [
    { direction: "backward", blockId: "p1", offset: 99 },
  ]);
  assert.deepEqual(plan.ranges, [
    { blockId: "p1", start: 4, end: 5, direction: "backward" },
  ]);
  assert.deepEqual(plan.caret, { blockId: "p1", offset: 4 });
});

test("planDeletions ignores an unknown block without deleting elsewhere", () => {
  const model = documentOf({ p1: "Alpha" });
  const plan = __test.planDeletions(model, { blockId: "nope", offset: 3 }, [
    { direction: "backward" },
  ]);
  assert.deepEqual(plan.ranges, []);
  assert.equal(plan.caret, null);
  assert.equal(plan.skipped[0].reason, "no-caret");
});

/* ---------------------------------------------------------- the queue */

test("the deletion queue resolves the flush before applying a batch", async () => {
  const order = [];
  const queue = __test.createDeletionQueue({
    async flush() {
      order.push("flush");
    },
    async apply(batch) {
      order.push(`apply:${batch.length}`);
    },
  });
  await queue.enqueue({ direction: "backward" });
  assert.deepEqual(order, ["flush", "apply:1"]);
  assert.equal(queue.applied(), 1);
  assert.equal(queue.pending(), 0);
  assert.equal(queue.busy(), false);
});

test("the deletion queue applies presses in order and never twice", async () => {
  const seen = [];
  let release = null;
  let held = false;
  const queue = __test.createDeletionQueue({
    flush: () => Promise.resolve(),
    apply(batch) {
      seen.push(...batch.map((intent) => intent.offset));
      // Hold the first batch so later presses join the same drain instead of
      // starting a second one.
      if (!held) {
        held = true;
        return new Promise((resolve) => {
          release = () => resolve();
        });
      }
      return Promise.resolve();
    },
  });
  const first = queue.enqueue({ direction: "backward", offset: 5 });
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(seen, [5], "the first press opens the drain");
  const second = queue.enqueue({ direction: "backward", offset: 4 });
  const third = queue.enqueue({ direction: "backward", offset: 3 });
  assert.equal(first, second, "one drain serves every press");
  assert.equal(second, third);
  assert.equal(
    queue.pending(),
    2,
    "presses that arrived mid-drain stay queued",
  );
  release();
  await first;
  assert.deepEqual(seen, [5, 4, 3], "the queued presses are applied in order");
  assert.equal(queue.applied(), 3);
  assert.equal(queue.pending(), 0);
});

test("a failed apply clears the queue and surfaces the error once", async () => {
  const errors = [];
  const queue = __test.createDeletionQueue({
    flush: () => Promise.resolve(),
    apply() {
      throw new Error("apply failed");
    },
    onError(error) {
      errors.push(error.message);
    },
  });
  await queue.enqueue({ direction: "backward" });
  assert.deepEqual(errors, ["apply failed"]);
  assert.equal(queue.pending(), 0, "stale intents must not delete later");
  assert.equal(queue.busy(), false, "the queue is reusable after a failure");
});

test("a Backspace immediately after typing deletes exactly once", async () => {
  // This is the e2e shape: the model already holds the typed text, the Host
  // does not (its flush is in flight), and the key arrives mid-flush.
  const host = { text: "Alpha paragraph." };
  const model = documentOf({ p1: "Alpha paragraph.X" });
  let flushResolve = null;
  const queue = __test.createDeletionQueue({
    flush: () =>
      new Promise((resolve) => {
        flushResolve = () => {
          // The flush sends the typed text and the Host confirms it.
          host.text = textOf(model, "p1");
          resolve();
        };
      }),
    apply(batch) {
      const plan = __test.planDeletions(
        model,
        { blockId: "p1", offset: textOf(model, "p1").length },
        batch,
      );
      for (const range of plan.ranges) {
        const entry = __test.collectBlocks(model).byId.get(range.blockId);
        __test.applyTextReplace(entry.block, range.start, range.end, "");
      }
    },
  });
  const drain = queue.enqueue({
    direction: "backward",
    blockId: "p1",
    offset: 17,
    range: { start: 16, end: 17 },
  });
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(textOf(model, "p1"), "Alpha paragraph.X", "flush runs first");
  flushResolve();
  await drain;
  assert.equal(textOf(model, "p1"), "Alpha paragraph.");
  assert.equal(host.text, "Alpha paragraph.X");
  // The model is one character behind the confirmed Host text: the next flush
  // sends exactly one deletion, so the browser and the op never both delete.
  assert.deepEqual(__test.diffTextRange(host.text, textOf(model, "p1")), {
    start: 16,
    end: 17,
    insert: "",
  });
});
