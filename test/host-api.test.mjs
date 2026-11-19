/**
 * Host half (`dsh-plugin/index.js`) editor-API tests.
 *
 * Run with: `node --test dsh-plugin/test/host-api.test.mjs`
 *
 * The Host module imports `./engine.js`, which only exists in the built bundle,
 * so build it once per checkout before running these tests:
 *
 *   node_modules\.bin\tsc.CMD -p packages/core/tsconfig.json
 *   node tools/build-dsh-plugin.mjs
 *
 * The test imports the plain `dsh-plugin/index.js`. When `dsh-plugin/engine.js`
 * is absent the Host half falls back to `dist/dsh-plugin/engine.js` (the build
 * output), so nothing has to be copied into the source tree.
 */
import assert from "node:assert/strict";
import { mkdtemp, rm, stat } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { after, before, test } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, "../..");
const API = "/api/dsh-documents";

const { apply } = await import(
  pathToFileURL(path.join(repoRoot, "dsh-plugin", "index.js")).href
);

/* --------------------------------------------------------------- harness */

/**
 * Starts the Host half with a mock Cordis context whose `webServer` records
 * the route table, and returns the registered route handler.
 */
function startHost(config = {}) {
  const routes = [];
  const warnings = [];
  const tools = [];
  const webServer = {
    origin: "http://127.0.0.1:9999",
    getOrigin() {
      return this.origin;
    },
    register(options) {
      routes.push(options);
      return () => {};
    },
  };
  const ctx = {
    logger: {
      warn(message) {
        warnings.push(String(message));
      },
    },
    tools: {
      register(definition) {
        tools.push(definition);
        return () => {};
      },
    },
    effect(fn) {
      const dispose = fn();
      return () => dispose?.();
    },
    get(name) {
      return name === "webServer" ? webServer : undefined;
    },
  };
  apply(ctx, config);
  assert.equal(routes.length, 1, "the Host half must register one route table");
  return { routes, warnings, tools, handler: routes[0].handler };
}

function makeRequest({
  method = "GET",
  urlPath = "/",
  headers = {},
  body,
  omitHeader = false,
} = {}) {
  const payload =
    body === undefined ? undefined : Buffer.from(JSON.stringify(body), "utf8");
  const requestHeaders = { host: "127.0.0.1:9999", ...headers };
  if (!omitHeader) requestHeaders["x-dsh-documents"] = "1";
  return {
    method,
    url: urlPath,
    headers: requestHeaders,
    async *[Symbol.asyncIterator]() {
      if (payload) yield payload;
    },
  };
}

function makeResponse() {
  return {
    statusCode: 0,
    headers: {},
    chunks: [],
    writeHead(status, headers) {
      this.statusCode = status;
      this.headers = headers ?? {};
    },
    end(chunk) {
      if (chunk !== undefined && chunk !== null) {
        this.chunks.push(Buffer.from(chunk));
      }
    },
    bodyBuffer() {
      return Buffer.concat(this.chunks);
    },
    json() {
      return JSON.parse(this.bodyBuffer().toString("utf8"));
    },
  };
}

async function call(handler, options = {}) {
  const req = makeRequest({
    urlPath: `${API}${options.route ?? "/"}`,
    ...options,
  });
  const res = makeResponse();
  await handler(req, res);
  const contentType = String(res.headers["content-type"] ?? "");
  return {
    status: res.statusCode,
    headers: res.headers,
    body: res.bodyBuffer(),
    // `/export` streams document bytes; every other route answers JSON.
    json: contentType.includes("application/json") ? res.json() : undefined,
  };
}

/* ---------------------------------------------------- document utilities */

function bodyBlockIds(document) {
  const ids = [];
  const visit = (blocks) => {
    for (const block of blocks ?? []) {
      ids.push(block.id);
      if (block.type === "table") {
        for (const row of block.rows ?? []) {
          for (const cell of row.cells ?? []) visit(cell.blocks);
        }
      }
    }
  };
  for (const section of document.sections ?? []) visit(section.blocks);
  return ids;
}

function paragraphText(block) {
  return (block.inlines ?? [])
    .map((inline) => (inline.type === "text" ? inline.text : ""))
    .join("");
}

function documentText(document) {
  const parts = [];
  const collect = (block) => {
    if (block.type === "paragraph") parts.push(paragraphText(block));
    else if (block.type === "table") {
      for (const row of block.rows ?? []) {
        for (const cell of row.cells ?? []) {
          for (const inner of cell.blocks ?? []) collect(inner);
        }
      }
    }
  };
  for (const section of document.sections ?? []) {
    for (const block of section.blocks ?? []) collect(block);
  }
  return parts.join("\n");
}

function firstParagraph(document) {
  for (const section of document.sections ?? []) {
    for (const block of section.blocks ?? []) {
      if (block.type === "paragraph") return block;
    }
  }
  throw new Error("document has no paragraph");
}

/* ------------------------------------------------------------------ suite */

let workspace;
let host;

before(async () => {
  workspace = await mkdtemp(path.join(os.tmpdir(), "dsh-doc-host-api-"));
  host = startHost({ workspaceDir: workspace });
});

after(async () => {
  if (workspace) await rm(workspace, { recursive: true, force: true });
});

async function createDoc(title = "Quarterly report", text = "Hello world") {
  const res = await call(host.handler, {
    method: "POST",
    route: "/create",
    body: { title, text },
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  return res.json.documentId;
}

async function bootstrap(documentId) {
  const res = await call(host.handler, {
    route: `/bootstrap?id=${encodeURIComponent(documentId)}`,
  });
  assert.equal(res.status, 200);
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  return res.json;
}

async function runOp(documentId, op, args, route = "/op") {
  const body =
    route === "/ops"
      ? { id: documentId, ops: args }
      : { id: documentId, op, args };
  const res = await call(host.handler, { method: "POST", route, body });
  return res.json;
}

test("registers the tools the editor ribbon needs", () => {
  assert.deepEqual(host.warnings, [], "every DEFAULT_TOOLS name must exist");
  const names = host.tools.map((tool) => tool.name);
  for (const required of [
    "document_insert_text",
    "document_insert_image",
    "document_insert_hyperlink",
    "document_insert_page_break",
    "document_insert_section_break",
    "document_insert_heading",
    "document_insert_list",
    "document_edit_table",
    "document_undo",
    "document_redo",
  ]) {
    assert.ok(names.includes(required), `missing tool ${required}`);
  }
});

test("GET /bootstrap returns the model, blockPages, stats and outline", async () => {
  const documentId = await createDoc("Bootstrap document", "Hello world");
  const data = await bootstrap(documentId);
  assert.equal(data.documentId, documentId);
  assert.equal(data.title, "Bootstrap document");
  assert.equal(typeof data.dirty, "boolean");
  assert.ok(Number.isInteger(data.revision) && data.revision >= 0);
  assert.equal(data.sourceFormat, "json");
  assert.equal(data.readOnly, false);
  assert.equal(data.workspace, path.resolve(workspace));

  assert.ok(Array.isArray(data.document.sections));
  assert.ok(data.document.sections.length >= 1);
  const ids = bodyBlockIds(data.document);
  assert.ok(ids.length > 0, "the document needs body blocks");
  assert.deepEqual(
    Object.keys(data.blockPages).sort(),
    ids.slice().sort(),
    "blockPages must cover every body block",
  );
  for (const page of Object.values(data.blockPages)) {
    assert.ok(Number.isInteger(page) && page >= 1, "pages are 1-based");
  }

  assert.equal(typeof data.pageCount, "number");
  assert.ok(data.pageCount >= 1);
  assert.ok(data.stats.words >= 2, "stats come from the service");
  assert.equal(data.stats.tables, 0);
  assert.ok(Array.isArray(data.outline));
  assert.ok(data.formats.import.includes("docx"));
  assert.ok(data.formats.export.includes("pdf"));
  assert.equal(data.capabilities.formats, true);
  assert.equal(data.capabilities.convert, true);
  assert.equal(data.capabilities.renderer, true);
  assert.equal(typeof data.capabilities.ai, "boolean");
});

test("POST /op insertText changes text and revision; POST /save writes a DOCX", async () => {
  const documentId = await createDoc();
  const before = await bootstrap(documentId);
  const block = firstParagraph(before.document);
  const offset = paragraphText(block).length;

  const applied = await runOp(documentId, "insertText", {
    position: { blockId: block.id, offset },
    text: " again",
  });
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.equal(applied.dirty, true);
  assert.ok(applied.revision > before.revision);
  assert.equal(typeof applied.stats.words, "number");
  assert.equal(applied.pageCount, before.pageCount);

  const after = await bootstrap(documentId);
  assert.ok(documentText(after.document).includes("Hello world again"));

  const saved = await call(host.handler, {
    method: "POST",
    route: "/save",
    body: { id: documentId },
  });
  assert.equal(saved.json.ok, true, JSON.stringify(saved.json));
  assert.ok(saved.json.bytes > 0);
  assert.ok(saved.json.path.endsWith("document.docx"));
  assert.equal(saved.json.revision, after.revision);
  assert.equal(
    saved.json.path,
    path.join(workspace, "documents", documentId, "document.docx"),
  );
  const info = await stat(saved.json.path);
  assert.equal(info.size, saved.json.bytes);
});

test("POST /ops applies format and insert ops in order", async () => {
  const documentId = await createDoc("Batch document", "Formatted text");
  const before = await bootstrap(documentId);
  const block = firstParagraph(before.document);
  const data = await runOp(
    documentId,
    undefined,
    [
      {
        op: "formatText",
        args: {
          range: {
            start: { blockId: block.id, offset: 0 },
            end: { blockId: block.id, offset: 9 },
          },
          props: { bold: true },
        },
      },
      {
        op: "insertText",
        args: { position: { blockId: block.id, offset: 0 }, text: "Prelude " },
      },
    ],
    "/ops",
  );
  assert.equal(data.ok, true, JSON.stringify(data));
  assert.equal(data.applied, 2);
  assert.ok(data.revision >= before.revision + 2);
  assert.equal(data.dirty, true);
  assert.ok(
    documentText((await bootstrap(documentId)).document).includes(
      "Prelude Formatted text",
    ),
  );
});

test("setMetadata renames the document and ignores engine-owned fields", async () => {
  const documentId = await createDoc("Old title", "Rename me");
  const renamed = await runOp(documentId, "setMetadata", {
    patch: {
      title: "Renamed report",
      subject: "Q3",
      keywords: ["quarterly", "budget"],
      // Unknown/engine-owned keys must be ignored, never written.
      lastModifiedBy: "someone",
      custom: { injected: "yes" },
    },
  });
  assert.equal(renamed.ok, true, JSON.stringify(renamed));
  assert.equal(renamed.title, "Renamed report");
  const data = await bootstrap(documentId);
  assert.equal(data.title, "Renamed report");
  assert.equal(data.document.metadata.title, "Renamed report");
  assert.equal(data.document.metadata.subject, "Q3");
  assert.deepEqual(data.document.metadata.keywords, ["quarterly", "budget"]);
  assert.notEqual(data.document.metadata.lastModifiedBy, "someone");
  assert.notEqual(data.document.metadata.custom?.injected, "yes");
});

test("unknown ops are rejected with UNKNOWN_OPERATION", async () => {
  const documentId = await createDoc("Ops document", "Ops text");
  const unknown = await runOp(documentId, "explodeEverything", {});
  assert.equal(unknown.ok, false);
  assert.equal(unknown.error.code, "UNKNOWN_OPERATION");
  assert.equal(unknown.applied, 0);
  for (const name of ["insertText", "replaceText", "undo", "redo"]) {
    assert.ok(unknown.error.message.includes(name), `error must list ${name}`);
  }

  const notAString = await runOp(documentId, 42, {});
  assert.equal(notAString.error.code, "UNKNOWN_OPERATION");

  const badArgs = await runOp(documentId, "insertText", "not an object");
  assert.equal(badArgs.ok, false);
  assert.equal(badArgs.error.code, "INVALID_ARGUMENT");

  const missingArg = await runOp(documentId, "insertText", {});
  assert.equal(missingArg.ok, false);
  assert.equal(missingArg.error.code, "INVALID_ARGUMENT");
});

test("requests without the Documents header are refused", async () => {
  const res = await call(host.handler, { route: "/list", omitHeader: true });
  assert.equal(res.status, 403);
  assert.equal(res.json.ok, false);
  assert.equal(res.json.error.code, "HEADER_REQUIRED");
});

test("cross-origin requests are refused", async () => {
  const res = await call(host.handler, {
    route: "/list",
    headers: { origin: "https://evil.example" },
  });
  assert.equal(res.status, 403);
  assert.equal(res.json.error.code, "ORIGIN_REFUSED");
});

test("invalid document ids are rejected before the service is touched", async () => {
  const traversal = await call(host.handler, {
    route: `/bootstrap?id=${encodeURIComponent("../secrets")}`,
  });
  assert.equal(traversal.json.ok, false);
  assert.equal(traversal.json.error.code, "INVALID_DOCUMENT_ID");

  const opTraversal = await call(host.handler, {
    method: "POST",
    route: "/op",
    body: {
      id: "../../etc/passwd",
      op: "insertText",
      args: { position: { blockId: "p1" }, text: "x" },
    },
  });
  assert.equal(opTraversal.json.ok, false);
  assert.equal(opTraversal.json.error.code, "INVALID_DOCUMENT_ID");

  const notOpen = await call(host.handler, {
    method: "POST",
    route: "/op",
    body: {
      id: "doc-not-open",
      op: "insertText",
      args: { position: { blockId: "p1" }, text: "x" },
    },
  });
  assert.equal(notOpen.json.ok, false);
  assert.equal(notOpen.json.error.code, "DOCUMENT_NOT_OPEN");
});

test("read-only mode refuses every mutation and leaves the workspace alone", async () => {
  const readOnlyWorkspace = await mkdtemp(
    path.join(os.tmpdir(), "dsh-doc-host-api-ro-"),
  );
  try {
    const readOnlyHost = startHost({
      workspaceDir: readOnlyWorkspace,
      readOnly: true,
    });
    const list = await call(readOnlyHost.handler, { route: "/list" });
    assert.equal(list.json.ok, true);
    assert.deepEqual(list.json.documents, []);

    for (const probe of [
      { route: "/create", body: { title: "Nope" }, code: "READ_ONLY_MODE" },
      {
        route: "/op",
        body: {
          id: "doc-1",
          op: "insertText",
          args: { position: { blockId: "p1" }, text: "x" },
        },
        code: "READ_ONLY_MODE",
      },
      {
        route: "/ops",
        body: { id: "doc-1", ops: [{ op: "undo", args: {} }] },
        code: "READ_ONLY_MODE",
      },
      { route: "/save", body: { id: "doc-1" }, code: "READ_ONLY_MODE" },
      {
        route: "/saveAs",
        body: { id: "doc-1", path: "copy.docx" },
        code: "READ_ONLY_MODE",
      },
      { route: "/delete", body: { id: "doc-1" }, code: "READ_ONLY_MODE" },
    ]) {
      const res = await call(readOnlyHost.handler, {
        method: "POST",
        route: probe.route,
        body: probe.body,
      });
      assert.equal(res.json.ok, false, `${probe.route} must be refused`);
      assert.equal(res.json.error.code, probe.code, probe.route);
    }

    const after = await call(readOnlyHost.handler, { route: "/list" });
    assert.deepEqual(after.json.documents, [], "no document was created");
  } finally {
    await rm(readOnlyWorkspace, { recursive: true, force: true });
  }
});

test("undo returns the whole model and restores the text", async () => {
  const documentId = await createDoc("Undo document", "Base text");
  const before = await bootstrap(documentId);
  const block = firstParagraph(before.document);
  const marker = " TEMPORARY-MARKER";
  const offset = paragraphText(block).length;

  const inserted = await runOp(documentId, "insertText", {
    position: { blockId: block.id, offset },
    text: marker,
  });
  assert.equal(inserted.ok, true, JSON.stringify(inserted));
  const afterInsert = await bootstrap(documentId);
  assert.ok(documentText(afterInsert.document).includes(marker));

  const undone = await runOp(documentId, "undo", {});
  assert.equal(undone.ok, true, JSON.stringify(undone));
  assert.ok(undone.document && typeof undone.document === "object");
  assert.ok(Array.isArray(undone.document.sections));
  assert.equal(undone.revision, afterInsert.revision + 1);
  assert.equal(undone.dirty, true);
  assert.ok(
    !documentText(undone.document).includes(marker),
    "undo must restore the previous model",
  );

  const redone = await runOp(documentId, "redo", {});
  assert.equal(redone.ok, true, JSON.stringify(redone));
  assert.ok(redone.document);
  assert.ok(documentText(redone.document).includes(marker));
});

test("GET /search returns block-relative matches with pages", async () => {
  const documentId = await createDoc("Search document", "Find the budget line");
  const res = await call(host.handler, {
    route: `/search?id=${encodeURIComponent(documentId)}&q=budget`,
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.equal(res.json.matches.length, 1);
  const [match] = res.json.matches;
  assert.equal(match.blockId, "p1");
  assert.equal(match.start, 9);
  assert.equal(match.end, 15);
  assert.equal(match.text, "budget");
  assert.equal(match.page, 1);

  const regex = await call(host.handler, {
    route: `/search?id=${encodeURIComponent(documentId)}&q=${encodeURIComponent("b.dget")}&regex=1`,
  });
  assert.equal(regex.json.matches.length, 1);
});

test("POST /saveAs writes the format implied by the extension", async () => {
  const documentId = await createDoc("Save-as document", "Copy me");
  const target = path.join(workspace, "exports", "copy.html");
  const res = await call(host.handler, {
    method: "POST",
    route: "/saveAs",
    body: { id: documentId, path: target },
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.equal(res.json.path, target);
  assert.ok(res.json.bytes > 0);
  const info = await stat(target);
  assert.equal(info.size, res.json.bytes);

  const outside = await call(host.handler, {
    method: "POST",
    route: "/saveAs",
    body: { id: documentId, path: path.join(os.tmpdir(), "escape.docx") },
  });
  assert.equal(outside.json.ok, false);
  assert.equal(outside.json.error.code, "PATH_TRAVERSAL");
});

test("GET /document and GET /list keep working", async () => {
  const documentId = await createDoc("Listing document", "Hello listing");
  const list = await call(host.handler, { route: "/list" });
  assert.equal(list.json.ok, true);
  assert.equal(list.json.workspace, path.resolve(workspace));
  const entry = list.json.documents.find((row) => row.id === documentId);
  assert.ok(entry, "the created document is listed");
  assert.equal(entry.title, "Listing document");
  assert.equal(entry.format, "json");

  const doc = await call(host.handler, {
    route: `/document?id=${encodeURIComponent(documentId)}`,
  });
  assert.equal(doc.json.ok, true);
  assert.ok(doc.json.metadata.metadata.title === "Listing document");
  assert.ok(doc.json.text.includes("Hello listing"));
  assert.ok(Array.isArray(doc.json.outline));
  assert.equal(doc.json.sourceFormat, "json");
});

test("GET /export streams bytes and POST /delete removes the document", async () => {
  const documentId = await createDoc("Export document", "Hello export");
  const exported = await call(host.handler, {
    route: `/export?id=${encodeURIComponent(documentId)}&format=html`,
  });
  assert.equal(exported.status, 200);
  assert.match(String(exported.headers["content-type"]), /text\/html/);
  assert.ok(exported.body.length > 0);
  assert.ok(exported.body.toString("utf8").includes("Hello export"));

  const deleted = await call(host.handler, {
    method: "POST",
    route: "/delete",
    body: { id: documentId },
  });
  assert.equal(deleted.json.ok, true, JSON.stringify(deleted.json));
  assert.deepEqual(
    (await call(host.handler, { route: "/list" })).json.documents.filter(
      (row) => row.id === documentId,
    ),
    [],
  );
  const gone = await call(host.handler, {
    route: `/bootstrap?id=${encodeURIComponent(documentId)}`,
  });
  assert.equal(gone.json.ok, false);
  assert.equal(gone.json.error.code, "DOCUMENT_NOT_OPEN");
});

test("unknown routes answer 404 with NOT_FOUND", async () => {
  const res = await call(host.handler, { route: "/nope" });
  assert.equal(res.status, 404);
  assert.equal(res.json.error.code, "NOT_FOUND");
});
