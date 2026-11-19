/**
 * Host half (`dsh-plugin/index.js`) folder-browsing, workspace-switching and
 * path-opening tests for the "Workspaces, folder browsing and opening files"
 * section of `dsh-plugin/EDITOR-API.md`.
 *
 * Run with: `node --test dsh-plugin/test/host-browse.test.mjs`
 *
 * The Host module imports `./engine.js`, which only exists in the built bundle,
 * so build it once per checkout before running these tests:
 *
 *   node_modules\.bin\tsc.CMD -p packages/core/tsconfig.json
 *   node tools/build-dsh-plugin.mjs
 */
import assert from "node:assert/strict";
import {
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import { after, test } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, "../..");
const API = "/api/dsh-documents";
const FIXTURE_DOCX = path.join(
  repoRoot,
  "fixtures",
  "docx_docs_renderer_feature_test_cases.docx",
);

const { apply } = await import(
  pathToFileURL(path.join(repoRoot, "dsh-plugin", "index.js")).href
);

/* --------------------------------------------------------------- harness */

/**
 * Fake Host `workspaceRegistry`: `resolveByPath` answers only for a path that
 * was created before, and `create` records the call so tests can prove the
 * route finds an existing workspace instead of creating a second one.
 */
function fakeRegistry() {
  const byPath = new Map();
  const calls = [];
  let counter = 0;
  const key = (target) => path.resolve(target).toLowerCase();
  return {
    calls,
    async resolveByPath(target) {
      calls.push({ method: "resolveByPath", path: path.resolve(target) });
      return byPath.get(key(target));
    },
    async create(target, name) {
      calls.push({
        method: "create",
        path: path.resolve(target),
        name,
      });
      counter += 1;
      const workspace = {
        id: `ws-${counter}`,
        path: path.resolve(target),
        name,
      };
      byPath.set(key(target), workspace);
      return workspace;
    },
  };
}

/** Starts the Host half with a mock Cordis context. */
function startHost(config = {}, options = {}) {
  const routes = [];
  const warnings = [];
  const tools = [];
  const workspaceRegistry = options.workspaceRegistry;
  const webServer = {
    origin: "http://127.0.0.1:9999",
    getOrigin() {
      return this.origin;
    },
    register(entry) {
      routes.push(entry);
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
      if (name === "webServer") return webServer;
      if (name === "workspaceRegistry") return workspaceRegistry;
      return undefined;
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
  bodyBuffer,
} = {}) {
  const payload =
    bodyBuffer ??
    (body === undefined
      ? undefined
      : Buffer.from(JSON.stringify(body), "utf8"));
  return {
    method,
    url: urlPath,
    headers: {
      host: "127.0.0.1:9999",
      "x-dsh-documents": "1",
      ...headers,
    },
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
    json: contentType.includes("application/json") ? res.json() : undefined,
  };
}

/* ------------------------------------------------------------- utilities */

const temporary = [];

/** A temp directory inside the session working directory (a browse root). */
async function tempDir(prefix) {
  const dir = await mkdtemp(path.join(process.cwd(), prefix));
  temporary.push(dir);
  return dir;
}

after(async () => {
  for (const dir of temporary.splice(0)) {
    await rm(dir, { recursive: true, force: true });
  }
});

async function createDocument(handler, title, text = "Hello") {
  const res = await call(handler, {
    method: "POST",
    route: "/create",
    body: { title, text },
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  return res.json.documentId;
}

async function bootstrap(handler, documentId) {
  const res = await call(handler, {
    route: `/bootstrap?id=${encodeURIComponent(documentId)}`,
  });
  assert.equal(res.status, 200);
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  return res.json;
}

/* ------------------------------------------------------------------ tests */

test("GET /browse without a path lists Home, the workspace and the cwd", async () => {
  const workspace = await tempDir("dsh-doc-roots-");
  const host = startHost({ workspaceDir: workspace });
  const res = await call(host.handler, { route: "/browse" });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.ok(Array.isArray(res.json.roots));

  const labels = res.json.roots.map((root) => root.label);
  for (const label of ["Home", "Workspace", "Working directory"]) {
    assert.ok(labels.includes(label), `missing root ${label}`);
  }
  const workspaceRoot = res.json.roots.find(
    (root) => root.label === "Workspace",
  );
  assert.equal(workspaceRoot.path, path.resolve(workspace));
  for (const root of res.json.roots) {
    assert.ok(path.isAbsolute(root.path), `${root.label} must be absolute`);
  }

  if (process.platform === "win32") {
    const drives = res.json.roots.filter((root) =>
      /^[A-Z]:\\$/.test(root.path),
    );
    assert.ok(drives.length >= 1, "Windows must expose drive roots");
    const system = `${String(process.env.SystemDrive ?? "C:")
      .slice(0, 1)
      .toUpperCase()}:\\`;
    assert.ok(
      drives.some((root) => root.path === system),
      `the ${system} drive root must be listed`,
    );
  }
});

test("GET /browse lists directories first with kinds and importable flags", async () => {
  const workspace = await tempDir("dsh-doc-list-");
  const host = startHost({ workspaceDir: workspace });

  const base = await tempDir("dsh-doc-list-base-");
  const folder = path.join(base, "Reports");
  await mkdir(path.join(folder, "Archive"), { recursive: true });
  const docxPath = path.join(folder, "Q3 report.docx");
  await copyFile(FIXTURE_DOCX, docxPath);
  await writeFile(path.join(folder, "notes.txt"), "notes", "utf8");
  await writeFile(path.join(folder, "data.bin"), Buffer.from([1, 2, 3]));

  const res = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(folder)}`,
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.equal(res.json.path, path.resolve(folder));
  assert.equal(res.json.parent, path.resolve(base));
  assert.equal(res.json.truncated, false);
  assert.equal(res.json.skipped, 0);

  const names = res.json.entries.map((entry) => entry.name);
  assert.deepEqual(names, [
    "Archive",
    "data.bin",
    "notes.txt",
    "Q3 report.docx",
  ]);

  const byName = new Map(res.json.entries.map((entry) => [entry.name, entry]));
  const archive = byName.get("Archive");
  assert.equal(archive.kind, "directory");
  assert.equal(archive.importable, false);
  assert.equal(archive.size, undefined, "directories have no size");

  const docx = byName.get("Q3 report.docx");
  assert.equal(docx.kind, "document");
  assert.equal(docx.importable, true);
  assert.equal(docx.path, docxPath);
  assert.equal(docx.size, (await stat(docxPath)).size);
  assert.ok(!Number.isNaN(Date.parse(docx.modified)));

  const text = byName.get("notes.txt");
  assert.equal(text.kind, "document");
  assert.equal(text.importable, true, "txt is importable in this build");

  const binary = byName.get("data.bin");
  assert.equal(binary.kind, "other");
  assert.equal(binary.importable, false);
});

test("GET /browse refuses paths outside the browse roots", async () => {
  const workspace = await tempDir("dsh-doc-outside-");
  const host = startHost({ workspaceDir: workspace });
  const outside = process.platform === "win32" ? "C:\\Windows" : "/etc";
  const res = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(outside)}`,
  });
  assert.equal(res.json.ok, false, JSON.stringify(res.json));
  assert.equal(res.json.error.code, "PATH_OUTSIDE_WORKSPACE");

  const missing = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(path.join(workspace, "absent"))}`,
  });
  assert.equal(missing.json.ok, false);
  assert.equal(missing.json.error.code, "FILE_NOT_FOUND");
});

test("GET /browse caps a listing at 500 entries and says it truncated", async () => {
  const workspace = await tempDir("dsh-doc-cap-");
  const host = startHost({ workspaceDir: workspace });
  const folder = path.join(workspace, "many");
  await mkdir(folder, { recursive: true });
  await Promise.all(
    Array.from({ length: 505 }, (_, index) =>
      writeFile(
        path.join(folder, `file-${String(index).padStart(3, "0")}.txt`),
        "x",
        "utf8",
      ),
    ),
  );

  const res = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(folder)}`,
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.equal(res.json.truncated, true);
  assert.equal(res.json.entries.length, 500);
  assert.equal(res.json.entries[0].name, "file-000.txt");
});

test("GET /workspace reports the active workspace and open documents", async () => {
  const workspace = await tempDir("dsh-doc-wsget-");
  const host = startHost({ workspaceDir: workspace });
  const empty = await call(host.handler, { route: "/workspace" });
  assert.equal(empty.json.ok, true);
  assert.equal(empty.json.path, path.resolve(workspace));
  assert.equal(empty.json.readOnly, false);
  assert.deepEqual(empty.json.documents, []);

  const documentId = await createDocument(host.handler, "Listed document");
  const listed = await call(host.handler, { route: "/workspace" });
  const entry = listed.json.documents.find((row) => row.id === documentId);
  assert.ok(entry, "the open document is listed");
  assert.equal(entry.title, "Listed document");
});

test("POST /workspace switches roots, retires sessions and reuses the registry", async () => {
  const registry = fakeRegistry();
  const workspace = await tempDir("dsh-doc-ws-");
  const host = startHost(
    { workspaceDir: workspace },
    { workspaceRegistry: registry },
  );
  const oldDocument = await createDocument(
    host.handler,
    "Before switch",
    "old",
  );

  const nextRoot = await tempDir("dsh-doc-ws-next-");
  const switched = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: nextRoot },
  });
  assert.equal(switched.json.ok, true, JSON.stringify(switched.json));
  assert.equal(switched.json.path, path.resolve(nextRoot));
  assert.equal(switched.json.created, false);
  assert.equal(switched.json.dshWorkspaceId, "ws-1");
  assert.equal(registry.calls[0].method, "resolveByPath");
  assert.equal(registry.calls[0].path, path.resolve(nextRoot));
  assert.equal(registry.calls[1].method, "create");
  assert.equal(registry.calls[1].path, path.resolve(nextRoot));
  assert.equal(registry.calls[1].name, path.basename(nextRoot));

  const current = await call(host.handler, { route: "/workspace" });
  assert.equal(current.json.path, path.resolve(nextRoot));
  assert.ok(
    current.json.documents.some((row) => row.id === oldDocument),
    "documents opened before the switch stay listed",
  );

  // The pre-switch document is not destroyed: it still bootstraps, edits and
  // saves inside its original workspace.
  const oldModel = await bootstrap(host.handler, oldDocument);
  assert.equal(oldModel.workspace, path.resolve(nextRoot));
  const edited = await call(host.handler, {
    method: "POST",
    route: "/op",
    body: {
      id: oldDocument,
      op: "insertText",
      args: { position: { blockId: "p1", offset: 3 }, text: "-kept-" },
    },
  });
  assert.equal(edited.json.ok, true, JSON.stringify(edited.json));

  // New documents land in the new root.
  const newDocument = await createDocument(host.handler, "After switch", "new");
  const newModel = await bootstrap(host.handler, newDocument);
  assert.equal(newModel.workspace, path.resolve(nextRoot));
  const stored = path.join(nextRoot, "documents", newDocument, "document.json");
  assert.ok((await stat(stored)).isFile(), `expected ${stored}`);
  const previous = path.join(
    workspace,
    "documents",
    oldDocument,
    "document.json",
  );
  assert.ok((await stat(previous)).isFile(), "the old root keeps its document");

  // A second switch to the same folder reuses the registry entry.
  const again = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: nextRoot },
  });
  assert.equal(again.json.ok, true, JSON.stringify(again.json));
  assert.equal(again.json.dshWorkspaceId, "ws-1");
  assert.equal(again.json.created, false);
  assert.equal(
    registry.calls.filter((row) => row.method === "create").length,
    1,
    "create must not run twice for the same folder",
  );
  assert.equal(
    registry.calls.filter((row) => row.method === "resolveByPath").length,
    2,
  );
  const stillThere = await bootstrap(host.handler, oldDocument);
  assert.equal(stillThere.ok, true, "the retired document is still readable");
});

test("POST /workspace creates a folder with create:true and warns without a registry", async () => {
  const workspace = await tempDir("dsh-doc-wsc-");
  const host = startHost({ workspaceDir: workspace });
  const target = path.join(workspace, "brand", "new");
  const created = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: target, create: true },
  });
  assert.equal(created.json.ok, true, JSON.stringify(created.json));
  assert.equal(created.json.created, true);
  assert.equal(created.json.path, path.resolve(target));
  assert.equal(created.json.dshWorkspaceId, undefined);
  assert.equal(created.json.warning, "workspace-registry-unavailable");
  assert.ok((await stat(target)).isDirectory());

  const missing = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: path.join(workspace, "never-made") },
  });
  assert.equal(missing.json.ok, false);
  assert.equal(missing.json.error.code, "DIRECTORY_NOT_FOUND");

  const outside = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: process.platform === "win32" ? "C:\\Windows" : "/etc" },
  });
  assert.equal(outside.json.ok, false);
  assert.equal(outside.json.error.code, "PATH_OUTSIDE_WORKSPACE");
});

test("POST /open-path opens a browsed DOCX and rejects bad targets", async () => {
  const workspace = await tempDir("dsh-doc-open-");
  const host = startHost({ workspaceDir: workspace });
  const folder = await tempDir("dsh-doc-open-src-");
  const docxPath = path.join(folder, "Seed.docx");
  await copyFile(FIXTURE_DOCX, docxPath);
  await writeFile(path.join(folder, "weird.xyz"), "not a document", "utf8");

  const browsed = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(folder)}`,
  });
  const seed = browsed.json.entries.find((entry) => entry.name === "Seed.docx");
  assert.equal(seed.kind, "document");
  assert.equal(seed.importable, true);

  const opened = await call(host.handler, {
    method: "POST",
    route: "/open-path",
    body: { path: docxPath },
  });
  assert.equal(opened.json.ok, true, JSON.stringify(opened.json));
  assert.equal(typeof opened.json.documentId, "string");
  assert.equal(opened.json.path, path.resolve(docxPath));
  assert.equal(opened.json.directory, path.resolve(folder));
  assert.equal(typeof opened.json.title, "string");
  assert.ok(Array.isArray(opened.json.warnings));

  const model = await bootstrap(host.handler, opened.json.documentId);
  assert.equal(model.documentId, opened.json.documentId);

  const badFormat = await call(host.handler, {
    method: "POST",
    route: "/open-path",
    body: { path: path.join(folder, "weird.xyz") },
  });
  assert.equal(badFormat.json.ok, false);
  assert.equal(badFormat.json.error.code, "FORMAT_UNSUPPORTED");

  const missing = await call(host.handler, {
    method: "POST",
    route: "/open-path",
    body: { path: path.join(folder, "ghost.docx") },
  });
  assert.equal(missing.json.ok, false);
  assert.equal(missing.json.error.code, "FILE_NOT_FOUND");

  const outside = await call(host.handler, {
    method: "POST",
    route: "/open-path",
    body: {
      path:
        process.platform === "win32"
          ? "C:\\Windows\\System32\\drivers\\etc\\hosts.docx"
          : "/etc/hosts.docx",
    },
  });
  assert.equal(outside.json.ok, false);
  assert.equal(outside.json.error.code, "PATH_OUTSIDE_WORKSPACE");
});

test("POST /open returns the directory of the imported copy", async () => {
  const workspace = await tempDir("dsh-doc-open-upload-");
  const host = startHost({ workspaceDir: workspace });
  const documentId = await createDocument(
    host.handler,
    "Saved copy",
    "copy me",
  );
  const saved = await call(host.handler, {
    method: "POST",
    route: "/saveAs",
    body: { id: documentId, path: "exports/copy.docx" },
  });
  assert.equal(saved.json.ok, true, JSON.stringify(saved.json));

  const opened = await call(host.handler, {
    method: "POST",
    route: "/open",
    body: { path: saved.json.path },
  });
  assert.equal(opened.json.ok, true, JSON.stringify(opened.json));
  assert.equal(
    opened.json.directory,
    path.join(path.resolve(workspace), "exports"),
  );
  assert.notEqual(opened.json.documentId, documentId);
});

test("POST /open accepts a multipart upload and saves it in the workspace", async () => {
  const workspace = await tempDir("dsh-doc-open-multi-");
  const host = startHost({ workspaceDir: workspace });
  const boundary = "----dsh-documents-test-boundary";
  const bytes = await readFile(FIXTURE_DOCX);
  const bodyBuffer = Buffer.concat([
    Buffer.from(
      `--${boundary}\r\n` +
        'Content-Disposition: form-data; name="file"; filename="Uploaded.docx"\r\n' +
        "Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n",
      "utf8",
    ),
    bytes,
    Buffer.from(`\r\n--${boundary}--\r\n`, "utf8"),
  ]);

  const res = await call(host.handler, {
    method: "POST",
    route: "/open",
    bodyBuffer,
    headers: { "content-type": `multipart/form-data; boundary=${boundary}` },
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.equal(
    res.json.directory,
    path.join(path.resolve(workspace), "imports"),
  );
  assert.equal(path.dirname(res.json.path), res.json.directory);
  assert.equal((await stat(res.json.path)).size, bytes.length);
  assert.equal(typeof res.json.documentId, "string");
});

test("read-only mode refuses create:true but still switches existing folders", async () => {
  const workspace = await tempDir("dsh-doc-ro-");
  const host = startHost({ workspaceDir: workspace, readOnly: true });
  const target = path.join(workspace, "new-folder");
  const refused = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: target, create: true },
  });
  assert.equal(refused.json.ok, false);
  assert.equal(refused.json.error.code, "READ_ONLY_MODE");

  const other = await tempDir("dsh-doc-ro-next-");
  const switched = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: other },
  });
  assert.equal(switched.json.ok, true, JSON.stringify(switched.json));
  assert.equal(switched.json.created, false);

  const current = await call(host.handler, { route: "/workspace" });
  assert.equal(current.json.readOnly, true, "read-only mode survives a switch");
  assert.equal(current.json.path, path.resolve(other));

  const created = await call(host.handler, {
    method: "POST",
    route: "/create",
    body: { title: "Nope" },
  });
  assert.equal(created.json.ok, false);
  assert.equal(created.json.error.code, "READ_ONLY_MODE");
});
