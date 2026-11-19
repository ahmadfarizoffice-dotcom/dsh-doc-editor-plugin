/**
 * Host half (`dsh-plugin/index.js`) folder-browsing, workspace-switching and
 * path-opening tests for the "Workspaces, folder browsing and opening files"
 * section of `dsh-plugin/EDITOR-API.md`.
 *
 * The rule these tests pin down is "any absolute local directory the Host can
 * read". Browsing, switching the workspace and opening a file must work from a
 * folder outside Home, the document workspace and the session working
 * directory, while shape rules stay: absolute paths only, `..` canonicalized
 * before use, directories and files distinguished, importable extensions only
 * for opening, and listings that never carry file contents.
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
  realpath,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import os from "node:os";
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

after(async () => {
  for (const dir of temporary.splice(0)) {
    await rm(dir, { recursive: true, force: true });
  }
});

/** A temp directory inside the session working directory (an old browse root). */
async function tempDir(prefix) {
  const dir = await mkdtemp(path.join(process.cwd(), prefix));
  temporary.push(dir);
  return dir;
}

/**
 * A temp directory the browser used to refuse: outside the document workspace
 * and outside the session working directory. Windows exposes no writable folder
 * that is *always* outside the profile, so the first writable candidate wins;
 * `os.tmpdir()` is the portable fallback.
 */
async function outsideDir(prefix) {
  const candidates = [];
  if (process.platform === "win32") {
    candidates.push(
      path.join(
        `${String(process.env.SystemDrive ?? "C:")
          .slice(0, 1)
          .toUpperCase()}:\\`,
        "Windows",
        "Temp",
      ),
    );
    candidates.push(
      path.join(process.env.PUBLIC ?? "C:\\Users\\Public", "Documents"),
    );
    candidates.push("C:\\ProgramData");
  }
  candidates.push(os.tmpdir());
  for (const base of candidates) {
    try {
      if (!(await stat(base)).isDirectory()) continue;
      const dir = await mkdtemp(path.join(base, prefix));
      temporary.push(dir);
      return dir;
    } catch {
      // Not writable here; try the next candidate.
    }
  }
  throw new Error("no writable temp folder is available for this test");
}

/** Lexical "is `candidate` inside `root`" used only to prove the setup. */
function isInsidePath(root, candidate) {
  const relative = path.relative(path.resolve(root), path.resolve(candidate));
  return (
    relative === "" ||
    (!relative.startsWith("..") && !path.isAbsolute(relative))
  );
}

/** The path a route is expected to report (realpath, not the typed path). */
function canonical(target) {
  return realpath(path.resolve(target));
}

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

test("GET /browse without a path lists Home, the workspace, the cwd and filesystem roots", async () => {
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
      drives.some((root) => root.path.toUpperCase() === system),
      `the ${system} drive root must be listed`,
    );
    assert.ok(
      res.json.roots.some((root) => root.label.startsWith("Local disk (")),
      "drive roots are offered as shortcuts",
    );
  } else {
    assert.ok(
      res.json.roots.some((root) => root.path === "/"),
      "POSIX must expose the / root",
    );
  }
});

test("GET /browse lists a folder outside Home, the workspace and the cwd", async () => {
  const workspace = await tempDir("dsh-doc-outside-ws-");
  const host = startHost({ workspaceDir: workspace });
  const base = await outsideDir("dsh-doc-anywhere-");
  // The point of the change: this folder is outside every root the Host used to
  // confine browsing to (Home is checked only when the platform allows it).
  assert.ok(
    !isInsidePath(workspace, base),
    "the fixture folder must stay outside the document workspace",
  );
  assert.ok(
    !isInsidePath(process.cwd(), base),
    "the fixture folder must stay outside the session working directory",
  );

  const folder = path.join(base, "Reports");
  await mkdir(path.join(folder, "Archive"), { recursive: true });
  const docxPath = path.join(folder, "Q3 report.docx");
  await copyFile(FIXTURE_DOCX, docxPath);
  await writeFile(path.join(folder, "notes.txt"), "notes", "utf8");
  await writeFile(path.join(folder, "data.bin"), Buffer.from([1, 2, 3]));

  const folderPath = await canonical(folder);
  const res = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(folder)}`,
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.equal(res.json.path, folderPath);
  assert.equal(res.json.parent, path.dirname(folderPath));
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
  assert.deepEqual(
    Object.keys(archive).sort(),
    ["importable", "kind", "name", "path"],
    "a listing entry carries metadata only, never file contents",
  );

  const docx = byName.get("Q3 report.docx");
  assert.equal(docx.kind, "document");
  assert.equal(docx.importable, true);
  assert.equal(docx.path, path.join(folderPath, "Q3 report.docx"));
  assert.equal(docx.size, (await stat(docxPath)).size);
  assert.ok(!Number.isNaN(Date.parse(docx.modified)));
  assert.deepEqual(
    Object.keys(docx).sort(),
    ["importable", "kind", "modified", "name", "path", "size"],
    "a file entry has size/modified but no contents field",
  );

  const text = byName.get("notes.txt");
  assert.equal(text.kind, "document");
  assert.equal(text.importable, true, "txt is importable in this build");

  const binary = byName.get("data.bin");
  assert.equal(binary.kind, "other");
  assert.equal(binary.importable, false);

  // Every entry keeps a lexical path under the declared, canonical directory.
  for (const entry of res.json.entries) {
    assert.equal(path.dirname(entry.path), folderPath);
  }
});

test("GET /browse accepts system folders that used to be refused", async () => {
  const workspace = await tempDir("dsh-doc-system-");
  const host = startHost({ workspaceDir: workspace });
  const system = process.platform === "win32" ? "C:\\Windows" : "/etc";
  const res = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(system)}`,
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.equal(res.json.path, await canonical(system));
  assert.ok(Array.isArray(res.json.entries));
  assert.equal(res.json.parent, path.dirname(await canonical(system)));
});

test("GET /browse lists a filesystem root without throwing", async () => {
  const workspace = await tempDir("dsh-doc-root-");
  const host = startHost({ workspaceDir: workspace });
  const root =
    process.platform === "win32"
      ? `${String(process.env.SystemDrive ?? "C:")
          .slice(0, 1)
          .toUpperCase()}:\\`
      : "/";
  const res = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(root)}`,
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.equal(res.json.parent, null, "a filesystem root has no parent");
  assert.ok(Array.isArray(res.json.entries));
  assert.equal(typeof res.json.skipped, "number");
  assert.ok(res.json.skipped >= 0);
  // Permission-denied entries are counted, never fatal: a root still lists.
  assert.ok(
    res.json.entries.length + res.json.skipped > 0,
    "a filesystem root returns entries or an explicit skipped count",
  );
});

test("GET /browse rejects relative input, missing folders and files", async () => {
  const workspace = await tempDir("dsh-doc-invalid-");
  const host = startHost({ workspaceDir: workspace });

  const relative = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent("relative/folder")}`,
  });
  assert.equal(relative.json.ok, false, JSON.stringify(relative.json));
  assert.equal(relative.json.error.code, "INVALID_PATH");

  const missing = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(path.join(workspace, "absent"))}`,
  });
  assert.equal(missing.json.ok, false);
  assert.equal(missing.json.error.code, "DIRECTORY_NOT_FOUND");

  const asDirectory = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(FIXTURE_DOCX)}`,
  });
  assert.equal(asDirectory.json.ok, false);
  assert.equal(asDirectory.json.error.code, "NOT_A_DIRECTORY");
});

test("GET /browse canonicalizes .. instead of confining it", async () => {
  const workspace = await tempDir("dsh-doc-dots-");
  const host = startHost({ workspaceDir: workspace });
  const base = await outsideDir("dsh-doc-dots-base-");
  const deep = path.join(base, "Reports", "2026");
  await mkdir(deep, { recursive: true });
  await writeFile(path.join(deep, "notes.txt"), "notes", "utf8");

  // Escaping out of Reports/2026 with `..` still names a real directory, so it
  // is allowed: the path is canonicalized before anything is listed.
  const traversal = path.join(deep, "..", "..");
  const res = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(traversal)}`,
  });
  assert.equal(res.json.ok, true, JSON.stringify(res.json));
  assert.equal(res.json.path, await canonical(base));
  assert.ok(
    res.json.entries.some((entry) => entry.name === "Reports"),
    "the canonicalized parent is listed",
  );

  // The same shape that resolves nowhere is a missing folder, not an escape.
  const nowhere = path.join(deep, "..", "no-such-folder-here");
  const missing = await call(host.handler, {
    route: `/browse?path=${encodeURIComponent(nowhere)}`,
  });
  assert.equal(missing.json.ok, false, JSON.stringify(missing.json));
  assert.equal(missing.json.error.code, "DIRECTORY_NOT_FOUND");
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

test("POST /workspace switches to a folder outside every old root", async () => {
  const registry = fakeRegistry();
  const workspace = await tempDir("dsh-doc-ws-outside-");
  const host = startHost(
    { workspaceDir: workspace },
    { workspaceRegistry: registry },
  );
  const outside = await outsideDir("dsh-doc-ws-switch-");
  assert.ok(!isInsidePath(workspace, outside));

  const switched = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: outside },
  });
  assert.equal(switched.json.ok, true, JSON.stringify(switched.json));
  const outsidePath = await canonical(outside);
  assert.equal(switched.json.path, outsidePath);
  assert.equal(switched.json.created, false);
  assert.equal(switched.json.dshWorkspaceId, "ws-1");
  assert.equal(registry.calls[0].method, "resolveByPath");
  assert.equal(registry.calls[0].path, outsidePath);
  assert.equal(registry.calls[1].method, "create");
  assert.equal(registry.calls[1].path, outsidePath);
  assert.equal(registry.calls[1].name, path.basename(outsidePath));

  // New documents land under that outside root, not under the old workspace.
  const documentId = await createDocument(host.handler, "Outside document");
  const model = await bootstrap(host.handler, documentId);
  assert.equal(model.workspace, outsidePath);
  const stored = path.join(
    outsidePath,
    "documents",
    documentId,
    "document.json",
  );
  assert.ok((await stat(stored)).isFile(), `expected ${stored}`);

  // The same folder reuses its registry entry instead of creating another.
  const again = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: outside },
  });
  assert.equal(again.json.ok, true, JSON.stringify(again.json));
  assert.equal(again.json.dshWorkspaceId, "ws-1");
  assert.equal(
    registry.calls.filter((row) => row.method === "create").length,
    1,
    "create must not run twice for the same folder",
  );
});

test("POST /workspace creates an absent outside folder and refuses bad targets", async () => {
  const registry = fakeRegistry();
  const workspace = await tempDir("dsh-doc-wsc-");
  const host = startHost(
    { workspaceDir: workspace },
    { workspaceRegistry: registry },
  );
  const base = await outsideDir("dsh-doc-wsc-outside-");
  const target = path.join(base, "brand", "new");
  const created = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: target, create: true },
  });
  assert.equal(created.json.ok, true, JSON.stringify(created.json));
  assert.equal(created.json.created, true);
  const targetPath = await canonical(target);
  assert.equal(created.json.path, targetPath);
  assert.equal(created.json.dshWorkspaceId, "ws-1");
  assert.ok((await stat(targetPath)).isDirectory());

  // A document created after that switch lands under the created root.
  const documentId = await createDocument(host.handler, "Brand new");
  const model = await bootstrap(host.handler, documentId);
  assert.equal(model.workspace, targetPath);
  assert.ok(
    (
      await stat(
        path.join(targetPath, "documents", documentId, "document.json"),
      )
    ).isFile(),
    "the created workspace stores its own documents",
  );

  const missing = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: path.join(base, "never-made") },
  });
  assert.equal(missing.json.ok, false);
  assert.equal(missing.json.error.code, "DIRECTORY_NOT_FOUND");

  // A relative path never names a folder, and a file is not a workspace.
  const relative = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: "relative/folder" },
  });
  assert.equal(relative.json.ok, false);
  assert.equal(relative.json.error.code, "INVALID_PATH");

  const asFolder = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: FIXTURE_DOCX },
  });
  assert.equal(asFolder.json.ok, false);
  assert.equal(asFolder.json.error.code, "NOT_A_DIRECTORY");

  // A filesystem root is a typo for a workspace, not a destination.
  const root =
    process.platform === "win32"
      ? `${String(process.env.SystemDrive ?? "C:")
          .slice(0, 1)
          .toUpperCase()}:\\`
      : "/";
  const rootRefused = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: root },
  });
  assert.equal(rootRefused.json.ok, false, JSON.stringify(rootRefused.json));
  assert.equal(rootRefused.json.error.code, "INVALID_PATH");
});

test("POST /open-path opens a DOCX from any local folder", async () => {
  const workspace = await tempDir("dsh-doc-open-");
  const host = startHost({ workspaceDir: workspace });
  const folder = await outsideDir("dsh-doc-open-any-");
  assert.ok(!isInsidePath(workspace, folder));
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
  assert.equal(opened.json.path, await canonical(docxPath));
  assert.equal(opened.json.directory, await canonical(folder));
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

  const asFile = await call(host.handler, {
    method: "POST",
    route: "/open-path",
    body: { path: folder },
  });
  assert.equal(asFile.json.ok, false);
  assert.equal(asFile.json.error.code, "NOT_A_FILE");

  const relative = await call(host.handler, {
    method: "POST",
    route: "/open-path",
    body: { path: "Seed.docx" },
  });
  assert.equal(relative.json.ok, false);
  assert.equal(relative.json.error.code, "INVALID_PATH");
});

test("POST /workspace keeps sessions opened before an outside switch", async () => {
  const workspace = await tempDir("dsh-doc-ws-retire-");
  const host = startHost({ workspaceDir: workspace });
  const oldDocument = await createDocument(
    host.handler,
    "Before switch",
    "old",
  );

  const nextRoot = await outsideDir("dsh-doc-ws-retire-next-");
  const switched = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: nextRoot },
  });
  assert.equal(switched.json.ok, true, JSON.stringify(switched.json));

  const current = await call(host.handler, { route: "/workspace" });
  assert.equal(current.json.path, await canonical(nextRoot));
  assert.ok(
    current.json.documents.some((row) => row.id === oldDocument),
    "documents opened before the switch stay listed",
  );

  // The pre-switch document is not destroyed: it still bootstraps and edits,
  // and its model keeps living in the workspace it was created in.
  const oldModel = await bootstrap(host.handler, oldDocument);
  assert.equal(oldModel.documentId, oldDocument);
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
  assert.ok(
    (
      await stat(
        path.join(workspace, "documents", oldDocument, "document.json"),
      )
    ).isFile(),
    "the old root keeps its document",
  );
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
  const base = await outsideDir("dsh-doc-ro-outside-");
  const target = path.join(base, "new-folder");
  const refused = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: target, create: true },
  });
  assert.equal(refused.json.ok, false);
  assert.equal(refused.json.error.code, "READ_ONLY_MODE");
  await assert.rejects(
    stat(target),
    "read-only mode must not create the folder",
  );

  const other = await outsideDir("dsh-doc-ro-next-");
  const switched = await call(host.handler, {
    method: "POST",
    route: "/workspace",
    body: { path: other },
  });
  assert.equal(switched.json.ok, true, JSON.stringify(switched.json));
  assert.equal(switched.json.created, false);
  assert.equal(
    switched.json.warning,
    "workspace-registry-unavailable",
    "without a registry the switch still succeeds and says why there is no id",
  );

  const current = await call(host.handler, { route: "/workspace" });
  assert.equal(current.json.readOnly, true, "read-only mode survives a switch");
  assert.equal(current.json.path, await canonical(other));

  const created = await call(host.handler, {
    method: "POST",
    route: "/create",
    body: { title: "Nope" },
  });
  assert.equal(created.json.ok, false);
  assert.equal(created.json.error.code, "READ_ONLY_MODE");
});
