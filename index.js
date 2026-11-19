/**
 * Host half of the DSH Documents bundle.
 *
 * Registers document tools on `ctx.tools` and, when the Web server is present,
 * the same-origin editor API used by the Documents page in the main column.
 * Every change goes through the bundled engine (`./engine.js`), which is the
 * same document service the standalone editor and MCP server use, so the panel
 * never reimplements editing logic. `EDITOR-API.md` is the contract.
 */
import os from "node:os";
import path from "node:path";
import { existsSync } from "node:fs";
import { mkdir, readdir, realpath, stat, writeFile } from "node:fs/promises";

/**
 * The engine bundle sits next to this file in the published plugin. A source
 * checkout has no `dsh-plugin/engine.js`, so fall back to the build output
 * produced by `node tools/build-dsh-plugin.mjs` (after building packages/core).
 * That keeps the plain Host half importable from tests without copying the
 * multi-megabyte bundle into the source tree.
 */
async function loadEngine() {
  try {
    return await import("./engine.js");
  } catch (error) {
    if (error?.code !== "ERR_MODULE_NOT_FOUND") throw error;
    try {
      return await import("../dist/dsh-plugin/engine.js");
    } catch {
      throw new Error(
        "dsh-documents: engine.js was not found. Build the bundle first: " +
          "`node_modules\\.bin\\tsc.CMD -p packages/core/tsconfig.json` then " +
          "`node tools/build-dsh-plugin.mjs`.",
      );
    }
  }
}

const engine = await loadEngine();
const {
  DocumentService,
  invokeTool,
  buildToolManifest,
  providerStatus,
  PAGE_PRESETS,
} = engine;

export const inject = ["tools"];

const API_PREFIX = "/api/dsh-documents";
const REQUEST_HEADER = "x-dsh-documents";

/** Ordinary JSON routes stay small; editor ops carry base64 images. */
const BODY_LIMIT = 1024 * 1024;
const LARGE_BODY_LIMIT = 16 * 1024 * 1024;
const LARGE_BODY_ROUTES = new Set(["/op", "/ops", "/save", "/saveAs"]);

/** Routes that change the document model or the workspace. */
const MUTATING_ROUTES = new Set([
  "/create",
  "/delete",
  "/op",
  "/ops",
  "/save",
  "/saveAs",
]);

/** Tools exposed to the agent. Grouped by ribbon section, kept cheap. */
const DEFAULT_TOOLS = [
  // Lifecycle and reading.
  "document_list_formats",
  "document_create",
  "document_open",
  "document_get_metadata",
  "document_get_outline",
  "document_get_text",
  "document_search",
  // Text editing.
  "document_insert_text",
  "document_replace_text",
  "document_format_text",
  "document_apply_style",
  // Structural insertion (Home/Insert ribbons).
  "document_insert_heading",
  "document_insert_list",
  "document_insert_table",
  "document_edit_table",
  "document_insert_image",
  "document_insert_hyperlink",
  "document_insert_page_break",
  "document_insert_section_break",
  // History.
  "document_undo",
  "document_redo",
  // AI.
  "document_ai_review",
  "document_ai_summarize",
  // Conversion.
  "document_convert",
  "document_export_pdf",
  "document_save",
];

function resolveWorkspace(config) {
  const configured =
    typeof config?.workspaceDir === "string" ? config.workspaceDir.trim() : "";
  if (configured) return path.resolve(configured);
  const env = (process.env.DSH_DOC_WORKSPACE ?? "").trim();
  if (env) return path.resolve(env);
  return path.join(os.homedir(), ".dsh-documents");
}

/** Folder browsing is read-only and bounded; contents are never read. */
const BROWSE_ENTRY_LIMIT = 500;
const BROWSE_SCAN_LIMIT = 5000;

/**
 * Extensions the folder browser shows as documents even when this build has no
 * importer for them (PDF, spreadsheet and presentation previews). Whether a
 * file is *importable* always comes from the format registry, never this list.
 */
const DOCUMENT_EXTENSIONS = new Set([
  "doc",
  "docm",
  "docx",
  "dot",
  "dotm",
  "dotx",
  "gdoc",
  "htm",
  "html",
  "log",
  "md",
  "mdown",
  "mkd",
  "markdown",
  "odt",
  "ott",
  "pages",
  "pdf",
  "potm",
  "potx",
  "pptm",
  "pptx",
  "rtf",
  "text",
  "txt",
  "xhtml",
  "xlsm",
  "xlsx",
  "xltm",
  "xltx",
]);

/** Case-insensitive comparison on Windows, exact elsewhere. */
function samePath(left, right) {
  const a = path.resolve(left);
  const b = path.resolve(right);
  return process.platform === "win32"
    ? a.toLowerCase() === b.toLowerCase()
    : a === b;
}

/** Real path when the target exists, `undefined` when it does not. */
async function realPathOf(target) {
  try {
    return await realpath(path.resolve(target));
  } catch {
    return undefined;
  }
}

/** Lowercased extension without the dot; empty when there is none. */
function extensionOf(file) {
  const base = path.basename(file);
  const index = base.lastIndexOf(".");
  return index > 0 ? base.slice(index + 1).toLowerCase() : "";
}

/**
 * Existing Windows drive roots (`C:\`), probed with `existsSync` so no shell or
 * subprocess is ever spawned. Empty on other platforms.
 */
function driveRoots() {
  if (process.platform !== "win32") return [];
  const candidates = [];
  if (process.env.SystemDrive) candidates.push(process.env.SystemDrive);
  for (let code = 65; code <= 90; code += 1) {
    candidates.push(String.fromCharCode(code));
  }
  const roots = [];
  const seen = new Set();
  for (const candidate of candidates) {
    const letter = String(candidate)
      .replace(/[^A-Za-z]/g, "")
      .toUpperCase();
    if (!/^[A-Z]$/.test(letter)) continue;
    const root = `${letter}:\\`;
    if (seen.has(root) || !existsSync(root)) continue;
    seen.add(root);
    roots.push(root);
  }
  return roots.sort();
}

/**
 * Every local filesystem root this Host can read. Windows exposes the drive
 * letters that exist (probed with `existsSync`, never a shell); every other
 * platform has the single `/` root.
 */
function filesystemRoots() {
  return process.platform === "win32" ? driveRoots() : ["/"];
}

/** True when `candidate` is one of the filesystem roots, not a folder to use. */
function isFilesystemRoot(candidate) {
  const resolved = path.resolve(candidate);
  return filesystemRoots().some((root) => samePath(root, resolved));
}

/**
 * The roots `/browse` offers: Home, the document workspace, the session working
 * directory, then every filesystem root (each existing drive on Windows, `/`
 * elsewhere). The panel can browse far beyond these roots now; the list is a
 * set of shortcuts, not a boundary.
 */
function browseRoots(workspaceDir) {
  const roots = [];
  const seen = new Set();
  const add = (label, target) => {
    const resolved = path.resolve(target);
    const key =
      process.platform === "win32" ? resolved.toLowerCase() : resolved;
    if (seen.has(key)) return;
    seen.add(key);
    roots.push({ label, path: resolved });
  };
  add("Home", os.homedir());
  add("Workspace", workspaceDir);
  add("Working directory", process.cwd());
  if (process.platform === "win32") {
    for (const drive of driveRoots()) {
      add(`Local disk (${drive.replace(/\\$/, "")})`, drive);
    }
  } else {
    add("File system", "/");
  }
  return roots;
}

/**
 * Validates the *shape* of a caller-supplied path before anything touches the
 * filesystem: it must be an absolute local path, never empty and never carry a
 * NUL. `..` segments are deliberately allowed here; the path is canonicalized
 * with `realpath` next, so a traversal can only name a path that really exists
 * after resolution, and anything that resolves nowhere fails as missing. That
 * keeps "no `..` escape" true without inventing a root to escape from.
 */
function requireAbsolutePath(input, what) {
  if (typeof input !== "string" || !input.trim()) {
    throw new HttpError("PATH_REQUIRED", `Give an absolute ${what}.`);
  }
  const raw = input.trim();
  if (raw.includes("\0")) {
    throw new HttpError("INVALID_PATH", "The path has an invalid character.");
  }
  if (!path.isAbsolute(raw)) {
    throw new HttpError(
      "INVALID_PATH",
      `Give an absolute ${what}: relative paths are not resolved against any root.`,
    );
  }
  return raw;
}

/**
 * Resolves a directory the folder browser may list: any absolute local
 * directory the Host can realpath, including a filesystem root and a UNC share
 * the OS resolves. Relative input, a missing target and a file are refused.
 */
async function requireBrowsableDirectory(input) {
  const canonical = await realPathOf(requireAbsolutePath(input, "folder path"));
  if (!canonical) {
    throw new HttpError("DIRECTORY_NOT_FOUND", "Folder not found.");
  }
  const info = await stat(canonical).catch(() => undefined);
  if (!info) throw new HttpError("DIRECTORY_NOT_FOUND", "Folder not found.");
  if (!info.isDirectory()) {
    throw new HttpError("NOT_A_DIRECTORY", "Not a folder.");
  }
  return canonical;
}

/**
 * Resolves a document workspace target: any absolute local directory the Host
 * can reach, created on demand. A filesystem root is refused because a
 * workspace keeps its own `documents/` and `audit.log` entries, so `C:\` or `/`
 * is a mistake rather than a destination.
 */
async function requireWorkspaceDirectory(input, { create = false } = {}) {
  const requested = requireAbsolutePath(input, "folder path");
  const canonical = await realPathOf(requested);
  if (canonical) {
    const info = await stat(canonical).catch(() => undefined);
    if (info && !info.isDirectory()) {
      throw new HttpError("NOT_A_DIRECTORY", "Not a folder.");
    }
    if (isFilesystemRoot(canonical)) {
      throw new HttpError(
        "INVALID_PATH",
        "Choose a folder for the workspace, not a filesystem root.",
      );
    }
    return { path: canonical, created: false };
  }
  if (!create) {
    throw new HttpError(
      "DIRECTORY_NOT_FOUND",
      "Folder not found; pass create:true to make it.",
    );
  }
  if (isFilesystemRoot(path.resolve(requested))) {
    throw new HttpError(
      "INVALID_PATH",
      "Choose a folder for the workspace, not a filesystem root.",
    );
  }
  await mkdir(path.resolve(requested), { recursive: true });
  const createdPath = await realPathOf(requested);
  if (!createdPath) {
    throw new HttpError("DIRECTORY_NOT_FOUND", "Folder not found.");
  }
  return { path: createdPath, created: true };
}

/** Resolves an open target: any absolute local file the Host can reach. */
async function requireOpenFile(input) {
  const canonical = await realPathOf(requireAbsolutePath(input, "file path"));
  if (!canonical) throw new HttpError("FILE_NOT_FOUND", "File not found.");
  const info = await stat(canonical).catch(() => undefined);
  if (!info) throw new HttpError("FILE_NOT_FOUND", "File not found.");
  if (!info.isFile()) throw new HttpError("NOT_A_FILE", "Not a file.");
  return canonical;
}

/** Parent for the browser's Up button, or null at a filesystem root. */
function browseParent(resolved) {
  const parent = path.dirname(resolved);
  return samePath(parent, resolved) ? null : parent;
}

/** Directory entries first, then by name (numeric-aware, case-insensitive). */
function compareBrowseEntries(left, right) {
  const leftDir = left.kind === "directory";
  const rightDir = right.kind === "directory";
  if (leftDir !== rightDir) return leftDir ? -1 : 1;
  return left.name.localeCompare(right.name, undefined, {
    numeric: true,
    sensitivity: "base",
  });
}

/**
 * Lists one directory: names, kinds and sizes only — file contents are never
 * read. Every entry keeps its lexical path under the listed (already
 * canonicalized) directory; a symlink is listed without changing that declared
 * path, and browsing it later re-validates wherever it really points. An entry
 * whose `stat` fails is counted as `skipped` instead of failing the listing.
 */
async function browseDirectory(dir, importable) {
  const dirents = await readdir(dir, { withFileTypes: true });
  const scanned = dirents.slice(0, BROWSE_SCAN_LIMIT);
  const entries = [];
  let skipped = 0;
  for (const dirent of scanned) {
    const full = path.join(dir, dirent.name);
    let info;
    try {
      info = await stat(full);
    } catch {
      skipped += 1;
      continue;
    }
    const isDirectory = info.isDirectory();
    const extension = isDirectory ? "" : extensionOf(dirent.name);
    const canImport = !isDirectory && importable.has(extension);
    const kind = isDirectory
      ? "directory"
      : canImport || DOCUMENT_EXTENSIONS.has(extension)
        ? "document"
        : "other";
    const entry = {
      name: dirent.name,
      path: full,
      kind,
      importable: canImport,
    };
    if (!isDirectory) {
      entry.size = info.size;
      entry.modified = info.mtime.toISOString();
    }
    entries.push(entry);
  }
  entries.sort(compareBrowseEntries);
  const truncated =
    dirents.length > scanned.length || entries.length > BROWSE_ENTRY_LIMIT;
  return {
    entries: entries.slice(0, BROWSE_ENTRY_LIMIT),
    truncated,
    skipped,
  };
}

function json(res, status, value) {
  const body = Buffer.from(JSON.stringify(value), "utf8");
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "content-length": String(body.length),
    "cache-control": "no-store",
  });
  res.end(body);
}

function originAllowed(req, origin) {
  const value = req.headers.origin;
  if (value === undefined || value === "") return true;
  try {
    return new URL(value).origin === new URL(origin).origin;
  } catch {
    return false;
  }
}

/** Route-level failure with a stable error code (never document content). */
class HttpError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

async function readJsonBody(req, limit = BODY_LIMIT) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > limit) {
      throw new HttpError(
        "PAYLOAD_TOO_LARGE",
        `Request body exceeds the ${Math.round(limit / (1024 * 1024))} MiB limit for this route.`,
      );
    }
    chunks.push(chunk);
  }
  if (total === 0) return {};
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw new HttpError("INVALID_JSON", "Request body is not valid JSON.");
  }
}

/** Reads the raw request body, bounded by `limit` bytes. */
async function readBodyBuffer(req, limit = LARGE_BODY_LIMIT) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > limit) {
      throw new HttpError(
        "PAYLOAD_TOO_LARGE",
        `Request body exceeds the ${Math.round(limit / (1024 * 1024))} MiB limit for this route.`,
      );
    }
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

/** The multipart boundary, or undefined for a non-multipart request. */
function multipartBoundary(req) {
  const type = String(req.headers?.["content-type"] ?? "");
  if (!/^multipart\/form-data/i.test(type)) return undefined;
  const match = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(type);
  return match ? (match[1] ?? match[2] ?? "").trim() || undefined : undefined;
}

/**
 * Extracts the first file part of a multipart body. The Documents page uploads
 * exactly one document, so the parser stays deliberately small: it walks the
 * boundary markers, reads the part headers and keeps that part's bytes.
 */
function firstFilePart(buffer, boundary) {
  const marker = Buffer.from(`--${boundary}`);
  const separator = Buffer.from("\r\n\r\n");
  let cursor = buffer.indexOf(marker);
  while (cursor !== -1) {
    const start = cursor + marker.length;
    // `--` right after the marker ends the body.
    if (buffer.subarray(start, start + 2).toString("latin1") === "--") break;
    const headersStart = start + 2; // CRLF after the boundary line
    const headersEnd = buffer.indexOf(separator, headersStart);
    if (headersEnd === -1) break;
    const headers = buffer
      .subarray(headersStart, headersEnd)
      .toString("latin1");
    const bodyStart = headersEnd + separator.length;
    const next = buffer.indexOf(marker, bodyStart);
    if (next === -1) break;
    const bodyEnd = next - 2; // CRLF before the next boundary
    const disposition = /content-disposition:([^\r\n]+)/i.exec(headers)?.[1];
    if (disposition) {
      const name = /filename="([^"]*)"/i.exec(disposition)?.[1];
      if (name !== undefined) {
        return {
          filename: name,
          data: buffer.subarray(bodyStart, Math.max(bodyStart, bodyEnd)),
        };
      }
    }
    cursor = next;
  }
  return undefined;
}

/** Keeps an uploaded filename a single, harmless path segment. */
function safeUploadName(name) {
  const base = path
    .basename(String(name ?? ""))
    .replace(/[^\w .-]+/g, "_")
    .replace(/^\.+/, "")
    .slice(0, 120);
  return base || "upload";
}

function safeId(value) {
  const id = String(value ?? "");
  if (!/^[A-Za-z0-9_-]{1,64}$/.test(id)) {
    throw new HttpError("INVALID_DOCUMENT_ID", "Invalid document id.");
  }
  return id;
}

function errorPayload(error) {
  const code =
    typeof error?.code === "string" && error.code
      ? error.code
      : "DOCUMENT_ERROR";
  const message = error instanceof Error ? error.message : String(error);
  return { code, message };
}

/* ------------------------------------------------------------ arg helpers */

function asObject(value) {
  return value && typeof value === "object" && !Array.isArray(value)
    ? value
    : undefined;
}

function optionalString(value) {
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function optionalNumber(value) {
  return typeof value === "number" && Number.isFinite(value)
    ? value
    : undefined;
}

function optionalInteger(value) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.trunc(number) : undefined;
}

function integer(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.trunc(number) : fallback;
}

function positiveInteger(value, field) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 1) {
    throw new HttpError(
      "INVALID_ARGUMENT",
      `${field} must be a positive integer.`,
    );
  }
  return Math.trunc(number);
}

function stringArray(value, field) {
  if (!Array.isArray(value) || value.length === 0) {
    throw new HttpError(
      "INVALID_ARGUMENT",
      `${field} must be a non-empty array of block ids.`,
    );
  }
  return value.map((entry) => String(entry));
}

function queryFlag(params, name) {
  const value = params.get(name);
  return value === "1" || value === "true" || value === "yes";
}

/** Resolves a caller-supplied save/image path inside the workspace. */
function workspacePath(root, value) {
  if (typeof value !== "string" || !value.trim()) return undefined;
  const input = value.trim();
  return path.isAbsolute(input)
    ? path.resolve(input)
    : path.resolve(root, input);
}

/* ------------------------------------------------------- operation helpers */

const SECTION_BREAK_TYPES = new Set([
  "nextPage",
  "continuous",
  "oddPage",
  "evenPage",
]);
const HEADER_FOOTER_PARTS = new Set(["header", "footer"]);
const HEADER_FOOTER_VARIANTS = new Set(["default", "first", "even"]);
const BREAK_KINDS = new Set(["page", "line", "column"]);

function requireArgs(op, args, keys) {
  for (const key of keys) {
    if (args[key] === undefined || args[key] === null) {
      throw new HttpError("INVALID_ARGUMENT", `${op} requires args.${key}.`);
    }
  }
}

function headingLevel(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    throw new HttpError(
      "INVALID_ARGUMENT",
      "insertHeading requires a numeric args.level.",
    );
  }
  return Math.min(9, Math.max(1, Math.trunc(number)));
}

function listKind(value) {
  if (value === "bullet" || value === "number") return value;
  throw new HttpError(
    "INVALID_ARGUMENT",
    'insertList requires args.kind "bullet" or "number".',
  );
}

function tableData(value) {
  if (value === undefined) return undefined;
  if (!Array.isArray(value)) {
    throw new HttpError("INVALID_ARGUMENT", "data must be a 2-D array.");
  }
  return value.map((row) =>
    Array.isArray(row) ? row.map((cell) => String(cell ?? "")) : [],
  );
}

/** Finds a page-size preset by id or label, case-insensitively. */
function findPagePreset(value) {
  if (typeof value !== "string") return undefined;
  const wanted = value.trim().toLowerCase();
  if (!wanted) return undefined;
  const entries = Object.entries(PAGE_PRESETS ?? {});
  for (const [name, preset] of entries) {
    if (name.toLowerCase() === wanted) return preset;
    if (String(preset.label).toLowerCase() === wanted) return preset;
  }
  for (const [, preset] of entries) {
    if (String(preset.label).toLowerCase().startsWith(wanted)) return preset;
  }
  return undefined;
}

function marginPatch(value) {
  const source = asObject(value);
  if (!source) return undefined;
  const out = {};
  for (const key of [
    "top",
    "right",
    "bottom",
    "left",
    "header",
    "footer",
    "gutter",
  ]) {
    const number = optionalNumber(source[key]);
    if (number !== undefined && number >= 0) out[key] = number;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Maps the contract's `setPageLayout` args (`pageSize`/`orientation`/`margins`)
 * onto the engine's section-page patch, expanding named presets to points.
 */
function pageLayoutArgs(args) {
  const page = {};
  const size = args.pageSize;
  if (typeof size === "string" && size.trim()) {
    const preset = findPagePreset(size);
    if (!preset) {
      throw new HttpError("INVALID_ARGUMENT", `Unknown page size: ${size}.`);
    }
    page.width = preset.width;
    page.height = preset.height;
    page.preset = preset.label;
  } else if (asObject(size)) {
    const width = optionalNumber(size.width);
    const height = optionalNumber(size.height);
    if (width !== undefined) page.width = width;
    if (height !== undefined) page.height = height;
  }
  if (args.orientation === "portrait" || args.orientation === "landscape") {
    page.orientation = args.orientation;
  }
  const margins = marginPatch(args.margins);
  if (margins) {
    page.margins = margins;
  }
  return { sectionIndex: integer(args.sectionIndex, 0), page };
}

/** Normalizes the image payload, including a `data:` URL, before the engine. */
function imageArgs(args, workspaceDir) {
  let dataBase64 =
    typeof args.dataBase64 === "string" ? args.dataBase64 : undefined;
  let contentType = optionalString(args.contentType);
  if (dataBase64) {
    const match = /^data:([^;,]+);base64,(.*)$/s.exec(dataBase64);
    if (match) {
      contentType = contentType ?? match[1];
      dataBase64 = match[2];
    }
  }
  const sourcePath = workspacePath(workspaceDir, args.path);
  return {
    position: args.position,
    path: sourcePath,
    dataBase64,
    contentType,
    width: optionalNumber(args.width),
    height: optionalNumber(args.height),
    alt: optionalString(args.alt),
    floating: args.floating === true,
  };
}

/**
 * The editor's op table. Each entry maps one contract op onto a
 * `DocumentService` method; nothing here reimplements engine behaviour.
 */
const OP_ACTIONS = {
  insertText: {
    keys: ["position"],
    run: (service, id, args) =>
      service.insertText(id, {
        position: args.position,
        text: typeof args.text === "string" ? args.text : "",
        props: asObject(args.props),
      }),
  },
  replaceText: {
    keys: ["range"],
    run: (service, id, args) =>
      service.replaceText(id, {
        range: args.range,
        text: typeof args.text === "string" ? args.text : "",
        props: asObject(args.props),
      }),
  },
  deleteRange: {
    keys: ["range"],
    run: (service, id, args) => service.deleteRange(id, { range: args.range }),
  },
  formatText: {
    keys: ["range", "props"],
    run: (service, id, args) =>
      service.formatText(id, {
        range: args.range,
        props: asObject(args.props) ?? {},
        clear: args.clear === true,
      }),
  },
  formatParagraph: {
    keys: ["blockIds", "props"],
    run: (service, id, args) =>
      service.formatParagraph(id, {
        blockIds: stringArray(args.blockIds, "blockIds"),
        props: asObject(args.props) ?? {},
      }),
  },
  applyStyle: {
    keys: ["blockIds", "styleId"],
    run: (service, id, args) =>
      service.applyStyle(id, {
        blockIds: stringArray(args.blockIds, "blockIds"),
        styleId: String(args.styleId),
        clearDirectFormatting: args.clearDirectFormatting === true,
      }),
  },
  insertHeading: {
    keys: ["text", "level"],
    run: (service, id, args) =>
      service.insertHeading(id, {
        text: String(args.text),
        level: headingLevel(args.level),
        afterBlockId: optionalString(args.afterBlockId),
        styleId: optionalString(args.styleId),
      }),
  },
  insertList: {
    keys: ["items", "kind"],
    run: (service, id, args) =>
      service.insertList(id, {
        items: stringArray(args.items, "items"),
        kind: listKind(args.kind),
        afterBlockId: optionalString(args.afterBlockId),
        level: optionalInteger(args.level),
        startAt: optionalInteger(args.startAt),
      }),
  },
  insertTable: {
    keys: ["rows", "columns"],
    run: (service, id, args) =>
      service.insertTable(id, {
        rows: positiveInteger(args.rows, "rows"),
        columns: positiveInteger(args.columns, "columns"),
        data: tableData(args.data),
        style: optionalString(args.style),
        afterBlockId: optionalString(args.afterBlockId),
        headerRow: args.headerRow === true,
      }),
  },
  editTable: {
    keys: ["tableId", "operation"],
    run: (service, id, args) =>
      service.editTable(id, {
        tableId: String(args.tableId),
        operation: String(args.operation),
        index: optionalInteger(args.index),
        region: asObject(args.region),
        cellId: optionalString(args.cellId),
      }),
  },
  insertImage: {
    keys: ["position"],
    run: (service, id, args, workspaceDir) =>
      service.insertImage(id, imageArgs(args, workspaceDir)),
  },
  insertHyperlink: {
    keys: ["range"],
    run: (service, id, args) => {
      const href = optionalString(args.href);
      const anchor = optionalString(args.anchor);
      if (!href && !anchor) {
        throw new HttpError(
          "INVALID_ARGUMENT",
          "insertHyperlink requires args.href or args.anchor.",
        );
      }
      return service.insertHyperlink(id, {
        range: args.range,
        href,
        anchor,
        tooltip: optionalString(args.tooltip),
      });
    },
  },
  insertPageBreak: {
    keys: ["position"],
    run: (service, id, args) =>
      service.insertPageBreak(id, {
        position: args.position,
        kind: BREAK_KINDS.has(args.kind) ? args.kind : "page",
      }),
  },
  insertSectionBreak: {
    run: (service, id, args) =>
      service.insertSectionBreak(id, {
        sectionIndex: optionalInteger(args.sectionIndex),
        blockIndex: optionalInteger(args.blockIndex),
        type: SECTION_BREAK_TYPES.has(args.type) ? args.type : undefined,
        landscape:
          typeof args.landscape === "boolean" ? args.landscape : undefined,
      }),
  },
  insertFootnote: {
    keys: ["position", "text"],
    run: (service, id, args) =>
      service.addFootnote(id, {
        position: args.position,
        text: String(args.text),
        type: args.type === "endnote" ? "endnote" : "footnote",
      }),
  },
  insertToc: {
    run: (service, id, args) =>
      service.insertToc(id, {
        afterBlockId: optionalString(args.afterBlockId),
        minLevel: optionalInteger(args.minLevel),
        maxLevel: optionalInteger(args.maxLevel),
        includeTitle:
          typeof args.includeTitle === "boolean"
            ? args.includeTitle
            : undefined,
      }),
  },
  insertEquation: {
    keys: ["position", "linear"],
    run: (service, id, args) =>
      service.insertEquation(id, {
        position: args.position,
        linear: String(args.linear),
        display: args.display === true,
      }),
  },
  setPageLayout: {
    keys: ["sectionIndex"],
    run: (service, id, args) => service.setPageLayout(id, pageLayoutArgs(args)),
  },
  setColumns: {
    keys: ["sectionIndex", "count"],
    run: (service, id, args) =>
      service.setPageLayout(id, {
        sectionIndex: integer(args.sectionIndex, 0),
        columns: {
          count: positiveInteger(args.count, "count"),
          space: optionalNumber(args.space) ?? 36,
        },
      }),
  },
  editHeaderFooter: {
    keys: ["part"],
    run: (service, id, args) => {
      if (!HEADER_FOOTER_PARTS.has(args.part)) {
        throw new HttpError(
          "INVALID_ARGUMENT",
          'editHeaderFooter requires args.part "header" or "footer".',
        );
      }
      return service.editHeaderFooter(id, {
        sectionIndex: optionalInteger(args.sectionIndex),
        part: args.part,
        variant: HEADER_FOOTER_VARIANTS.has(args.variant)
          ? args.variant
          : undefined,
        text: typeof args.text === "string" ? args.text : undefined,
        blocks: Array.isArray(args.blocks) ? args.blocks : undefined,
      });
    },
  },
  addComment: {
    keys: ["range", "text"],
    run: (service, id, args) =>
      service.addComment(id, {
        range: args.range,
        text: String(args.text),
        author: optionalString(args.author),
      }),
  },
  setMetadata: {
    keys: ["patch"],
    run: (service, id, args) =>
      service.setDocumentMetadata(id, asObject(args.patch) ?? {}),
  },
  trackChanges: {
    keys: ["enabled"],
    run: (service, id, args) => {
      if (typeof args.enabled !== "boolean") {
        throw new HttpError(
          "INVALID_ARGUMENT",
          "trackChanges requires a boolean args.enabled.",
        );
      }
      return service.setTrackChanges(id, args.enabled);
    },
  },
  acceptChanges: {
    run: (service, id, args) =>
      service.acceptChanges(id, optionalString(args.revisionId)),
  },
  rejectChanges: {
    run: (service, id, args) =>
      service.rejectChanges(id, optionalString(args.revisionId)),
  },
  undo: { run: (service, id) => service.undo(id) },
  redo: { run: (service, id) => service.redo(id) },
};

const OP_NAMES = Object.keys(OP_ACTIONS).sort();
/** Ops that replace the whole model, so the caller also gets the document. */
const HISTORY_OPS = new Set(["undo", "redo"]);

function requireOperation(value) {
  const name = typeof value === "string" ? value : "";
  if (!name || !Object.hasOwn(OP_ACTIONS, name)) {
    throw new HttpError(
      "UNKNOWN_OPERATION",
      `Unknown op ${JSON.stringify(value ?? null)}. Allowed ops: ${OP_NAMES.join(", ")}.`,
    );
  }
  return name;
}

function applyOperation(service, documentId, name, args, workspaceDir) {
  const action = OP_ACTIONS[name];
  if (args !== undefined && args !== null && !asObject(args)) {
    throw new HttpError("INVALID_ARGUMENT", `${name} args must be an object.`);
  }
  const input = asObject(args) ?? {};
  if (action.keys) requireArgs(name, input, action.keys);
  return action.run(service, documentId, input, workspaceDir);
}

/** Small, content-free summary returned with every mutating response. */
function documentSnapshot(service, documentId) {
  const session = service.session(documentId);
  const stats = service.metadata(documentId).stats;
  const pageCount = service.layout(session).pageCount;
  return {
    revision: session.revision,
    dirty: session.dirty,
    title: session.document.metadata.title ?? "Untitled document",
    stats,
    pageCount,
  };
}

/**
 * 1-based page for every body block, skipping header/footer/note regions.
 * Blocks the layout engine did not mark inherit the page of the block before
 * them, so the page can always draw a separator for a body block.
 */
function blockPages(document, layout) {
  const pages = layout.pageNumberByBlock;
  const out = {};
  let current = 1;
  const visit = (blocks) => {
    for (const block of blocks ?? []) {
      const page = pages.get(block.id);
      if (typeof page === "number") current = page;
      out[block.id] = current;
      if (block.type === "table") {
        for (const row of block.rows ?? []) {
          for (const cell of row.cells ?? []) visit(cell.blocks);
        }
      }
    }
  };
  for (const section of document.sections ?? []) visit(section.blocks);
  return out;
}

function capabilities() {
  let ai = false;
  try {
    ai = providerStatus().available === true;
  } catch {
    ai = false;
  }
  return { ai, formats: true, convert: true, renderer: true };
}

export function apply(ctx, config = {}) {
  const initialWorkspace = resolveWorkspace(config);
  const readOnly =
    config?.readOnly === true || process.env.DSH_DOC_READ_ONLY === "1";
  const toolNames =
    Array.isArray(config?.tools) && config.tools.length > 0
      ? config.tools
      : DEFAULT_TOOLS;
  const cwdRoot = path.resolve(process.cwd());
  const homeRoot = path.resolve(os.homedir());
  /** Every filesystem root this Host can read, resolved once per activation. */
  const readableRoots = filesystemRoots();

  /**
   * Engine options for one document workspace root. The service is given every
   * local read root (all drives, or `/`) so the panel and the agent tools can
   * import a document from any folder the Host can read. The engine still
   * confines *writes* — save, saveAs and delete stay inside `workspaceRoot`.
   */
  const engineOptions = (root) => ({
    workspaceRoot: root,
    allowedReadRoots: Array.from(
      new Set([root, cwdRoot, homeRoot, ...readableRoots]),
    ),
    readOnly,
    auditPath: path.join(root, "audit.log"),
  });

  /**
   * The active document workspace. `POST /workspace` replaces it with a new
   * service; the previous one is retired (not closed) so an open dirty
   * document keeps its model and stays readable for the rest of the process.
   */
  let workspaceDir = initialWorkspace;
  let service = new DocumentService(engineOptions(workspaceDir));
  let ready = service.init().then(async () => {
    await mkdir(workspaceDir, { recursive: true });
    return service;
  });
  const retired = [];

  /** Awaits initialization and returns the service routes should use now. */
  async function active() {
    await ready;
    return service;
  }

  /** The service that holds `documentId`: the active one, then retired ones. */
  function serviceFor(documentId) {
    try {
      service.session(documentId);
      return service;
    } catch {
      // Not in the active workspace; look at the ones it replaced.
    }
    for (const previous of retired) {
      try {
        previous.session(documentId);
        return previous;
      } catch {
        // Not open there either.
      }
    }
    return service;
  }

  function allServices() {
    return [service, ...retired];
  }

  /** Importable extensions reported by this build's format registry. */
  let importExtensions;
  function importableExtensions(svc) {
    if (!importExtensions) {
      importExtensions = svc
        .listFormats()
        .then((listed) => {
          const extensions = Array.isArray(listed?.import) ? listed.import : [];
          return new Set(extensions.map((ext) => String(ext).toLowerCase()));
        })
        .catch(() => {
          importExtensions = undefined;
          return new Set(["docx", "dotx", "docm"]);
        });
    }
    return importExtensions;
  }

  /** Side-list rows for every open document, including retired workspaces. */
  function listDocuments() {
    const rows = [];
    for (const svc of allServices()) {
      for (const session of svc.sessions_()) {
        let words;
        try {
          words = svc.metadata(session.id).stats?.words;
        } catch {
          words = undefined;
        }
        rows.push({
          id: session.id,
          title: session.document.metadata.title ?? "Untitled document",
          dirty: session.dirty,
          format: session.sourceFormat ?? "docx",
          words,
          modified: session.document.metadata.modified,
          sections: session.document.sections.length,
          blocks: session.document.sections.reduce(
            (sum, section) => sum + section.blocks.length,
            0,
          ),
        });
      }
    }
    return rows;
  }

  const manifest = buildToolManifest();
  const byName = new Map(manifest.tools.map((tool) => [tool.name, tool]));

  for (const name of toolNames) {
    const entry = byName.get(name);
    if (!entry) {
      ctx.logger?.warn?.(`dsh-documents: unknown tool ${name} was skipped`);
      continue;
    }
    const schema = { ...entry.inputSchema };
    delete schema.$schema;
    ctx.effect(() =>
      ctx.tools.register({
        name,
        description: entry.description,
        parameters: schema,
        output: {
          schema: { type: "object", additionalProperties: true },
          render(_args, value) {
            return [{ type: "text", text: JSON.stringify(value, null, 2) }];
          },
        },
        timeoutMs: 120_000,
        async execute(args) {
          await ready;
          const input = args ?? {};
          // A document opened before a workspace switch lives in a retired
          // service; tools must keep reaching it.
          const documentId =
            typeof input.documentId === "string" ? input.documentId : "";
          const target = documentId ? serviceFor(documentId) : await active();
          const result = await invokeTool(target, name, input);
          if (result.ok === false) {
            const detail = result.error;
            throw new Error(
              `${detail?.code ?? "DOCUMENT_ERROR"}: ${detail?.message ?? "Document operation failed"}`,
            );
          }
          return result.data ?? null;
        },
      }),
    );
  }

  /**
   * The Web carrier may not be listening when this plugin activates during a
   * Harness boot, so route registration waits for the service instead of
   * returning early (an early return left every editor route at 404 until the
   * next hot reload).
   */
  const attachRoutes = (routeCtx, webServer) => {
    routeCtx.effect(() => {
      const disposers = [];

      disposers.push(
        webServer.register({
          kind: "prefix",
          path: API_PREFIX,
          async handler(req, res) {
            try {
              if (!originAllowed(req, webServer.getOrigin())) {
                req.resume?.();
                json(res, 403, {
                  ok: false,
                  error: {
                    code: "ORIGIN_REFUSED",
                    message: "Cross-origin request refused.",
                  },
                });
                return;
              }
              if (req.headers[REQUEST_HEADER] !== "1") {
                req.resume?.();
                json(res, 403, {
                  ok: false,
                  error: {
                    code: "HEADER_REQUIRED",
                    message: "Missing the Documents page request header.",
                  },
                });
                return;
              }

              const method = (req.method ?? "GET").toUpperCase();
              const url = new URL(req.url ?? "/", "http://localhost");
              const route = url.pathname.slice(API_PREFIX.length) || "/";

              // Read-only mode never reaches the service for a mutating route.
              if (method === "POST" && readOnly && MUTATING_ROUTES.has(route)) {
                req.resume?.();
                json(res, 200, {
                  ok: false,
                  error: {
                    code: "READ_ONLY_MODE",
                    message:
                      "The Documents panel is read-only (DSH_DOC_READ_ONLY=1).",
                  },
                });
                return;
              }

              await ready;
              // Snapshot the active workspace/service so one request always talks
              // to the same engine even if another request switches workspaces.
              const activeRoot = workspaceDir;
              const svc = service;

              let body;
              let upload;
              if (method === "POST") {
                const boundary = multipartBoundary(req);
                if (boundary && route === "/open") {
                  // Raw upload: the Host saves the copy inside the workspace.
                  upload = firstFilePart(
                    await readBodyBuffer(req, LARGE_BODY_LIMIT),
                    boundary,
                  );
                  if (!upload) {
                    throw new HttpError(
                      "INVALID_ARGUMENT",
                      "The upload had no file part.",
                    );
                  }
                  body = {};
                } else {
                  body = await readJsonBody(
                    req,
                    LARGE_BODY_ROUTES.has(route)
                      ? LARGE_BODY_LIMIT
                      : BODY_LIMIT,
                  );
                }
              }

              if (method === "GET" && route === "/browse") {
                const roots = browseRoots(activeRoot);
                const requested = url.searchParams.get("path");
                if (requested === null || requested.trim() === "") {
                  json(res, 200, { ok: true, roots });
                  return;
                }
                // Any absolute local directory the Host can realpath: another
                // drive, a public/shared folder, or a project outside the
                // session working directory. `realpath` resolves `..` before
                // the listing, so a traversal only names a real directory.
                const target = await requireBrowsableDirectory(requested);
                const importable = await importableExtensions(svc);
                const listing = await browseDirectory(target, importable);
                json(res, 200, {
                  ok: true,
                  path: target,
                  parent: browseParent(target),
                  roots,
                  entries: listing.entries,
                  truncated: listing.truncated,
                  skipped: listing.skipped,
                });
                return;
              }

              if (method === "GET" && route === "/workspace") {
                json(res, 200, {
                  ok: true,
                  path: activeRoot,
                  readOnly,
                  documents: listDocuments(),
                });
                return;
              }

              if (method === "POST" && route === "/workspace") {
                const create = body.create === true;
                // Creating a folder is a filesystem write, so read-only refuses
                // it. Switching to an existing folder is not a mutation.
                if (create && readOnly) {
                  throw new HttpError(
                    "READ_ONLY_MODE",
                    "The Documents panel is read-only (DSH_DOC_READ_ONLY=1).",
                  );
                }
                // Any absolute local directory, not just Home/workspace/cwd.
                // A filesystem root is refused as a workspace target.
                const targetDirectory = await requireWorkspaceDirectory(
                  body.path,
                  { create },
                );
                const target = targetDirectory.path;
                const created = targetDirectory.created;

                if (!samePath(target, workspaceDir)) {
                  // Build and initialize the replacement before switching so a
                  // failed init cannot leave the panel without a service.
                  const next = new DocumentService(engineOptions(target));
                  await next.init();
                  retired.push(service);
                  service = next;
                  workspaceDir = target;
                  ready = Promise.resolve(next);
                }

                let dshWorkspaceId;
                let warning;
                const registry = ctx.get("workspaceRegistry");
                if (
                  registry &&
                  (typeof registry.resolveByPath === "function" ||
                    typeof registry.create === "function")
                ) {
                  try {
                    let workspace =
                      typeof registry.resolveByPath === "function"
                        ? await registry.resolveByPath(target)
                        : undefined;
                    if (!workspace && typeof registry.create === "function") {
                      workspace = await registry.create(
                        target,
                        path.basename(target) || target,
                      );
                    }
                    if (workspace && typeof workspace.id === "string") {
                      dshWorkspaceId = workspace.id;
                    }
                  } catch {
                    warning = "workspace-registry-failed";
                  }
                } else {
                  warning = "workspace-registry-unavailable";
                }

                json(res, 200, {
                  ok: true,
                  path: target,
                  created,
                  ...(dshWorkspaceId ? { dshWorkspaceId } : {}),
                  ...(warning ? { warning } : {}),
                });
                return;
              }

              if (method === "POST" && route === "/open-path") {
                // Any absolute local file the Host can read; the importable
                // extension gate below is the only format restriction.
                const target = await requireOpenFile(body.path);
                const importable = await importableExtensions(svc);
                if (!importable.has(extensionOf(target))) {
                  throw new HttpError(
                    "FORMAT_UNSUPPORTED",
                    "That file format cannot be imported by this build.",
                  );
                }
                // Re-opening a file that is already open reuses its session, so the
                // document list never shows the same file twice.
                const alreadyOpen = svc
                  .sessions_()
                  .find(
                    (session) =>
                      session.filePath !== undefined &&
                      samePath(session.filePath, target),
                  );
                if (alreadyOpen) {
                  json(res, 200, {
                    ok: true,
                    documentId: alreadyOpen.id,
                    path: target,
                    directory: path.dirname(target),
                    title:
                      alreadyOpen.document.metadata.title ??
                      path.basename(target).replace(/\.[^.]+$/, ""),
                    warnings: [],
                    reused: true,
                  });
                  return;
                }
                const opened = await svc.openDocument(target);
                const fallbackTitle = path
                  .basename(target)
                  .replace(/\.[^.]+$/, "");
                if (!opened.document?.metadata?.title) {
                  // Give imported files their file name as the document title.
                  try {
                    svc.setDocumentMetadata(opened.documentId, {
                      title: fallbackTitle,
                    });
                  } catch {
                    // The title is cosmetic; never fail the open for it.
                  }
                }
                json(res, 200, {
                  ok: true,
                  documentId: opened.documentId,
                  path: target,
                  directory: path.dirname(target),
                  title:
                    svc.session(opened.documentId).document.metadata.title ??
                    fallbackTitle,
                  warnings: opened.warnings ?? [],
                });
                return;
              }

              if (method === "GET" && route === "/list") {
                json(res, 200, {
                  ok: true,
                  workspace: activeRoot,
                  readOnly,
                  documents: listDocuments(),
                });
                return;
              }

              if (method === "GET" && route === "/bootstrap") {
                const documentId = safeId(url.searchParams.get("id"));
                const owner = serviceFor(documentId);
                const session = owner.session(documentId);
                const layout = owner.layout(session);
                const meta = owner.metadata(documentId);
                let formats = { import: [], export: [] };
                try {
                  const listed = await owner.listFormats();
                  formats = {
                    import: Array.isArray(listed?.import) ? listed.import : [],
                    export: Array.isArray(listed?.export) ? listed.export : [],
                  };
                } catch {
                  // The model still edits fine without the adapter registry.
                }
                json(res, 200, {
                  ok: true,
                  workspace: activeRoot,
                  readOnly,
                  formats,
                  documentId,
                  title: session.document.metadata.title ?? "Untitled document",
                  dirty: session.dirty,
                  revision: session.revision,
                  sourceFormat: session.sourceFormat ?? "docx",
                  document: session.document,
                  blockPages: blockPages(session.document, layout),
                  pageCount: layout.pageCount,
                  stats: meta.stats,
                  outline: owner.getOutline(documentId),
                  capabilities: capabilities(),
                });
                return;
              }

              if (method === "GET" && route === "/search") {
                const documentId = safeId(url.searchParams.get("id"));
                const query = url.searchParams.get("q") ?? "";
                if (!query) {
                  throw new HttpError(
                    "INVALID_ARGUMENT",
                    "The q search parameter is required.",
                  );
                }
                const found = serviceFor(documentId).search(documentId, query, {
                  regex: queryFlag(url.searchParams, "regex"),
                  caseSensitive: queryFlag(url.searchParams, "caseSensitive"),
                  includePages: true,
                });
                json(res, 200, {
                  ok: true,
                  matches: found.matches,
                  total: found.total,
                });
                return;
              }

              if (method === "POST" && route === "/op") {
                const documentId = safeId(body.id);
                const owner = serviceFor(documentId);
                try {
                  const name = requireOperation(body.op);
                  const result = await applyOperation(
                    owner,
                    documentId,
                    name,
                    body.args,
                    owner.workspace.root,
                  );
                  if (HISTORY_OPS.has(name)) {
                    json(res, 200, {
                      ok: true,
                      ...documentSnapshot(owner, documentId),
                      document: owner.session(documentId).document,
                    });
                    return;
                  }
                  json(res, 200, {
                    ok: true,
                    ...documentSnapshot(owner, documentId),
                    result: result ?? null,
                  });
                } catch (error) {
                  json(res, 200, {
                    ok: false,
                    applied: 0,
                    error: errorPayload(error),
                  });
                }
                return;
              }

              if (method === "POST" && route === "/ops") {
                const documentId = safeId(body.id);
                const owner = serviceFor(documentId);
                if (!Array.isArray(body.ops) || body.ops.length === 0) {
                  throw new HttpError(
                    "INVALID_ARGUMENT",
                    "ops must be a non-empty array.",
                  );
                }
                let applied = 0;
                let result;
                let history = false;
                for (const entry of body.ops) {
                  try {
                    if (!asObject(entry)) {
                      throw new HttpError(
                        "INVALID_ARGUMENT",
                        "Each ops entry must be an object.",
                      );
                    }
                    const name = requireOperation(entry.op);
                    result = await applyOperation(
                      owner,
                      documentId,
                      name,
                      entry.args,
                      owner.workspace.root,
                    );
                    if (HISTORY_OPS.has(name)) history = true;
                    applied += 1;
                  } catch (error) {
                    json(res, 200, {
                      ok: false,
                      applied,
                      error: errorPayload(error),
                    });
                    return;
                  }
                }
                json(res, 200, {
                  ok: true,
                  applied,
                  ...documentSnapshot(owner, documentId),
                  result: result ?? null,
                  ...(history
                    ? { document: owner.session(documentId).document }
                    : {}),
                });
                return;
              }

              if (method === "POST" && route === "/save") {
                const documentId = safeId(body.id);
                const owner = serviceFor(documentId);
                const target = workspacePath(owner.workspace.root, body.path);
                const saved = await owner.saveDocument(documentId, target);
                json(res, 200, {
                  ok: true,
                  path: saved.path,
                  bytes: saved.bytes,
                  revision: owner.session(documentId).revision,
                });
                return;
              }

              if (method === "POST" && route === "/saveAs") {
                const documentId = safeId(body.id);
                const owner = serviceFor(documentId);
                const target = workspacePath(owner.workspace.root, body.path);
                if (!target) {
                  throw new HttpError(
                    "PATH_REQUIRED",
                    "Give a path for the saved copy.",
                  );
                }
                const saved = await owner.saveDocumentAs(documentId, target);
                json(res, 200, {
                  ok: true,
                  path: saved.path,
                  bytes: saved.bytes,
                  revision: owner.session(documentId).revision,
                });
                return;
              }

              if (method === "POST" && route === "/create") {
                const created = await svc.createDocument({
                  title:
                    typeof body.title === "string" && body.title.trim()
                      ? body.title.trim()
                      : "Untitled document",
                  paragraphs:
                    typeof body.text === "string" && body.text.trim()
                      ? [body.text.trim()]
                      : undefined,
                });
                json(res, 200, { ok: true, documentId: created.documentId });
                return;
              }

              if (method === "POST" && route === "/open") {
                // Two shapes: an uploaded file (multipart, saved into the
                // workspace) and the original `{ path }` JSON form.
                let sourcePath;
                if (upload) {
                  if (readOnly) {
                    throw new HttpError(
                      "READ_ONLY_MODE",
                      "The Documents panel is read-only (DSH_DOC_READ_ONLY=1).",
                    );
                  }
                  const name = safeUploadName(upload.filename);
                  if (
                    !(await importableExtensions(svc)).has(extensionOf(name))
                  ) {
                    throw new HttpError(
                      "FORMAT_UNSUPPORTED",
                      "That file format cannot be imported by this build.",
                    );
                  }
                  const directory = path.join(activeRoot, "imports");
                  await mkdir(directory, { recursive: true });
                  sourcePath = path.join(
                    directory,
                    `${Date.now().toString(36)}-${name}`,
                  );
                  await writeFile(sourcePath, upload.data);
                } else {
                  if (typeof body.path !== "string" || !body.path.trim()) {
                    throw new HttpError(
                      "PATH_REQUIRED",
                      "Give a path to a document inside the workspace.",
                    );
                  }
                  sourcePath = workspacePath(activeRoot, body.path);
                }
                const opened = await svc.openDocument(sourcePath);
                json(res, 200, {
                  ok: true,
                  documentId: opened.documentId,
                  path: path.resolve(sourcePath),
                  directory: path.dirname(path.resolve(sourcePath)),
                  title:
                    svc.session(opened.documentId).document.metadata.title ??
                    path.basename(sourcePath).replace(/\.[^.]+$/, ""),
                  warnings: opened.warnings ?? [],
                  report: opened.report,
                  compatibility: opened.compatibility,
                });
                return;
              }

              if (method === "GET" && route === "/document") {
                const documentId = safeId(url.searchParams.get("id"));
                const owner = serviceFor(documentId);
                const session = owner.session(documentId);
                const metadata = owner.metadata(documentId);
                const outline = owner.getOutline(documentId);
                const text = owner.getText(documentId).text;
                json(res, 200, {
                  ok: true,
                  metadata,
                  outline,
                  text:
                    text.length > 6000
                      ? `${text.slice(0, 6000)}\n\n[Text truncated by the Documents panel]`
                      : text,
                  truncated: text.length > 6000,
                  layout: owner.getPageInfo(documentId),
                  sourceFormat: session.sourceFormat ?? "docx",
                });
                return;
              }

              if (method === "GET" && route === "/export") {
                const documentId = safeId(url.searchParams.get("id"));
                const format = String(
                  url.searchParams.get("format") ?? "docx",
                ).toLowerCase();
                const allowed = new Set([
                  "docx",
                  "dotx",
                  "docm",
                  "pdf",
                  "html",
                  "md",
                  "txt",
                  "odt",
                  "rtf",
                ]);
                if (!allowed.has(format)) {
                  json(res, 400, {
                    ok: false,
                    error: {
                      code: "FORMAT_UNSUPPORTED",
                      message: `Cannot export ${format}.`,
                    },
                  });
                  return;
                }
                const owner = serviceFor(documentId);
                const exported =
                  format === "pdf"
                    ? await owner.exportPdf(documentId)
                    : await owner.exportDocument(documentId, format);
                const bytes = Buffer.from(exported.dataBase64 ?? "", "base64");
                const types = {
                  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  dotx: "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
                  docm: "application/vnd.ms-word.document.macroEnabled.12",
                  pdf: "application/pdf",
                  html: "text/html; charset=utf-8",
                  md: "text/markdown; charset=utf-8",
                  txt: "text/plain; charset=utf-8",
                  odt: "application/vnd.oasis.opendocument.text",
                  rtf: "application/rtf",
                };
                const title = (
                  owner.metadata(documentId).metadata?.title ?? documentId
                )
                  .replace(/[^\w .-]+/g, "_")
                  .slice(0, 80);
                res.writeHead(200, {
                  "content-type": types[format] ?? "application/octet-stream",
                  "content-length": String(bytes.length),
                  "content-disposition": `attachment; filename="${title || "document"}.${format}"`,
                  "cache-control": "no-store",
                });
                res.end(bytes);
                return;
              }

              if (method === "POST" && route === "/delete") {
                const documentId = safeId(body.id);
                const owner = serviceFor(documentId);
                await owner
                  .closeDocument(documentId, { discardUnsaved: true })
                  .catch(() => undefined);
                await owner.deleteDocument(documentId);
                json(res, 200, { ok: true });
                return;
              }

              json(res, 404, {
                ok: false,
                error: {
                  code: "NOT_FOUND",
                  message: "Unknown Documents API route.",
                },
              });
            } catch (error) {
              json(res, 200, {
                ok: false,
                error: errorPayload(error),
              });
            }
          },
        }),
      );

      return () => {
        for (const dispose of disposers.reverse()) {
          try {
            dispose();
          } catch {
            // Disposal is best-effort during unload.
          }
        }
      };
    });
  };

  const immediateWebServer = ctx.get("webServer");
  if (immediateWebServer) {
    attachRoutes(ctx, immediateWebServer);
  } else if (typeof ctx.inject === "function") {
    ctx.inject(["webServer"], (routeCtx) => {
      const webServer = routeCtx.get("webServer") ?? routeCtx.webServer;
      if (webServer) attachRoutes(routeCtx, webServer);
    });
  } else {
    // Last-resort readiness wait for hosts without `ctx.inject`. The timer is
    // unref'd so a mock or headless host can still exit.
    const timer = setInterval(() => {
      const webServer = ctx.get("webServer");
      if (!webServer) return;
      clearInterval(timer);
      attachRoutes(ctx, webServer);
    }, 250);
    timer.unref?.();
    ctx.effect(() => () => clearInterval(timer));
  }
}
