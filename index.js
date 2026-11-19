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
import { mkdir } from "node:fs/promises";

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
  const workspaceDir = resolveWorkspace(config);
  const readOnly =
    config?.readOnly === true || process.env.DSH_DOC_READ_ONLY === "1";
  const toolNames =
    Array.isArray(config?.tools) && config.tools.length > 0
      ? config.tools
      : DEFAULT_TOOLS;
  const extraRoot = path.resolve(process.cwd());
  const allowedReadRoots = Array.from(new Set([workspaceDir, extraRoot]));

  const service = new DocumentService({
    workspaceRoot: workspaceDir,
    allowedReadRoots,
    readOnly,
    auditPath: path.join(workspaceDir, "audit.log"),
  });
  const ready = service.init().then(async () => {
    await mkdir(workspaceDir, { recursive: true });
    return service;
  });

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
          const result = await invokeTool(service, name, args ?? {});
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

  const webServer = ctx.get("webServer");
  if (!webServer) return;

  ctx.effect(() => {
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
            const body =
              method === "POST"
                ? await readJsonBody(
                    req,
                    LARGE_BODY_ROUTES.has(route)
                      ? LARGE_BODY_LIMIT
                      : BODY_LIMIT,
                  )
                : undefined;

            if (method === "GET" && route === "/list") {
              const documents = service.sessions_().map((session) => {
                let words;
                try {
                  words = service.metadata(session.id).stats?.words;
                } catch {
                  words = undefined;
                }
                return {
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
                };
              });
              json(res, 200, {
                ok: true,
                workspace: workspaceDir,
                readOnly,
                documents,
              });
              return;
            }

            if (method === "GET" && route === "/bootstrap") {
              const documentId = safeId(url.searchParams.get("id"));
              const session = service.session(documentId);
              const layout = service.layout(session);
              const meta = service.metadata(documentId);
              let formats = { import: [], export: [] };
              try {
                const listed = await service.listFormats();
                formats = {
                  import: Array.isArray(listed?.import) ? listed.import : [],
                  export: Array.isArray(listed?.export) ? listed.export : [],
                };
              } catch {
                // The model still edits fine without the adapter registry.
              }
              json(res, 200, {
                ok: true,
                workspace: workspaceDir,
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
                outline: service.getOutline(documentId),
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
              const found = service.search(documentId, query, {
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
              try {
                const name = requireOperation(body.op);
                const result = await applyOperation(
                  service,
                  documentId,
                  name,
                  body.args,
                  workspaceDir,
                );
                if (HISTORY_OPS.has(name)) {
                  json(res, 200, {
                    ok: true,
                    ...documentSnapshot(service, documentId),
                    document: service.session(documentId).document,
                  });
                  return;
                }
                json(res, 200, {
                  ok: true,
                  ...documentSnapshot(service, documentId),
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
                    service,
                    documentId,
                    name,
                    entry.args,
                    workspaceDir,
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
                ...documentSnapshot(service, documentId),
                result: result ?? null,
                ...(history
                  ? { document: service.session(documentId).document }
                  : {}),
              });
              return;
            }

            if (method === "POST" && route === "/save") {
              const documentId = safeId(body.id);
              const target = workspacePath(workspaceDir, body.path);
              const saved = await service.saveDocument(documentId, target);
              json(res, 200, {
                ok: true,
                path: saved.path,
                bytes: saved.bytes,
                revision: service.session(documentId).revision,
              });
              return;
            }

            if (method === "POST" && route === "/saveAs") {
              const documentId = safeId(body.id);
              const target = workspacePath(workspaceDir, body.path);
              if (!target) {
                throw new HttpError(
                  "PATH_REQUIRED",
                  "Give a path for the saved copy.",
                );
              }
              const saved = await service.saveDocumentAs(documentId, target);
              json(res, 200, {
                ok: true,
                path: saved.path,
                bytes: saved.bytes,
                revision: service.session(documentId).revision,
              });
              return;
            }

            if (method === "POST" && route === "/create") {
              const created = await service.createDocument({
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
              if (typeof body.path !== "string" || !body.path.trim()) {
                throw new HttpError(
                  "PATH_REQUIRED",
                  "Give a path to a document inside the workspace.",
                );
              }
              const opened = await service.openDocument(body.path.trim());
              json(res, 200, {
                ok: true,
                documentId: opened.documentId,
                warnings: opened.warnings,
                report: opened.report,
                compatibility: opened.compatibility,
              });
              return;
            }

            if (method === "GET" && route === "/document") {
              const documentId = safeId(url.searchParams.get("id"));
              const session = service.session(documentId);
              const metadata = service.metadata(documentId);
              const outline = service.getOutline(documentId);
              const text = service.getText(documentId).text;
              json(res, 200, {
                ok: true,
                metadata,
                outline,
                text:
                  text.length > 6000
                    ? `${text.slice(0, 6000)}\n\n[Text truncated by the Documents panel]`
                    : text,
                truncated: text.length > 6000,
                layout: service.getPageInfo(documentId),
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
              const exported =
                format === "pdf"
                  ? await service.exportPdf(documentId)
                  : await service.exportDocument(documentId, format);
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
                service.metadata(documentId).metadata?.title ?? documentId
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
              await service
                .closeDocument(documentId, { discardUnsaved: true })
                .catch(() => undefined);
              await service.deleteDocument(documentId);
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
}
