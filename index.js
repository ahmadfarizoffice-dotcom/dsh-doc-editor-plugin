/**
 * Host half of the DSH Documents bundle.
 *
 * Registers document tools on `ctx.tools` and, when the Web server is present,
 * a small same-origin API used by the Documents page in the main column. All
 * work goes through the bundled engine (`./engine.js`), which is the same
 * document service the standalone editor and MCP server use.
 */
import os from "node:os";
import path from "node:path";
import { mkdir } from "node:fs/promises";
import { DocumentService, invokeTool, buildToolManifest } from "./engine.js";

export const inject = ["tools"];

const PANEL_ID = "dsh-documents";
const API_PREFIX = "/api/dsh-documents";
const REQUEST_HEADER = "x-dsh-documents";

/** Tools exposed to the agent. Kept small enough to stay cheap in context. */
const DEFAULT_TOOLS = [
  "document_list_formats",
  "document_create",
  "document_open",
  "document_get_metadata",
  "document_get_outline",
  "document_get_text",
  "document_search",
  "document_insert_text",
  "document_replace_text",
  "document_format_text",
  "document_apply_style",
  "document_insert_table",
  "document_ai_review",
  "document_ai_summarize",
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

async function readJsonBody(req, limit = 1024 * 1024) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > limit) throw new Error("Request body is too large");
    chunks.push(chunk);
  }
  if (total === 0) return {};
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

function safeId(value) {
  const id = String(value ?? "");
  if (!/^[A-Za-z0-9_-]{1,64}$/.test(id)) throw new Error("Invalid document id");
  return id;
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
              json(res, 403, {
                ok: false,
                error: {
                  code: "HEADER_REQUIRED",
                  message: "Missing the Documents page request header.",
                },
              });
              return;
            }
            await ready;
            const url = new URL(req.url ?? "/", "http://localhost");
            const route = url.pathname.slice(API_PREFIX.length) || "/";

            if (req.method === "GET" && route === "/list") {
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

            if (req.method === "POST" && route === "/create") {
              const body = await readJsonBody(req);
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

            if (req.method === "POST" && route === "/open") {
              const body = await readJsonBody(req);
              if (typeof body.path !== "string" || !body.path.trim()) {
                json(res, 400, {
                  ok: false,
                  error: {
                    code: "PATH_REQUIRED",
                    message: "Give a path to a document inside the workspace.",
                  },
                });
                return;
              }
              const opened = await service.openDocument(body.path.trim());
              json(res, 200, {
                ok: true,
                documentId: opened.documentId,
                warnings: opened.warnings,
                report: opened.report,
              });
              return;
            }

            if (req.method === "GET" && route === "/document") {
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

            if (req.method === "GET" && route === "/export") {
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

            if (req.method === "POST" && route === "/delete") {
              const body = await readJsonBody(req);
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
              error: {
                code: error?.code ?? "DOCUMENT_ERROR",
                message: error?.message ?? String(error),
              },
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
