window.__ModuleLoader__.load({
  id: "dsh-doc-editor-plugin",
  factory(require) {
    const React = require("react");
    const h = React.createElement;
    const PANEL_ID = "dsh-documents";
    const API = "/api/dsh-documents";
    const HEADER = { "x-dsh-documents": "1" };

    function DocIcon(props) {
      const size = typeof props?.size === "number" ? props.size : 18;
      const active = props?.active === true;
      return h(
        "svg",
        {
          viewBox: "0 0 24 24",
          width: size,
          height: size,
          fill: "none",
          stroke: "currentColor",
          strokeWidth: 1.7,
          strokeLinecap: "round",
          strokeLinejoin: "round",
          "aria-hidden": true,
          focusable: false,
          style: { display: "block", opacity: active ? 1 : 0.86 },
        },
        h("path", { d: "M6.5 3.5h7.2l4.3 4.3v12.7H6.5z" }),
        h("path", { d: "M13.7 3.5v4.3h4.3" }),
        h("path", { d: "M9.2 12h6.1M9.2 15h6.1M9.2 18h3.6" }),
      );
    }

    function styles() {
      return h(
        "style",
        null,
        `
        .dshdoc-root { display:flex; flex-direction:column; height:100%; min-height:0; color:var(--dsh-text, inherit); font:inherit; }
        .dshdoc-head { display:flex; align-items:center; gap:10px; padding:14px 18px; border-bottom:1px solid var(--dsh-border, rgba(127,127,127,.22)); flex-wrap:wrap; }
        .dshdoc-title { font-size:16px; font-weight:650; margin:0; }
        .dshdoc-sub { font-size:12px; opacity:.62; margin-left:2px; }
        .dshdoc-spacer { flex:1 1 auto; }
        .dshdoc-btn { display:inline-flex; align-items:center; gap:6px; padding:6px 11px; border-radius:8px; border:1px solid var(--dsh-border, rgba(127,127,127,.28)); background:var(--dsh-surface, transparent); color:inherit; font:inherit; font-size:13px; cursor:pointer; }
        .dshdoc-btn:hover { background:var(--dsh-hover, rgba(127,127,127,.12)); }
        .dshdoc-btn[disabled] { opacity:.5; cursor:not-allowed; }
        .dshdoc-btn.primary { background:var(--dsh-accent, #247bbf); border-color:transparent; color:#fff; }
        .dshdoc-btn.danger { color:#b42318; }
        .dshdoc-input { padding:6px 9px; border-radius:8px; border:1px solid var(--dsh-border, rgba(127,127,127,.28)); background:transparent; color:inherit; font:inherit; font-size:13px; min-width:220px; }
        .dshdoc-body { display:grid; grid-template-columns:minmax(240px, 320px) 1fr; gap:0; flex:1 1 auto; min-height:0; }
        .dshdoc-list { border-right:1px solid var(--dsh-border, rgba(127,127,127,.22)); overflow:auto; padding:8px; }
        .dshdoc-row { display:flex; flex-direction:column; gap:2px; width:100%; text-align:left; padding:9px 10px; border-radius:9px; border:0; background:transparent; color:inherit; font:inherit; cursor:pointer; }
        .dshdoc-row:hover { background:var(--dsh-hover, rgba(127,127,127,.1)); }
        .dshdoc-row[aria-current="true"] { background:var(--dsh-selected, rgba(36,123,191,.14)); }
        .dshdoc-row strong { font-size:13px; font-weight:600; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
        .dshdoc-row span { font-size:11.5px; opacity:.62; }
        .dshdoc-empty { padding:26px 14px; text-align:center; font-size:13px; opacity:.66; line-height:1.5; }
        .dshdoc-detail { overflow:auto; padding:18px 22px 30px; min-width:0; }
        .dshdoc-h2 { margin:0 0 4px; font-size:18px; font-weight:650; }
        .dshdoc-meta { font-size:12px; opacity:.66; margin-bottom:14px; }
        .dshdoc-stats { display:flex; gap:8px; flex-wrap:wrap; margin:0 0 16px; }
        .dshdoc-chip { font-size:11.5px; padding:4px 9px; border-radius:999px; background:var(--dsh-hover, rgba(127,127,127,.12)); }
        .dshdoc-actions { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:18px; }
        .dshdoc-h3 { font-size:12px; text-transform:uppercase; letter-spacing:.06em; opacity:.6; margin:18px 0 8px; }
        .dshdoc-outline { margin:0; padding:0; list-style:none; font-size:13px; }
        .dshdoc-outline li { padding:3px 0 3px calc(var(--lvl, 0) * 14px); opacity:.92; }
        .dshdoc-pre { white-space:pre-wrap; font-size:13px; line-height:1.55; border:1px solid var(--dsh-border, rgba(127,127,127,.22)); border-radius:10px; padding:12px 14px; max-height:340px; overflow:auto; margin:0; background:var(--dsh-surface, transparent); }
        .dshdoc-error { margin:0 18px 14px; padding:9px 12px; border-radius:9px; background:rgba(180,35,24,.12); color:#b42318; font-size:12.5px; }
        .dshdoc-note { font-size:12px; opacity:.62; margin-top:14px; line-height:1.5; }
        @media (max-width: 760px) { .dshdoc-body { grid-template-columns:1fr; } .dshdoc-list { border-right:0; border-bottom:1px solid var(--dsh-border, rgba(127,127,127,.22)); max-height:220px; } }
        `,
      );
    }

    async function call(path, options = {}) {
      const response = await fetch(`${API}${path}`, {
        ...options,
        headers: { ...HEADER, ...(options.headers ?? {}) },
      });
      const type = response.headers.get("content-type") ?? "";
      if (!type.includes("application/json")) {
        if (!response.ok)
          throw new Error(`Request failed (${response.status})`);
        return response;
      }
      const payload = await response.json();
      if (!payload.ok) {
        const error = payload.error ?? {};
        throw new Error(error.message ?? "Request failed");
      }
      return payload;
    }

    function download(documentId, format, title) {
      return fetch(
        `${API}/export?id=${encodeURIComponent(documentId)}&format=${encodeURIComponent(format)}`,
        {
          headers: HEADER,
        },
      ).then(async (response) => {
        if (!response.ok) {
          const text = await response.text().catch(() => "");
          throw new Error(text || `Export failed (${response.status})`);
        }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = `${(title || "document").replace(/[^\w .-]+/g, "_").slice(0, 80)}.${format}`;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        URL.revokeObjectURL(url);
      });
    }

    function DocumentsPanel() {
      const [documents, setDocuments] = React.useState([]);
      const [workspace, setWorkspace] = React.useState("");
      const [readOnly, setReadOnly] = React.useState(false);
      const [selected, setSelected] = React.useState(null);
      const [detail, setDetail] = React.useState(null);
      const [busy, setBusy] = React.useState(false);
      const [error, setError] = React.useState("");
      const [openPath, setOpenPath] = React.useState("");
      const [newTitle, setNewTitle] = React.useState("");
      const [confirmDelete, setConfirmDelete] = React.useState(false);

      const refresh = React.useCallback(async () => {
        try {
          const payload = await call("/list");
          setDocuments(payload.documents ?? []);
          setWorkspace(payload.workspace ?? "");
          setReadOnly(payload.readOnly === true);
          setError("");
          return payload.documents ?? [];
        } catch (cause) {
          setError(cause.message);
          return [];
        }
      }, []);

      React.useEffect(() => {
        void refresh();
      }, [refresh]);

      const openDetail = React.useCallback(async (documentId) => {
        setSelected(documentId);
        setDetail(null);
        setConfirmDelete(false);
        try {
          const payload = await call(
            `/document?id=${encodeURIComponent(documentId)}`,
          );
          setDetail(payload);
          setError("");
        } catch (cause) {
          setError(cause.message);
        }
      }, []);

      const run = async (operation, success) => {
        setBusy(true);
        setError("");
        try {
          const result = await operation();
          const list = await refresh();
          if (typeof success === "function") await success(result, list);
        } catch (cause) {
          setError(cause.message);
        } finally {
          setBusy(false);
        }
      };

      const create = () =>
        run(
          () =>
            call("/create", {
              method: "POST",
              body: JSON.stringify({
                title: newTitle.trim() || "Untitled document",
              }),
            }),
          async (result, list) => {
            setNewTitle("");
            const id = result.documentId;
            const exists = list.some((item) => item.id === id);
            if (exists) await openDetail(id);
          },
        );

      const open = () =>
        run(
          () =>
            call("/open", {
              method: "POST",
              body: JSON.stringify({ path: openPath.trim() }),
            }),
          async (result) => {
            setOpenPath("");
            await openDetail(result.documentId);
          },
        );

      const remove = () =>
        run(
          () =>
            call("/delete", {
              method: "POST",
              body: JSON.stringify({ id: selected }),
            }),
          async () => {
            setSelected(null);
            setDetail(null);
          },
        );

      const meta = detail?.metadata ?? {};
      const stats = meta.stats ?? meta.statistics ?? {};
      const outline = detail?.outline ?? [];
      const selectedRow = documents.find((item) => item.id === selected);

      return h(
        "div",
        { className: "dshdoc-root" },
        styles(),
        h(
          "header",
          { className: "dshdoc-head" },
          h("h1", { className: "dshdoc-title" }, "Documents"),
          h(
            "span",
            { className: "dshdoc-sub" },
            readOnly ? "read-only" : workspace || "workspace",
          ),
          h("span", { className: "dshdoc-spacer" }),
          h("input", {
            className: "dshdoc-input",
            placeholder: "New document title",
            value: newTitle,
            disabled: busy || readOnly,
            onChange: (event) => setNewTitle(event.target.value),
            onKeyDown: (event) => {
              if (event.key === "Enter") void create();
            },
            "aria-label": "New document title",
          }),
          h(
            "button",
            {
              className: "dshdoc-btn primary",
              onClick: () => void create(),
              disabled: busy || readOnly,
            },
            "New document",
          ),
          h(
            "button",
            {
              className: "dshdoc-btn",
              onClick: () => void refresh(),
              disabled: busy,
            },
            "Refresh",
          ),
        ),
        error
          ? h("div", { className: "dshdoc-error", role: "alert" }, error)
          : null,
        h(
          "div",
          { className: "dshdoc-body" },
          h(
            "div",
            { className: "dshdoc-list", role: "list" },
            h(
              "div",
              { style: { display: "flex", gap: 6, padding: "2px 2px 10px" } },
              h("input", {
                className: "dshdoc-input",
                style: { minWidth: 0, flex: "1 1 auto" },
                placeholder: "Open a file path",
                value: openPath,
                disabled: busy,
                onChange: (event) => setOpenPath(event.target.value),
                onKeyDown: (event) => {
                  if (event.key === "Enter" && openPath.trim()) void open();
                },
                "aria-label": "Path of a document to open",
              }),
              h(
                "button",
                {
                  className: "dshdoc-btn",
                  onClick: () => void open(),
                  disabled: busy || !openPath.trim(),
                },
                "Open",
              ),
            ),
            documents.length === 0
              ? h(
                  "div",
                  { className: "dshdoc-empty" },
                  "No open documents. Create one here, open a DOCX/Pages/DOC file by path, or ask the agent to call the document tools.",
                )
              : documents.map((item) =>
                  h(
                    "button",
                    {
                      key: item.id,
                      className: "dshdoc-row",
                      role: "listitem",
                      "aria-current": item.id === selected ? "true" : undefined,
                      onClick: () => void openDetail(item.id),
                    },
                    h("strong", null, item.title || "Untitled document"),
                    h(
                      "span",
                      null,
                      [
                        item.format,
                        item.words ? `${item.words} words` : null,
                        item.dirty ? "unsaved" : null,
                        item.sections
                          ? `${item.sections} section${item.sections === 1 ? "" : "s"}`
                          : null,
                      ]
                        .filter(Boolean)
                        .join(" · "),
                    ),
                  ),
                ),
          ),
          h(
            "section",
            { className: "dshdoc-detail" },
            selected === null
              ? h(
                  "div",
                  { className: "dshdoc-empty" },
                  "Select a document to inspect it, export DOCX/PDF, or review it with the document tools.",
                )
              : detail === null
                ? h("div", { className: "dshdoc-empty" }, "Loading document…")
                : h(
                    React.Fragment,
                    null,
                    h(
                      "h2",
                      { className: "dshdoc-h2" },
                      meta.metadata?.title ?? selectedRow?.title ?? selected,
                    ),
                    h(
                      "div",
                      { className: "dshdoc-meta" },
                      `${detail.sourceFormat ?? "docx"} · ${detail.layout?.pageCount ?? "?"} page(s) · modified ${
                        meta.metadata?.modified
                          ? new Date(meta.metadata.modified).toLocaleString()
                          : "unknown"
                      }`,
                    ),
                    h(
                      "div",
                      { className: "dshdoc-stats" },
                      h(
                        "span",
                        { className: "dshdoc-chip" },
                        `${stats.words ?? 0} words`,
                      ),
                      h(
                        "span",
                        { className: "dshdoc-chip" },
                        `${stats.characters ?? 0} characters`,
                      ),
                      h(
                        "span",
                        { className: "dshdoc-chip" },
                        `${stats.paragraphs ?? 0} paragraphs`,
                      ),
                      h(
                        "span",
                        { className: "dshdoc-chip" },
                        `${stats.tables ?? 0} tables`,
                      ),
                      h(
                        "span",
                        { className: "dshdoc-chip" },
                        `${stats.images ?? 0} images`,
                      ),
                    ),
                    h(
                      "div",
                      { className: "dshdoc-actions" },
                      h(
                        "button",
                        {
                          className: "dshdoc-btn",
                          disabled: busy,
                          onClick: () =>
                            void download(
                              selected,
                              "docx",
                              meta.metadata?.title ?? selected,
                            ).catch((cause) => setError(cause.message)),
                        },
                        "Download DOCX",
                      ),
                      h(
                        "button",
                        {
                          className: "dshdoc-btn",
                          disabled: busy,
                          onClick: () =>
                            void download(
                              selected,
                              "pdf",
                              meta.metadata?.title ?? selected,
                            ).catch((cause) => setError(cause.message)),
                        },
                        "Export PDF",
                      ),
                      h(
                        "button",
                        {
                          className: "dshdoc-btn",
                          disabled: busy,
                          onClick: () =>
                            void download(
                              selected,
                              "html",
                              meta.metadata?.title ?? selected,
                            ).catch((cause) => setError(cause.message)),
                        },
                        "Export HTML",
                      ),
                      h(
                        "button",
                        {
                          className: "dshdoc-btn",
                          disabled: busy,
                          onClick: () =>
                            void download(
                              selected,
                              "md",
                              meta.metadata?.title ?? selected,
                            ).catch((cause) => setError(cause.message)),
                        },
                        "Markdown",
                      ),
                      readOnly
                        ? null
                        : h(
                            "button",
                            {
                              className: "dshdoc-btn danger",
                              disabled: busy,
                              onClick: () =>
                                confirmDelete
                                  ? void remove()
                                  : setConfirmDelete(true),
                              onBlur: () => setConfirmDelete(false),
                            },
                            confirmDelete ? "Confirm delete" : "Delete",
                          ),
                    ),
                    outline.length > 0
                      ? h(
                          React.Fragment,
                          null,
                          h("div", { className: "dshdoc-h3" }, "Outline"),
                          h(
                            "ul",
                            { className: "dshdoc-outline" },
                            outline.slice(0, 40).map((entry, index) =>
                              h(
                                "li",
                                {
                                  key: `${entry.blockId ?? index}-${index}`,
                                  style: {
                                    "--lvl": Math.max(
                                      0,
                                      (entry.level ?? 1) - 1,
                                    ),
                                  },
                                },
                                entry.text,
                              ),
                            ),
                          ),
                        )
                      : null,
                    h("div", { className: "dshdoc-h3" }, "Text"),
                    h(
                      "pre",
                      { className: "dshdoc-pre" },
                      detail.text || "(empty document)",
                    ),
                    detail.truncated
                      ? h(
                          "div",
                          { className: "dshdoc-note" },
                          "Text is truncated in this panel.",
                        )
                      : null,
                    h(
                      "div",
                      { className: "dshdoc-note" },
                      "Ask the agent to edit this document with the document_* tools, or run the standalone editor from the repository (see INSTALL.md) for the full Google Docs / Word / Pages / LibreOffice-style interface.",
                    ),
                  ),
          ),
        ),
      );
    }

    return {
      inject: ["slots"],
      apply(ctx) {
        ctx.slots.inject("sidebar.panellist", () =>
          ctx.slots.register(
            {
              name: "sidebar.panellist",
              id: PANEL_ID,
              order: 30,
              label: () => "Documents",
            },
            DocIcon,
          ),
        );
        ctx.slots.inject("main", () =>
          ctx.slots.register({ name: "main", key: PANEL_ID }, DocumentsPanel),
        );
      },
    };
  },
});
