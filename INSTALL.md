# Install Documents for DeepSeek Harness

This repository is the installable bundle. It is self-contained: `engine.js`
holds the full document engine, so there is no build step, no npm dependency and
no install script.

## One step

Paste this into **Plugins → Add plugin**:

```
https://github.com/ahmadfarizoffice-dotcom/dsh-doc-editor-plugin
```

or, from an agent session:

```
plugin_manager action: install_bundle target: github:ahmadfarizoffice-dotcom/dsh-doc-editor-plugin
```

Requirements: DeepSeek Harness **0.2.2-rc.3 or newer**, and pnpm available to
the profile. After installation, enable **Documents** if prompted. With HMR the
panel appears immediately; without HMR the Harness must be restarted.

## First run

1. Open the **Documents** icon in the left rail.
2. Type a title and choose **New document**, or paste a path inside the
   workspace and choose **Open**.
3. Inspect counts, outline and text; download DOCX, export PDF/HTML/Markdown.
4. Ask the agent in Chat to edit the same document with the `document_*` tools.

Documents live in `$DSH_DOC_WORKSPACE`, else `~/.dsh-documents`, unless you set
`workspaceDir` on the `dsh-documents` row.

## Uninstall

**Plugins → Documents → Remove**, or:

```
plugin_manager action: remove_bundle target: dsh-doc-editor-plugin
```

## Getting the standalone editor

The full editor UI (Google Docs / Word / Pages / LibreOffice-style skins),
MCP server, CLI and format adapters live in:

https://github.com/ahmadfarizoffice-dotcom/dsh-doc-editor

See that repository's `INSTALL.md` for the editor, Claude Code, Codex,
Antigravity and OpenCode integration.
