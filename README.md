# Documents for DeepSeek Harness

Reads, creates, edits, converts and exports documents from inside DeepSeek
Harness, with a **Documents** panel in the main column and document tools the
agent can call.

The bundled engine is the same one used by the standalone editor and the MCP
server in
[ahmadfarizoffice-dotcom/dsh-doc-editor](https://github.com/ahmadfarizoffice-dotcom/dsh-doc-editor):
a versioned document model, a real pagination/layout engine, DOCX round-trip
import/export, format adapters and a conversion engine with loss reporting.

![The Documents panel and the document tools](docs/preview.svg)

## Install

**Requirement:** DeepSeek Harness **0.2.2-rc.3 or newer** (declared as
`engines.dsh`). No other dependency, no build step, no registry package.

### From the Plugins page

1. Open **Plugins** in the left sidebar.
2. Choose **Add plugin**.
3. Paste this repository link and confirm:

   ```
   https://github.com/ahmadfarizoffice-dotcom/dsh-doc-editor-plugin
   ```

4. Enable **Documents** if it is not enabled yet.

The same spec is accepted by the agent tool:

```
plugin_manager action: install_bundle target: github:ahmadfarizoffice-dotcom/dsh-doc-editor-plugin
```

Both forms are the same Git package. If the page reports
`failedAt: "spec-host"` (“connection to github.com timed out”), the Harness’s
pre-install reachability check failed before pnpm ran; retry, or paste the
`github:` form, which skips that pre-check on this release. A spec pinned to a
commit (`…#<sha>`) can trip the same pre-check; install the branch spec and
update by removing and reinstalling.

Installation runs `pnpm` against the profile, so pnpm must be on `PATH` (the
Harness launcher usually supplies it). With HMR the bundle activates
immediately; without HMR, restart the runtime. The plugin has **no install
scripts and no dependencies**, so there is nothing to approve.

### Verify

- The left rail shows a **Documents** icon; the panel lists open documents.
- In Chat, ask _“Create a DOCX report with a table and export it as PDF”_ and
  the agent calls `document_create`, `document_insert_table`,
  `document_export_pdf`.
- **Plugins → Documents** shows this description, the icon and the detail image.

## What you get

### Documents panel — a Word-like editor

Click the **Documents** icon in the left rail. The panel is a real editor, not a
preview:

- **Ribbon with the seven Word tabs**: File · Home · Insert · Design · Layout ·
  References · View. The ribbon keeps **exactly 1/16 of the window height on every
  tab** (never below 56 px, so captions and labels are not clipped), so switching
  tabs never moves the document; wide tabs scroll sideways instead of growing
  taller, and the rich File details open as an overlay (New, Open, Export,
  Information, open documents) while the ribbon height stays fixed. Icons are
  **20 px with a 1.8 px stroke** and take their colour from the DeepSeek Harness
  theme tokens (`--dsw-alias-label-primary`, `--dsw-alias-brand-primary`, …), so
  they stay legible in light and dark mode instead of being washed out.
- **Any local folder**: **File → Choose folder…** opens the operating system's
  folder dialog and then an in-panel browser that can now list **any absolute
  local directory** (all drives, network shares the OS resolves, any project
  folder), not only the workspace or home directory; **Open document…** uses the
  system file dialog, and _Advanced: path_ accepts an absolute path anywhere on
  disk. Opening a file switches the document workspace to its folder, reuses an
  already-open session for the same file, names the document after the file and
  prepares a Harness workspace plus a chat session there (_Chat ready_ chip).
- **Real multiline paragraphs**: DOCX hard line breaks (`<w:br/>`), Word's
  Shift+Enter and imported multiline text render as real line breaks inside one
  paragraph (the ↵ mark appears only with View → Formatting marks), and
  Shift+Enter inserts a line break instead of splitting the paragraph.
- **Compact, icon-only commands**: buttons show only their icon and reveal the
  command name and shortcut in a hover tooltip (`Bold (Ctrl+B)`). **View → Show
  labels** restores text labels for anyone who prefers them, and the choice is
  remembered.
- **Native pickers instead of typing paths**: **File → Choose folder…** opens the
  operating system's folder dialog (macOS Finder / Windows Explorer) and then an
  in-panel browser with breadcrumbs, roots, folders and only the documents the
  Host can import; **File → Open document…** opens the system file dialog.
  Opening a file **switches the document workspace to that file's folder and
  prepares a Harness workspace and a new chat session there**, with a
  _Chat ready · <folder>_ chip in the panel; **Switch to chat** jumps to that
  session. A path field remains under _Advanced: path_ when no picker is
  available.
- **Edit by clicking into the page**: type in any paragraph or table cell,
  press Enter to split a paragraph, Backspace at the start to merge with the
  previous one, select text and use Bold/Italic/Underline, font family/size,
  colour, highlight, alignment, lists, indents, line and paragraph spacing, or
  the styles gallery. With a caret but no selection, character formatting is
  applied to the whole paragraph, as Word does at an insertion point.
- **Shortcuts**: `Ctrl+B/I/U`, `Ctrl+Z/Y`, `Ctrl+S`, `Ctrl+F`, `Ctrl+A`.
- **Insert** tables, images, page and section breaks, hyperlinks, headers and
  footers, a generated table of contents, footnotes and equations.
- **Layout** changes margins, orientation, paper size and columns; **View**
  changes zoom, ruler, formatting marks, page separators, read-only and focus
  mode.
- **Status bar**: page X of Y, word and character counts, language, zoom slider,
  editing state and save state.
- **File** tab and the left rail keep document management: new, open by path,
  save, save as, export DOCX/PDF/HTML/Markdown, rename and delete.

Every change goes through the same engine the agent tools use, so the panel and
the agent always see one document. Typing is applied locally for an instant
caret and synced a few hundred milliseconds later; structural changes refetch
the page map. Unsupported Word features (page colour, shapes, symbols,
captions, cross-references, page-number fields in headers) are shown **disabled
with a reason** rather than as dead buttons.

### Tools (26)

| Tool                              | Use                                                                      |
| --------------------------------- | ------------------------------------------------------------------------ |
| `document_list_formats`           | Import/export matrix for this build                                      |
| `document_create`                 | New document (title, page size, initial paragraphs)                      |
| `document_open`                   | Import DOCX/DOTX/DOCM, Pages, legacy DOC, ODT, RTF, HTML, Markdown, text |
| `document_get_metadata`           | Metadata, statistics, compatibility report                               |
| `document_get_outline`            | Heading outline with block ids                                           |
| `document_get_text`               | Plain text with counts                                                   |
| `document_search`                 | Literal/regex search with context                                        |
| `document_insert_text`            | Insert text at a position                                                |
| `document_replace_text`           | Replace a range, preserving formatting                                   |
| `document_format_text`            | Character formatting over a range                                        |
| `document_format_paragraph`       | Alignment, spacing, indents                                              |
| `document_apply_style`            | Apply a named style                                                      |
| `document_insert_heading`         | Insert a heading (1–9)                                                   |
| `document_insert_list`            | Bullet or numbered list                                                  |
| `document_insert_table`           | Table with data and header row                                           |
| `document_edit_table`             | Rows, columns, merges, splits                                            |
| `document_insert_image`           | Image from a workspace path or base64                                    |
| `document_insert_hyperlink`       | External URL or bookmark anchor                                          |
| `document_insert_page_break`      | Page, line or column break                                               |
| `document_insert_section_break`   | Section split, optionally landscape                                      |
| `document_undo` / `document_redo` | Undo and redo through the session history                                |
| `document_ai_review`              | Mechanical checks; works without a model                                 |
| `document_ai_summarize`           | Extractive summary offline; abstractive with a provider                  |
| `document_convert`                | Convert between DOCX/ODT/RTF/HTML/Markdown/TXT/PDF with losses           |
| `document_export_pdf`             | Page-accurate PDF from the layout engine                                 |
| `document_save`                   | Save DOCX + lossless model snapshot                                      |

The full 50-tool surface, resources and prompts are available from the MCP
server in the main repository; this bundle keeps the context cost small.

## Formats

| Format                | Import              | Export  | Notes                                                                                                                 |
| --------------------- | ------------------- | ------- | --------------------------------------------------------------------------------------------------------------------- |
| DOCX / DOTX / DOCM    | ✅ full             | ✅ full | Round-trip tested; macros preserved, never executed                                                                   |
| Legacy `.doc`         | ✅ text only        | ✖       | Formatting, tables and images are reported as lost                                                                    |
| Apple `.pages`        | ✅ text recovery    | ✖       | Modern Pages (IWA/Snappy) and Pages ’09 XML; structure/styles are not recovered and the result is marked as a preview |
| `.gdoc`               | ✖                   | ✖       | Google Docs pointer files cannot be fetched without Drive OAuth; download as DOCX instead                             |
| ODT / OTT             | ✅                  | ✅      | Headers/footers, tracked changes and math are reported as lost                                                        |
| RTF                   | ✅                  | ✅      | Images and headers are lost                                                                                           |
| HTML / Markdown / TXT | ✅                  | ✅      | Fidelity depends on the target; losses are reported                                                                   |
| PDF                   | ✖ **not supported** | ✅      | PDF export is page-accurate; PDF import is intentionally not implemented                                              |
| XLSX / PPTX           | preview only        | ✖       | Tables/text previews, not editors                                                                                     |

Every conversion returns a loss report instead of dropping content silently.

## Configuration

Edit the `dsh-documents` row's `config` in the profile patch, or override it in
your own profile patch:

```yaml
- id: dsh-documents
  config:
    workspaceDir: 'D:\Documents\Harness'
    readOnly: false
    tools:
      [
        "document_create",
        "document_open",
        "document_get_text",
        "document_export_pdf",
      ]
```

| Field          | Default                                       | Meaning                              |
| -------------- | --------------------------------------------- | ------------------------------------ |
| `workspaceDir` | `$DSH_DOC_WORKSPACE`, else `~/.dsh-documents` | Where documents and `audit.log` live |
| `readOnly`     | `false`                                       | Refuse every mutation                |
| `tools`        | built-in subset                               | Exact `document_*` names to expose   |

Also read: `DSH_DOC_WORKSPACE`, `DSH_DOC_READ_ONLY=1`, and for AI
`DSH_DOC_AI_PROVIDER` (`offline`/`openai`/`ollama`), `OPENAI_API_KEY`,
`OPENAI_BASE_URL`, `DSH_DOC_AI_MODEL`, `OLLAMA_HOST`.

## Security

- Reads only from `workspaceDir` plus the session working directory; writes
  only inside `workspaceDir`.
- Path traversal, invalid document ids and ZIP-slip entries are rejected.
- Macros, OLE objects and embedded scripts are never executed.
- Document content is untrusted data for prompts; the AI layer returns
  proposals and never applies them itself.
- Mutations are written to `audit.log` (action, outcome, size; not content).

## Troubleshooting

| Symptom                                    | Fix                                                                  |
| ------------------------------------------ | -------------------------------------------------------------------- |
| Install fails with `failedAt: 'spec-host'` | The GitHub host is unreachable; check proxy/VPN and retry            |
| Bundle installs but rows are missing       | Enable **Documents** in Plugins; without HMR restart the runtime     |
| `PATH_OUTSIDE_WORKSPACE`                   | Pass a path inside `workspaceDir` or set `workspaceDir`              |
| `FORMAT_UNAVAILABLE`                       | The format has no adapter in this build; see `document_list_formats` |
| `AI_PROVIDER_REQUIRED`                     | Configure a provider; review/summarise still work offline            |
| Panel is empty                             | Create a document or open a path inside the workspace                |

Uninstall: **Plugins → Documents → Remove**, or
`plugin_manager action: remove_bundle target: dsh-doc-editor-plugin`.

## Development

The engine is generated from the main repository:

```bash
git clone https://github.com/ahmadfarizoffice-dotcom/dsh-doc-editor
cd dsh-doc-editor
pnpm install
pnpm run build:plugin        # builds dist/dsh-plugin with a fresh engine.js
```

Then publish `dist/dsh-plugin/` as this repository. Rebuild after engine
changes; `engine.js` is committed so installs stay build-free.

## License

MIT — see [LICENSE](LICENSE).
