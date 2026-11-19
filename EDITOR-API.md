# Documents panel — editor API contract

The Documents page inside the DeepSeek Harness Web GUI is a **real editor**, not
a preview. The Client half (`client.js`) renders a Word-style ribbon and an
editable document surface; the Host half (`index.js`) owns the document model and
applies every change through the shared engine. This file is the contract both
halves implement.

All routes live under `/api/dsh-documents` and require the header
`x-dsh-documents: 1`. Responses are JSON unless stated. Failures return
`{ ok: false, error: { code, message } }` with HTTP 200 (the page renders the
message) except for export, which returns the file bytes.

## Workspaces, folder browsing and opening files

The panel never asks a user to type a path unless native pickers are
unavailable. Two Host routes back the pickers and one switches the document
workspace; opening a document also prepares a Harness workspace and chat
session at the file's folder.

**The rule: any absolute local directory (or file) the Host process can read.**
A document may live on another volume (`D:\Docs`), in a shared folder
(`C:\Users\Public\Documents`), or in a project outside the session working
directory. The old Home/workspace/cwd confinement is gone; what remains is the
path _shape_ (absolute, canonicalized, real) and the importable-extension gate
for opening. Reads are read-only and never execute or parse a file to build a
listing.

### `GET /browse?path=<absolute directory>`

Lists one directory for the in-panel folder browser. With no `path` (or an
empty one), returns the browse roots.

```json
{
  "ok": true,
  "path": "D:\\Docs\\Reports",
  "parent": "D:\\Docs",
  "roots": [
    { "label": "Home", "path": "C:\\Users\\me" },
    { "label": "Workspace", "path": "…" },
    { "label": "Working directory", "path": "…" },
    { "label": "Local disk (D:)", "path": "D:\\" }
  ],
  "entries": [
    {
      "name": "Q3 report.docx",
      "path": "D:\\Docs\\Reports\\Q3 report.docx",
      "kind": "document",
      "importable": true,
      "size": 41233,
      "modified": "2026-11-01T09:12:00.000Z"
    },
    {
      "name": "Archive",
      "path": "D:\\Docs\\Reports\\Archive",
      "kind": "directory",
      "importable": false
    }
  ],
  "truncated": false,
  "skipped": 0
}
```

`roots` is a shortcut list, not a boundary: Home, the current document
workspace, the session working directory, then every existing filesystem root
(each drive letter on Windows, probed without a shell; `/` on other platforms).

Browsing rules:

- Absolute paths only. A missing or empty `path` asks for the root list;
  a relative path fails with `INVALID_PATH`.
- The requested path is canonicalized with `realpath` before use. `..` is
  therefore allowed when it lands on a real directory (`D:\a\..\b` names
  `D:\b`) and fails with `DIRECTORY_NOT_FOUND` when the resolved target does
  not exist. A `..` never names a path the Host has not resolved.
- A file passed as a directory fails with `NOT_A_DIRECTORY`.
- `parent` is the containing directory, or `null` at a filesystem root.
- `entries` is sorted directories-first then by name; `kind` is
  `directory`, `document` or `other`, and `importable` comes from this build's
  format registry (never from the extension list the browser uses for `kind`).
  At most 500 entries are returned; `truncated` says more were found.
- An entry whose `stat` fails (permission denied, dangling symlink) is counted
  in `skipped` instead of failing the whole listing. A directory that cannot be
  read at all returns the route error.
- Only names, kinds, sizes and timestamps are returned — never file contents.
  A symlinked entry keeps its lexical path inside the listed directory;
  browsing it later re-validates wherever it really points.
- UNC shares (`\\server\share`) are browsable when the OS resolves them; the
  same `realpath` rule applies.

### `GET /workspace`

```json
{ "ok": true, "path": "…", "readOnly": false, "documents": [ … ] }
```

### `POST /workspace { "path": "<absolute directory>", "create": false }`

Switches the panel's document workspace to `path`, which may be **any absolute
local directory** the Host can reach (created when `create` is true, which the
folder picker uses for a new folder). It also **finds or creates the Harness
workspace** for that directory through the Host's `workspaceRegistry`, so the
caller can start a chat session there.

```json
{
  "ok": true,
  "path": "D:\\Docs\\Reports",
  "created": true,
  "dshWorkspaceId": "…"
}
```

Workspace rules:

- Relative paths fail with `INVALID_PATH`; a file fails with
  `NOT_A_DIRECTORY`; a missing folder without `create` fails with
  `DIRECTORY_NOT_FOUND`.
- A filesystem root (`C:\`, `/`) is refused with `INVALID_PATH`: a workspace
  keeps its own `documents/` and `audit.log` entries, so a root is a typo
  rather than a destination.
- In read-only mode (`DSH_DOC_READ_ONLY=1`) `create:true` still fails with
  `READ_ONLY_MODE`; switching to an existing folder is not a mutation and
  stays allowed.
- The reported `path` is the canonical (realpath) directory, so the Client can
  compare it with what `/browse` returns.

`dshWorkspaceId` is omitted when the profile has no workspace registry; the
Client then only switches the document workspace and says so. Switching keeps
documents that are already open in the previous workspace in memory but moves
new opens and saves to the new one.

### `POST /open-path { "path": "<absolute file path>" }`

Opens a document that the folder browser selected, without uploading it. The
file may live in **any absolute local directory** the Host can read; it is
parsed through the engine, which records the audit entry.

```json
{
  "ok": true,
  "documentId": "doc-…",
  "path": "D:\\Docs\\Reports\\Q3 report.docx",
  "directory": "D:\\Docs\\Reports",
  "title": "Q3 report",
  "warnings": []
}
```

Opening rules: a relative path fails with `INVALID_PATH`; a missing file with
`FILE_NOT_FOUND`; a directory with `NOT_A_FILE`; and any extension this build
cannot import with `FORMAT_UNSUPPORTED`. `directory` is the canonical folder,
which the Client passes to `POST /workspace`.

`POST /open` (multipart upload) returns the same `directory` field so the Client
can switch there too. The Client then calls `POST /workspace` and, when
`dshWorkspaceId` comes back, starts a new chat session in that workspace through
the Client's workspace service (`uiWorkspace.startSession`), switching to the
conversation only when the user asks (the panel shows a "Chat ready" chip).

Mutating routes stay confined: `POST /save`, `POST /saveAs` and `POST /delete`
write inside the active document workspace only, and read-only mode still
refuses them.

## Reads

### `GET /bootstrap?id=<documentId>`

The single call the editor needs to start.

```json
{
  "ok": true,
  "workspace": "C:\\Users\\me\\.dsh-documents",
  "readOnly": false,
  "formats": { "import": ["docx", "..."], "export": ["docx", "pdf", "..."] },
  "documentId": "doc-mwmvia2o-rx1r7r",
  "title": "Quarterly report",
  "dirty": false,
  "revision": 3,
  "sourceFormat": "docx",
  "document": { "schemaVersion": "1.0.0", "...": "the full DocDocument model" },
  "blockPages": { "p1": 1, "p2": 1, "p3": 2 },
  "pageCount": 2,
  "stats": {
    "words": 812,
    "characters": 5200,
    "paragraphs": 96,
    "tables": 3,
    "images": 2
  },
  "outline": [{ "level": 1, "text": "Introduction", "blockId": "p4" }],
  "capabilities": {
    "ai": true,
    "formats": true,
    "convert": true,
    "renderer": true
  }
}
```

`blockPages` maps every body block id to its 1-based page so the page can draw
page separators. `document` is the authoritative model; the Client keeps it in
memory and patches it locally for instant feedback.

### `GET /list`, `GET /document?id=`, `GET /export?id=&format=`

Unchanged from 1.0.1 (`/document` also returns metadata/outline/stats for the
side list; `/export` streams bytes).

### `GET /search?id=&q=&regex=&caseSensitive=`

```json
{
  "ok": true,
  "matches": [
    { "blockId": "p12", "start": 4, "end": 11, "text": "budget", "page": 2 }
  ]
}
```

## Writes

### `POST /op`

```json
{ "id": "doc-…", "op": "replaceText", "args": { "range": { "start": {...}, "end": {...} }, "text": "x" } }
```

Response:

```json
{
  "ok": true,
  "revision": 4,
  "dirty": true,
  "title": "Quarterly report",
  "stats": { "words": 813, "...": 0 },
  "pageCount": 2,
  "result": {
    "paragraphIds": ["p3"],
    "": "raw engine result (small, op specific)"
  }
}
```

### `POST /ops`

```json
{ "id": "doc-…", "ops": [ { "op": "insertText", "args": {...} }, { "op": "formatText", "args": {...} } ] }
```

Applies in order and returns the same envelope plus `"applied": 2`. On a failing
op it stops and returns `{ ok: false, applied, error }`; earlier ops stay applied
(the engine has no multi-op transaction; the Client refetches after a failure).

### `POST /save`

```json
{ "id": "doc-…", "path": "optional/within/workspace.docx" }
```

```json
{
  "ok": true,
  "path": "…/documents/<id>/document.docx",
  "bytes": 17421,
  "revision": 4
}
```

`POST /saveAs` accepts `{ id, path }` and writes the format implied by the
extension; `POST /create`, `POST /open`, `POST /delete` are unchanged.

## Supported `op` values

Positions are `{ blockId, offset }` (UTF-16 offset inside the paragraph text);
ranges are `{ start, end }`. A missing `offset` means the start of the block.

| op                                | args                                                                                          |
| --------------------------------- | --------------------------------------------------------------------------------------------- |
| `insertText`                      | `position`, `text`, `props?`                                                                  |
| `replaceText`                     | `range`, `text`, `props?`                                                                     |
| `deleteRange`                     | `range`                                                                                       |
| `formatText`                      | `range`, `props`, `clear?`                                                                    |
| `formatParagraph`                 | `blockIds`, `props`                                                                           |
| `applyStyle`                      | `blockIds`, `styleId`, `clearDirectFormatting?`                                               |
| `insertHeading`                   | `text`, `level`, `afterBlockId?`, `styleId?`                                                  |
| `insertList`                      | `items`, `kind`, `afterBlockId?`, `level?`, `startAt?`                                        |
| `insertTable`                     | `rows`, `columns`, `data?`, `style?`, `afterBlockId?`, `headerRow?`                           |
| `editTable`                       | `tableId`, `operation`, `index?`, `region?`, `cellId?`                                        |
| `insertImage`                     | `position`, `path?` or `dataBase64`, `contentType?`, `width?`, `height?`, `alt?`, `floating?` |
| `insertHyperlink`                 | `range`, `href?`, `anchor?`, `tooltip?`                                                       |
| `insertPageBreak`                 | `position`, `kind?`                                                                           |
| `insertSectionBreak`              | `sectionIndex?`, `blockIndex?`, `type?`, `landscape?`                                         |
| `insertFootnote`                  | `position`, `text`, `type?`                                                                   |
| `insertToc`                       | `afterBlockId?`, `minLevel?`, `maxLevel?`, `includeTitle?`                                    |
| `insertEquation`                  | `position`, `linear`, `display?`                                                              |
| `setPageLayout`                   | `sectionIndex`, `pageSize?`, `orientation?`, `margins?`                                       |
| `setColumns`                      | `sectionIndex`, `count`, `space?`                                                             |
| `editHeaderFooter`                | `sectionIndex?`, `variant?`, `part` (`header`/`footer`), `text`                               |
| `addComment`                      | `range`, `text`, `author?`                                                                    |
| `setMetadata`                     | `patch` (title/subject/keywords/creator)                                                      |
| `trackChanges`                    | `enabled`                                                                                     |
| `acceptChanges` / `rejectChanges` | `revisionId?`                                                                                 |
| `undo` / `redo`                   | —                                                                                             |

`undo` and `redo` return `{ ok: true, document }` (the full model) because the
caret and page map must be rebuilt; every other op returns the envelope above.

## Client responsibilities

- Keep the model in memory; typing patches the affected paragraph locally and
  posts a debounced `replaceText`/`insertText` so the caret never jumps.
- Structural or layout-affecting ops (table, image, break, TOC, page layout,
  style, undo/redo) refetch `/bootstrap` after the op settles.
- Never send an offset that the model cannot resolve; clamp to the paragraph's
  text length and verify the block exists before posting.
- Treat the document as untrusted data: text is rendered with `textContent`
  semantics (never `innerHTML` from document strings).
