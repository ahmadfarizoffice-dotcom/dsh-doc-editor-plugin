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
  "stats": { "words": 812, "characters": 5200, "paragraphs": 96, "tables": 3, "images": 2 },
  "outline": [{ "level": 1, "text": "Introduction", "blockId": "p4" }],
  "capabilities": { "ai": true, "formats": true, "convert": true, "renderer": true }
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
{ "ok": true, "matches": [{ "blockId": "p12", "start": 4, "end": 11, "text": "budget", "page": 2 }] }
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
  "result": { "paragraphIds": ["p3"], "": "raw engine result (small, op specific)" }
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
{ "ok": true, "path": "…/documents/<id>/document.docx", "bytes": 17421, "revision": 4 }
```

`POST /saveAs` accepts `{ id, path }` and writes the format implied by the
extension; `POST /create`, `POST /open`, `POST /delete` are unchanged.

## Supported `op` values

Positions are `{ blockId, offset }` (UTF-16 offset inside the paragraph text);
ranges are `{ start, end }`. A missing `offset` means the start of the block.

| op | args |
|---|---|
| `insertText` | `position`, `text`, `props?` |
| `replaceText` | `range`, `text`, `props?` |
| `deleteRange` | `range` |
| `formatText` | `range`, `props`, `clear?` |
| `formatParagraph` | `blockIds`, `props` |
| `applyStyle` | `blockIds`, `styleId`, `clearDirectFormatting?` |
| `insertHeading` | `text`, `level`, `afterBlockId?`, `styleId?` |
| `insertList` | `items`, `kind`, `afterBlockId?`, `level?`, `startAt?` |
| `insertTable` | `rows`, `columns`, `data?`, `style?`, `afterBlockId?`, `headerRow?` |
| `editTable` | `tableId`, `operation`, `index?`, `region?`, `cellId?` |
| `insertImage` | `position`, `path?` or `dataBase64`, `contentType?`, `width?`, `height?`, `alt?`, `floating?` |
| `insertHyperlink` | `range`, `href?`, `anchor?`, `tooltip?` |
| `insertPageBreak` | `position`, `kind?` |
| `insertSectionBreak` | `sectionIndex?`, `blockIndex?`, `type?`, `landscape?` |
| `insertFootnote` | `position`, `text`, `type?` |
| `insertToc` | `afterBlockId?`, `minLevel?`, `maxLevel?`, `includeTitle?` |
| `insertEquation` | `position`, `linear`, `display?` |
| `setPageLayout` | `sectionIndex`, `pageSize?`, `orientation?`, `margins?` |
| `setColumns` | `sectionIndex`, `count`, `space?` |
| `editHeaderFooter` | `sectionIndex?`, `variant?`, `part` (`header`/`footer`), `text` |
| `addComment` | `range`, `text`, `author?` |
| `setMetadata` | `patch` (title/subject/keywords/creator) |
| `trackChanges` | `enabled` |
| `acceptChanges` / `rejectChanges` | `revisionId?` |
| `undo` / `redo` | — |

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
