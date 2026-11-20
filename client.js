window.__ModuleLoader__.load({
  id: "dsh-doc-editor-plugin",
  factory(require) {
    const React = require("react");
    const h = React.createElement;
    const PANEL_ID = "dsh-documents";
    const API = "/api/dsh-documents";
    const HEADER = { "x-dsh-documents": "1" };

    /** Ribbon tabs, in Word order. `File` is the document-management tab. */
    const TABS = [
      { id: "file", label: "File" },
      { id: "home", label: "Home" },
      { id: "insert", label: "Insert" },
      { id: "design", label: "Design" },
      { id: "layout", label: "Layout" },
      { id: "references", label: "References" },
      { id: "view", label: "View" },
    ];

    /**
     * Honest reasons for controls whose backend the editor API does not expose.
     * Every one is rendered `disabled` with this text in its `title`.
     */
    const REASONS = {
      pageColor:
        "Page colour is not supported by the engine: the editor API exposes no page-background operation.",
      shapes:
        "Shapes are not supported by the editor API: there is no shape-insertion op.",
      symbols:
        "Symbols and special characters are not supported by the editor API: there is no symbol op.",
      pageBorders:
        "Page borders are not supported by the engine: the editor API exposes no page-border op.",
      paragraphBorders:
        "Paragraph borders and shading are not supported by the editor API: formatParagraph has no border or shading route.",
      bookmark:
        "Bookmark insertion is not exposed by the editor API (no bookmark op).",
      caption:
        "Captions are not exposed by the editor API (no caption or sequence field op).",
      crossReference:
        "Cross-references are not exposed by the editor API (no reference-field op).",
      pageNumber:
        "Page-number fields in headers and footers are not exposed by the editor API; only plain header/footer text can be edited.",
      checklist:
        "Checklist numbering is not supported by the document model; bulleted and numbered lists are.",
      listFromText:
        "The editor API only exposes insertList (a new list); converting the selected paragraphs in place is not available. Move the caret to an empty paragraph to start a list.",
      comment:
        "Comments require the review routes, which this panel does not drive yet; use the agent tools instead.",
      trackChanges:
        "Track Changes is not exposed by the editor API yet (no trackChanges route in this panel).",
      field: "Fields (PAGE, NUMPAGES, REF) are not exposed by the editor API.",
      clipboardUnavailable:
        "The browser blocked clipboard access. Use Ctrl+C, Ctrl+X or Ctrl+V inside the editor.",
      selectionRequired:
        "Select text in the document first: character formatting needs a range.",
      caretRequired:
        "Click into the document text first, then apply formatting (a collapsed caret formats the whole paragraph).",
      readOnly: "The document is open read-only, so edits are disabled.",
      offline:
        "The Host editor API is unavailable, so this command is disabled.",
      noDocument: "Open a document first.",
    };

    /**
     * `localStorage` key for the View → Show labels ribbon preference. The
     * panel is icon-only by default; the toggle restores text labels for users
     * who want them and remembers the choice for this browser.
     */
    const LABELS_KEY = "dsh-doc:labels";

    /**
     * `localStorage` key for the user-chosen ribbon height in CSS pixels. The
     * default is 1/10 of the viewport (clamped); dragging the handle stores an
     * explicit pixel value here and double-clicking removes it again.
     */
    const RIBBON_HEIGHT_KEY = "dsh-doc:ribbon-height";

    /** Never shorter than 64px: captions, labels mode and the gallery clip below. */
    const RIBBON_MIN_HEIGHT = 64;

    /** The 1/10 default is clamped to a band that fits every tab's content. */
    const RIBBON_DEFAULT_MIN = 72;
    const RIBBON_DEFAULT_MAX = 240;

    /** A dragged ribbon stops at half the viewport so the document keeps room. */
    const RIBBON_MAX_VIEWPORT_RATIO = 0.5;

    /** One arrow key or Home/End step while the handle has focus. */
    const RIBBON_KEYBOARD_STEP = 16;

    /**
     * `localStorage` key for the folder the in-panel file chooser reopens. It is
     * written when the user acts on a folder: opening a document from it,
     * picking it with the native picker, or switching the workspace to it.
     */
    const LAST_FOLDER_KEY = "dsh-doc:last-folder";

    /** The browser's viewport height, or 0 where there is no viewport (SSR). */
    function viewportHeight() {
      return typeof window !== "undefined" &&
        Number.isFinite(window.innerHeight) &&
        window.innerHeight > 0
        ? window.innerHeight
        : 0;
    }

    /** Largest allowed ribbon height: 50% of the viewport, never below the floor. */
    function ribbonMaxHeight(height) {
      const viewport = Number(height);
      if (!Number.isFinite(viewport) || viewport <= 0)
        return RIBBON_DEFAULT_MAX;
      return Math.max(
        RIBBON_MIN_HEIGHT,
        Math.round(viewport * RIBBON_MAX_VIEWPORT_RATIO),
      );
    }

    /** The 1/10-of-viewport default height, clamped to the sane band. */
    function defaultRibbonHeight(height) {
      const viewport = Number(height);
      const ratio =
        Number.isFinite(viewport) && viewport > 0
          ? viewport / 10
          : RIBBON_DEFAULT_MIN;
      return Math.round(
        Math.min(RIBBON_DEFAULT_MAX, Math.max(RIBBON_DEFAULT_MIN, ratio)),
      );
    }

    /** Clamp one requested height to 64px … 50vh, rounded to whole pixels. */
    function clampRibbonHeight(value, height) {
      if (value === null || value === undefined || value === "") return null;
      const numeric = Number(value);
      if (!Number.isFinite(numeric)) return null;
      return Math.max(
        RIBBON_MIN_HEIGHT,
        Math.min(ribbonMaxHeight(height), Math.round(numeric)),
      );
    }

    function readStoredRibbonHeight() {
      try {
        if (typeof localStorage === "undefined" || !localStorage) return null;
        const raw = localStorage.getItem(RIBBON_HEIGHT_KEY);
        if (raw === null || raw === "") return null;
        const numeric = Number(raw);
        return Number.isFinite(numeric) ? numeric : null;
      } catch (_error) {
        return null;
      }
    }

    function storeRibbonHeight(value) {
      try {
        if (typeof localStorage === "undefined" || !localStorage) return;
        if (value === null || value === undefined) {
          localStorage.removeItem(RIBBON_HEIGHT_KEY);
          return;
        }
        const numeric = Number(value);
        if (!Number.isFinite(numeric)) return;
        localStorage.setItem(RIBBON_HEIGHT_KEY, String(Math.round(numeric)));
      } catch (_error) {
        /* Storage may be unavailable (private mode): resizing still works. */
      }
    }

    function readStoredFolder() {
      try {
        if (typeof localStorage === "undefined" || !localStorage) return "";
        return String(localStorage.getItem(LAST_FOLDER_KEY) || "").trim();
      } catch (_error) {
        return "";
      }
    }

    function storeFolder(path) {
      const value = String(path || "").trim();
      if (!value) return;
      try {
        if (typeof localStorage === "undefined" || !localStorage) return;
        localStorage.setItem(LAST_FOLDER_KEY, value);
      } catch (_error) {
        /* Storage may be unavailable (private mode): browsing still works. */
      }
    }

    /** The directory that contains `path` (the folder the chooser remembers). */
    function folderPathOf(path) {
      const trimmed = trimTrailingSeparator(path);
      if (!trimmed) return "";
      const index = Math.max(
        trimmed.lastIndexOf("/"),
        trimmed.lastIndexOf("\\"),
      );
      if (index < 0) return "";
      if (index === 0) return trimmed.slice(0, 1);
      const parent = trimmed.slice(0, index);
      return /^[A-Za-z]:$/.test(parent) ? `${parent}\\` : parent;
    }

    /** Keyboard shortcuts the editor implements, used in icon-only tooltips. */
    const SHORTCUTS = {
      undo: "Ctrl+Z",
      redo: "Ctrl+Y",
      bold: "Ctrl+B",
      boldAlt: "Ctrl+Shift+B",
      italic: "Ctrl+I",
      underline: "Ctrl+U",
      find: "Ctrl+F",
      selectAll: "Ctrl+A",
      save: "Ctrl+S",
      open: "Ctrl+O",
      hyperlink: "Ctrl+K",
      numbering: "Ctrl+Shift+7",
      bullets: "Ctrl+Shift+8",
    };

    /**
     * The panel's own keyboard commands.
     *
     * `code` is the physical key (`KeyboardEvent.code`), never `event.key`:
     * Shift+7 reports "&" as `key` on a US layout and a dead key elsewhere,
     * while the physical key stays `Digit7`. `mod` is Ctrl or Cmd.
     *
     * `scope` is the input owner a command may consume:
     *   - `editable` — the document must own the caret (formatting, history,
     *     select all, link, lists).
     *   - `panel` — any focus inside the panel root (save, find, open).
     * `repeat` decides whether an auto-repeated key runs the command again:
     * history may repeat, a formatting toggle must not (a held Ctrl+B would
     * otherwise flip on and off dozens of times).
     */
    const SHORTCUT_TABLE = [
      {
        id: "bold",
        code: "KeyB",
        scope: "editable",
        repeat: false,
        action: "bold",
        label: "Ctrl+B",
      },
      {
        id: "boldShift",
        code: "KeyB",
        shift: true,
        scope: "editable",
        repeat: false,
        action: "bold",
        label: "Ctrl+Shift+B",
      },
      {
        id: "italic",
        code: "KeyI",
        scope: "editable",
        repeat: false,
        action: "italic",
        label: "Ctrl+I",
      },
      {
        id: "underline",
        code: "KeyU",
        scope: "editable",
        repeat: false,
        action: "underline",
        label: "Ctrl+U",
      },
      {
        id: "undo",
        code: "KeyZ",
        scope: "editable",
        repeat: true,
        action: "undo",
        label: "Ctrl+Z",
      },
      {
        id: "redo",
        code: "KeyY",
        scope: "editable",
        repeat: true,
        action: "redo",
        label: "Ctrl+Y",
      },
      {
        id: "redoShift",
        code: "KeyZ",
        shift: true,
        scope: "editable",
        repeat: true,
        action: "redo",
        label: "Ctrl+Shift+Z",
      },
      {
        id: "save",
        code: "KeyS",
        scope: "panel",
        repeat: false,
        action: "save",
        label: "Ctrl+S",
      },
      {
        id: "find",
        code: "KeyF",
        scope: "panel",
        repeat: false,
        action: "find",
        label: "Ctrl+F",
      },
      {
        id: "selectAll",
        code: "KeyA",
        scope: "editable",
        repeat: true,
        action: "selectAll",
        label: "Ctrl+A",
      },
      {
        id: "open",
        code: "KeyO",
        scope: "panel",
        repeat: false,
        action: "open",
        label: "Ctrl+O",
      },
      {
        id: "link",
        code: "KeyK",
        scope: "editable",
        repeat: false,
        action: "link",
        label: "Ctrl+K",
      },
      {
        id: "numberedList",
        code: "Digit7",
        shift: true,
        scope: "editable",
        repeat: false,
        action: "numbering",
        label: "Ctrl+Shift+7",
      },
      {
        id: "bulletList",
        code: "Digit8",
        shift: true,
        scope: "editable",
        repeat: false,
        action: "bullets",
        label: "Ctrl+Shift+8",
      },
    ];

    /** Shown when the host already owns Ctrl+B (Desktop Windows/macOS). */
    const BOLD_RESERVED_NOTE =
      "Ctrl+B is reserved by DeepSeek Harness (sidebar toggle). " +
      "Use Ctrl+Shift+B, or remap it in Settings → Shortcuts.";

    /** Once per page session, not once per panel mount. */
    let boldReservedNoticeShown = false;

    /**
     * Whether the host catalog already binds primary+KeyB to one of its own
     * commands. That binding is consumed by the Desktop main process before the
     * renderer (and outranks local regions in the registry), so the panel must
     * advertise the Ctrl+Shift+B fallback instead of pretending Ctrl+B works.
     */
    function detectBoldReservation(service) {
      if (!service) return false;
      const ours = new Set(["docEditor.bold", "docEditor.boldShift"]);
      try {
        if (
          service.catalog &&
          typeof service.catalog.getSnapshot === "function"
        ) {
          const rows = service.catalog.getSnapshot() || [];
          for (const row of rows) {
            if (!row || ours.has(row.id)) continue;
            const binding = row.binding || null;
            if (binding && binding.code === "KeyB") {
              const modifiers = Array.isArray(binding.modifiers)
                ? binding.modifiers
                : [];
              const plain =
                modifiers.includes("primary") &&
                !modifiers.includes("shift") &&
                !modifiers.includes("alt");
              if (plain) return true;
            }
          }
        }
        if (typeof service.describeBinding === "function") {
          const probe = service.describeBinding({
            code: "KeyB",
            modifiers: ["primary"],
          });
          const conflicts = (probe && probe.conflicts) || [];
          if (conflicts.some((id) => !ours.has(id))) return true;
        }
      } catch (_error) {
        /* an unreadable catalog cannot prove a reservation */
      }
      return false;
    }

    /** Stable per-command identity for the DSH `shortcuts` service. */
    const SHORTCUT_COMMAND_PREFIX = "docEditor.";

    /**
     * Swap `Control` for the device primary modifier, so the same table works on
     * macOS (Cmd) and Windows/Linux (Ctrl) without branching at the call site.
     */
    function shortcutCommandDefaults(entry, primary) {
      const modifiers = [primary];
      if (entry.shift) modifiers.push("shift");
      return {
        "desktop:macos": { code: entry.code, modifiers },
        "desktop:windows": { code: entry.code, modifiers },
        "desktop:linux": { code: entry.code, modifiers },
      };
    }

    /**
     * Whether a native or React-synthetic event is part of an IME composition.
     * React's SyntheticEvent does not expose `isComposing`, so the native event
     * has to be consulted or an IME key would look like a shortcut.
     */
    function eventIsComposing(event) {
      if (!event) return false;
      if (event.isComposing === true) return true;
      return Boolean(
        event.nativeEvent && event.nativeEvent.isComposing === true,
      );
    }

    /**
     * Normalise a keyboard event into the facts the matcher needs, so the
     * predicate is a pure function of data that tests can build directly.
     */
    function shortcutCombo(event) {
      if (!event || typeof event !== "object") return null;
      return {
        code: typeof event.code === "string" ? event.code : "",
        key: typeof event.key === "string" ? event.key : "",
        mod: Boolean(event.ctrlKey || event.metaKey),
        control: Boolean(event.ctrlKey),
        meta: Boolean(event.metaKey),
        shift: Boolean(event.shiftKey),
        alt: Boolean(event.altKey),
      };
    }

    /** The table row a normalised combo selects, or null. Pure. */
    function shortcutForCombo(combo) {
      if (!combo || !combo.mod || combo.alt) return null;
      for (const entry of SHORTCUT_TABLE) {
        if (entry.code !== combo.code) continue;
        if (Boolean(entry.shift) !== Boolean(combo.shift)) continue;
        return entry;
      }
      return null;
    }

    /**
     * Where the keyboard input is coming from, in the panel's terms.
     *
     * `region` drives the `scope` check; `textField`, `dialog`, `composing`,
     * `outsideModal` and `repeat` are independent vetoes so each condition can
     * be exercised on its own (MC/DC decision table in the unit tests).
     */
    function shortcutEnvironment(event, rootElement, options) {
      const opts = options || {};
      const target = event && event.target;
      const element =
        target && target.nodeType === 3 ? target.parentElement : target;
      const withinRoot = Boolean(
        rootElement &&
        element &&
        typeof rootElement.contains === "function" &&
        rootElement.contains(element),
      );
      const active =
        typeof document !== "undefined" ? document.activeElement : null;
      const focusInsideRoot = Boolean(
        rootElement &&
        active &&
        typeof rootElement.contains === "function" &&
        rootElement.contains(active),
      );
      const editable = Boolean(
        withinRoot &&
        element &&
        typeof element.closest === "function" &&
        element.closest(
          '[data-block-id][contenteditable="true"], [data-block-id][contenteditable=""]',
        ),
      );
      const textField = Boolean(
        withinRoot &&
        element &&
        typeof element.closest === "function" &&
        element.closest("input, textarea, select"),
      );
      return {
        region: editable
          ? "editable"
          : withinRoot || focusInsideRoot
            ? "panel"
            : "outside",
        textField,
        dialog: opts.dialog === true,
        composing: eventIsComposing(event) || opts.composing === true,
        repeat: Boolean(event && event.repeat),
        outsideModal: opts.outsideModal === true,
      };
    }

    /** True when a modal outside the panel owns focus (e.g. a DSH dialog). */
    function modalOutsidePanel(rootElement) {
      if (typeof document === "undefined" || !document.activeElement)
        return false;
      const active = document.activeElement;
      if (!active || active === document.body) return false;
      if (
        rootElement &&
        typeof rootElement.contains === "function" &&
        rootElement.contains(active)
      )
        return false;
      return Boolean(
        typeof active.closest === "function" &&
        active.closest(
          '[role="dialog"], [role="alertdialog"], [aria-modal="true"]',
        ),
      );
    }

    /**
     * Decide whether one keyboard event is one of the panel's shortcuts.
     * Returns the table entry or null; never touches the event.
     */
    function matchShortcut(event, environment) {
      if (!event || !environment) return null;
      if (environment.composing) return null;
      if (environment.outsideModal) return null;
      if (environment.dialog) return null;
      if (environment.textField) return null;
      if (environment.region === "outside") return null;
      const entry = shortcutForCombo(shortcutCombo(event));
      if (!entry) return null;
      if (entry.scope === "editable" && environment.region !== "editable")
        return null;
      if (entry.repeat === false && environment.repeat) return null;
      return entry;
    }

    /**
     * Install the window **capture-phase** keydown listener that runs before
     * DSH's own window-bubble shortcut dispatcher. Capture at `window` is the
     * only DOM position that is guaranteed to see the event before a bubble
     * listener on `window`, before React's root delegation, and before any
     * default editing behaviour.
     */
    const shortcutRuntime = {
      bindings: 0,
      capture: 0,
      events: 0,
      matched: 0,
      ignored: 0,
      lastCommand: null,
      serviceRegistered: 0,
      serviceSkipped: [],
      serviceError: null,
      executed: 0,
      lastExec: null,
    };

    function bindShortcutCapture(options) {
      if (typeof window === "undefined" || !window.addEventListener)
        return () => {};
      const opts = options || {};
      const onKeyDown = (event) => {
        shortcutRuntime.events += 1;
        let root = null;
        try {
          root = opts.getRoot ? opts.getRoot() : null;
        } catch (_error) {
          root = null;
        }
        let state = null;
        try {
          state = opts.getState ? opts.getState() : null;
        } catch (_error) {
          state = null;
        }
        const environment = shortcutEnvironment(event, root, {
          dialog: Boolean(state && state.dialog),
          composing: Boolean(state && state.composing),
          outsideModal: modalOutsidePanel(root),
        });
        const entry = matchShortcut(event, environment);
        if (!entry) {
          shortcutRuntime.ignored += 1;
          return;
        }
        // Consume before anything else runs: DSH's dispatcher checks
        // `defaultPrevented`, and capture-phase stop prevents its bubble
        // listener (and React's delegated one) from seeing the event again.
        event.preventDefault();
        event.stopPropagation();
        if (typeof event.stopImmediatePropagation === "function")
          event.stopImmediatePropagation();
        shortcutRuntime.matched += 1;
        shortcutRuntime.lastCommand = entry.id;
        try {
          if (opts.onCommand) opts.onCommand(entry, event);
        } catch (error) {
          if (typeof console !== "undefined" && console.error)
            console.error("Documents panel shortcut failed", error);
        }
      };
      window.addEventListener("keydown", onKeyDown, true);
      shortcutRuntime.capture += 1;
      shortcutRuntime.bindings += 1;
      return () => {
        window.removeEventListener("keydown", onKeyDown, true);
        shortcutRuntime.capture = Math.max(0, shortcutRuntime.capture - 1);
        shortcutRuntime.bindings = Math.max(0, shortcutRuntime.bindings - 1);
      };
    }

    /**
     * Mirrors the binding facts into a read-only diagnostic for tests and the
     * `data-dx-shortcuts` attribute visible to the live v2 probe.
     */
    function shortcutDiagnostics() {
      return {
        capturePhase: true,
        bound: shortcutRuntime.bindings > 0,
        bindings: shortcutRuntime.bindings,
        windowCaptureListeners: shortcutRuntime.capture,
        combos: SHORTCUT_TABLE.map((entry) => ({
          id: entry.id,
          combo: entry.label,
          code: entry.code,
          shift: Boolean(entry.shift),
          scope: entry.scope,
          repeat: entry.repeat,
          action: entry.action,
        })),
        events: shortcutRuntime.events,
        matched: shortcutRuntime.matched,
        ignored: shortcutRuntime.ignored,
        lastCommand: shortcutRuntime.lastCommand,
        executed: shortcutRuntime.executed,
        lastExec: shortcutRuntime.lastExec,
        service: {
          registered: shortcutRuntime.serviceRegistered,
          skipped: shortcutRuntime.serviceSkipped.slice(),
          error: shortcutRuntime.serviceError,
        },
      };
    }

    /**
     * Second path: contribute the same commands to the DSH `shortcuts` service
     * when the page has one. Every failure is contained: a missing, older or
     * conflicting registry must never break the panel, and the capture
     * listener keeps working regardless.
     */
    function registerShortcutCommands(options) {
      const opts = options || {};
      const service = opts.shortcuts;
      const disposers = [];
      if (!service || typeof service.register !== "function") {
        return disposers;
      }
      const onCommand = (id) => {
        if (typeof opts.onCommand === "function") opts.onCommand(id);
      };
      for (const entry of SHORTCUT_TABLE) {
        const commandId = `${SHORTCUT_COMMAND_PREFIX}${entry.id}`;
        try {
          const defaults = shortcutCommandDefaults(entry, "primary");
          if (typeof service.describeBinding === "function") {
            const probe = service.describeBinding({
              code: entry.code,
              modifiers: entry.shift ? ["primary", "shift"] : ["primary"],
            });
            const conflicts = (probe && probe.conflicts) || [];
            if (conflicts.length > 0) {
              shortcutRuntime.serviceSkipped.push({
                id: entry.id,
                reason: "conflict",
                conflicts: conflicts.slice(),
              });
              continue;
            }
            if (probe && probe.issue) {
              shortcutRuntime.serviceSkipped.push({
                id: entry.id,
                reason: "binding-issue",
                issue: String(probe.issue),
              });
              continue;
            }
          }
          const dispose = service.register({
            id: commandId,
            label: () => entry.label,
            aliases: [entry.label.toLowerCase()],
            defaults,
            // DSH models dialogs as modal owners (`modals`), not as a region.
            // `dialog` is kept in `regions` for the requested contract and is
            // inert: dispatch only compares page/editable/terminal.
            regions: ["editable", "page", "dialog"],
            modals: ["dx-dialog", "dialog", "other"],
            resolve: () => ({
              status: "handled",
              run: () => {
                onCommand(entry.id);
              },
            }),
          });
          if (typeof dispose === "function") disposers.push(dispose);
          shortcutRuntime.serviceRegistered += 1;
        } catch (error) {
          shortcutRuntime.serviceError = String(
            (error && error.message) || error,
          );
        }
      }
      return disposers;
    }

    /**
     * The document types the Host can import through `POST /open`; the hidden
     * `<input type="file">` and its accept filter use this exact list.
     */
    const DOCUMENT_ACCEPT =
      ".docx,.dotx,.docm,.doc,.dot,.odt,.ott,.rtf,.html,.htm,.xhtml,.md,.markdown,.txt,.pages,.gdoc";

    const FONT_FALLBACKS = [
      "Calibri",
      "Calibri Light",
      "Arial",
      "Cambria",
      "Candara",
      "Consolas",
      "Courier New",
      "Georgia",
      "Times New Roman",
      "Trebuchet MS",
      "Verdana",
    ];

    const FONT_SIZES = [
      8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72,
    ];

    const ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 1.75, 2];

    const LINE_SPACING = [
      { value: 1, label: "1.0" },
      { value: 1.15, label: "1.15" },
      { value: 1.5, label: "1.5" },
      { value: 2, label: "2.0" },
      { value: 2.5, label: "2.5" },
      { value: 3, label: "3.0" },
    ];

    /** Word's paragraph-spacing presets, applied with formatParagraph. */
    const SPACING_PRESETS = [
      { id: "none", label: "No spacing", before: 0, after: 0 },
      { id: "tight", label: "Tight", before: 0, after: 4 },
      { id: "normal", label: "Normal", before: 0, after: 8 },
      { id: "open", label: "Open", before: 6, after: 10 },
      { id: "loose", label: "Loose", before: 10, after: 16 },
    ];

    /** Page-size presets in points (A4, Letter, Legal). */
    const PAGE_SIZES = [
      { id: "a4", label: "A4", width: 595.3, height: 841.9 },
      { id: "letter", label: "Letter", width: 612, height: 792 },
      { id: "legal", label: "Legal", width: 612, height: 1008 },
    ];

    /** Margin presets in points (Word's Normal/Narrow/Moderate/Wide). */
    const MARGIN_PRESETS = [
      {
        id: "normal",
        label: "Normal",
        top: 72,
        right: 72,
        bottom: 72,
        left: 72,
      },
      {
        id: "narrow",
        label: "Narrow",
        top: 36,
        right: 36,
        bottom: 36,
        left: 36,
      },
      {
        id: "moderate",
        label: "Moderate",
        top: 72,
        right: 54,
        bottom: 72,
        left: 54,
      },
      { id: "wide", label: "Wide", top: 72, right: 144, bottom: 72, left: 144 },
    ];

    /** Word style gallery: ids resolved against the document style sheet. */
    const STYLE_GALLERY = [
      { id: "Normal", label: "Normal" },
      { id: "NoSpacing", label: "No Spacing" },
      { id: "Title", label: "Title" },
      { id: "Heading1", label: "Heading 1" },
      { id: "Heading2", label: "Heading 2" },
      { id: "Heading3", label: "Heading 3" },
      { id: "Quote", label: "Quote" },
      { id: "ListParagraph", label: "List Paragraph" },
    ];

    const THEME_COLORS = [
      "#000000",
      "#404040",
      "#7f7f7f",
      "#bfbfbf",
      "#ffffff",
      "#c00000",
      "#ff0000",
      "#ffc000",
      "#ffff00",
      "#92d050",
      "#00b050",
      "#00b0f0",
      "#0070c0",
      "#002060",
      "#7030a0",
      "#e7e6e6",
    ];

    /* ------------------------------------------------------------- icons */

    /**
     * Inline SVG icon set. Entries are path `d` strings or geometry shorthand,
     * so no network image, icon font or extra dependency is needed.
     */
    const ICONS = {
      undo: ["M9 7 4 12l5 5", "M4 12h9.5a5.5 5.5 0 0 1 5.5 5.5V19"],
      redo: ["m15 7 5 5-5 5", "M20 12H10.5A5.5 5.5 0 0 0 5 17.5V19"],
      cut: [
        { circle: [6.5, 7, 2.5] },
        { circle: [6.5, 17, 2.5] },
        "m8.5 8.5 11 10M8.5 15.5l11-10",
      ],
      copy: [
        { rect: [8, 8, 12, 12, 1.5] },
        "M16 8V5.5A2.5 2.5 0 0 0 13.5 3H5.5A2.5 2.5 0 0 0 3 5.5v8A2.5 2.5 0 0 0 5.5 16H8",
      ],
      paste: ["M8 4h8v3H8z", "M8 6H5v15h14V6h-3", "M8 11h8M8 15h8"],
      clearFormat: ["M7 5h10M12 5l-3 12", "M5 20h14", "m14 12 7 7M21 12l-7 7"],
      bulletList: [
        { circle: [4.5, 7, 1.2] },
        { circle: [4.5, 12, 1.2] },
        { circle: [4.5, 17, 1.2] },
        "M8 7h12M8 12h12M8 17h12",
      ],
      numberList: [
        "M4 5h2v4M3.5 9h3",
        "M3.5 13h3l-3 3h3",
        "M4 19h2M4 21h2",
        "M9 7h11M9 15h11M9 19h11",
      ],
      multilevelList: ["M4 6h6M4 12h9M4 18h12", "M14 5h6M17 10h3M19 15h1"],
      indentIncrease: ["M4 6h16M4 18h16M10 10h10M10 14h10", "m4 10 4 2-4 2z"],
      indentDecrease: ["M4 6h16M4 18h16M10 10h10M10 14h10", "m8 10-4 2 4 2z"],
      alignLeft: "M4 6h16M4 10h10M4 14h16M4 18h10",
      alignCenter: "M4 6h16M7 10h10M4 14h16M7 18h10",
      alignRight: "M4 6h16M10 10h10M4 14h16M10 18h10",
      alignJustify: "M4 6h16M4 10h16M4 14h16M4 18h16",
      lineSpacing: [
        "M9 6h11M9 12h11M9 18h11",
        "M4 4v16M2 6l2-2 2 2M2 18l2 2 2-2",
      ],
      borders: [{ rect: [4, 4, 16, 16, 1] }, "M4 9h16M4 15h16M9 4v16M15 4v16"],
      shading: [{ rect: [4, 4, 16, 16, 1] }, "M8 8h8v8H8z"],
      styles: ["M4 5h16M4 10h10M4 15h16M4 20h10"],
      find: [{ circle: [10.5, 10.5, 5.5] }, "m15 15 5 5"],
      replace: ["M4 7h12l-3-3M20 17H8l3 3", "M16 7v4M8 17v-4"],
      selectAll: [
        "M4 4h4M16 4h4M4 20h4M16 20h4M4 8v4M4 14v4M20 8v4M20 14v4",
        "M9 9h6v6H9z",
      ],
      pageBreak: [
        "M6 3h7l4.5 4.5V10.5H6z",
        "M13 3v4.5h4.5",
        "M6 13.5h12v3.5H6z",
        "M3 12h3M9 12h3M15 12h3",
      ],
      sectionBreak: [
        "M5 3.5h14v7H5z",
        "M5 13.5h14v7H5z",
        "M3 12h4M10 12h4M17 12h4",
      ],
      columnBreak: ["M4.5 4h6v16h-6z", "M13.5 4h6v16h-6z", "M12 3v18"],
      table: [
        { rect: [3.5, 5, 17, 14, 1.5] },
        "M3.5 9.8h17M3.5 14.5h17M9.6 5v14M15 5v14",
      ],
      image: [
        { rect: [3.5, 5, 17, 14, 1.5] },
        { circle: [8.8, 9.8, 1.5] },
        "m4.6 17.2 4.9-4.9 3.6 3.6 2.5-2.5 4 4",
      ],
      shape: [{ rect: [3, 8, 8, 8, 1] }, { circle: [17, 12, 4.5] }],
      link: [
        "M9.4 14.6 14.6 9.4",
        "M11.1 6.7l1.6-1.6a4 4 0 0 1 5.7 5.7l-1.6 1.6",
        "M12.9 17.3l-1.6 1.6a4 4 0 0 1-5.7-5.7l1.6-1.6",
      ],
      header: [
        "M5 3.5h14v17H5z",
        "M5 9h14",
        "M8.5 6h3.5M8.5 12.5h7M8.5 16h4.5",
      ],
      footer: [
        "M5 3.5h14v17H5z",
        "M5 15h14",
        "M8.5 6h7M8.5 9.5h4.5M8.5 18h3.5",
      ],
      pageNumber: ["M6 3h9l4 4v14H6z", "M15 3v4h4", "M9 12h3M9 16h3M14 12l2 4"],
      toc: [
        "M4.5 6h5M4.5 11h5M4.5 16h5",
        "M13.5 6h6M13.5 11h6M13.5 16h6",
        "M11.2 6h.01M11.2 11h.01M11.2 16h.01",
      ],
      footnote: [
        "M5.5 3.5h8L18 8.5v12H5.5z",
        "M13.5 3.5v5H18",
        "M9 11.5h5.5",
        "M11.8 15.8v3.4M10.3 16.7l3 1.7M13.3 16.7l-3 1.7",
      ],
      endnote: ["M5 3h9l5 5v13H5z", "M14 3v5h5", "M9 12h6M9 16h4"],
      equation: ["M17.5 4.5h-11l6 7.5-6 7.5h11"],
      symbol: ["M6 18V8l6 10 6-10v10", "M4 20h16"],
      pageColor: [{ rect: [4, 3, 16, 18, 1] }, "M4 15h16", "M8 7h8"],
      paragraphSpacing: ["M5 6h14M5 12h14M5 18h14", "M18 4v2M18 10v2M18 16v2"],
      fontSet: ["M4 19 10 5l6 14M6.5 14h7", "M17 19h4"],
      pageBorders: [{ rect: [4, 3, 16, 18, 0] }, { rect: [7, 6, 10, 12, 0] }],
      margins: ["M4 3.5h16v17H4z", "M8 7.5h8v9H8z"],
      orientation: [
        "M4.5 6h9v12h-9z",
        "M17.4 8.4a6 6 0 0 1 0 7.2",
        "m18.6 7.2 2.6 1.2-1.2 2.6",
      ],
      pageSize: [
        "M6 3.5h8.5L19 8v10H6z",
        "M14.5 3.5V8H19",
        "M4 20h16M4 18.5v3M20 18.5v3",
      ],
      columns: ["M4.5 4.5h15v15h-15z", "M9.5 4.5v15", "M14.5 4.5v15"],
      zoomIn: [{ circle: [10.5, 10.5, 5.5] }, "m15 15 5 5M8 10.5h5M10.5 8v5"],
      zoomOut: [{ circle: [10.5, 10.5, 5.5] }, "m15 15 5 5M8 10.5h5"],
      fitWidth: [
        "M4 6v12M20 6v12",
        "M8 12h8M8 12l2-2M8 12l2 2M16 12l-2-2M16 12l-2 2",
      ],
      lock: [{ rect: [5, 10, 14, 10, 1] }, "M8 10V7a4 4 0 0 1 8 0v3"],
      focus: [
        "M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5",
        { circle: [12, 12, 2.5] },
      ],
      ruler: ["M3 8h18v6H3z", "M6 8v3M9 8v2M12 8v3M15 8v2M18 8v3"],
      marks: [
        "M8.5 20V4h3.4a3.6 3.6 0 0 1 0 7.2H8.5",
        "M15.2 20V4",
        "M6 20h12",
      ],
      pageSeparators: ["M4 5h16M4 10h16", "M4 15h16", "M4 19h16"],
      sidebar: [{ rect: [3, 4, 18, 16, 1] }, "M9 4v16"],
      newDoc: ["M6 3h8l4 4v14H6z", "M14 3v4h4M9 13h6M12 10v6"],
      open: [
        "M3.5 8.6V6.4A1.4 1.4 0 0 1 4.9 5h3.9l1.9 2h7.9a1.4 1.4 0 0 1 1.4 1.4v.2",
        "M3.5 8.6h17.4l-1.9 9.3a1.5 1.5 0 0 1-1.5 1.2H5.9a1.5 1.5 0 0 1-1.5-1.2z",
      ],
      save: ["M4 4h13l3 3v13H4z", "M8 4v6h8V4M8 20v-6h8v6"],
      saveAs: [
        "M4 4.5h10.8L18 7.7V17H4z",
        "M8 4.5v5.2h6.2V4.5",
        "m14.5 20.4 5.2-5.2 1.9 1.9-5.2 5.2-2.6.6z",
      ],
      export: ["M12 15V3.6", "m8.3 7.3 3.7-3.7 3.7 3.7", "M5 13.5v6.2h14v-6.2"],
      refresh: ["M20 12a8 8 0 1 1-2.3-5.7", "M20 4v4h-4"],
      trash: ["M4 7h16", "M9 7V4h6v3", "M6 7v13h12V7", "M10 11v6M14 11v6"],
      check: "m5 12 4 4 10-10",
      close: "m6 6 12 12M18 6 6 18",
      plus: "M12 5v14M5 12h14",
      minus: "M5 12h14",
      chevronDown: "m6 9 6 6 6-6",
      chevronUp: "m6 15 6-6 6 6",
      info: [{ circle: [12, 12, 8] }, "M12 11v5M12 8h.01"],
      bookmark: ["M6 3h12v18l-6-4-6 4z"],
      caption: [{ rect: [4, 5, 16, 10, 1] }, "M7 19h10", "M9 9h6M9 12h4"],
      crossReference: [
        "M7 8h4M7 12h4M7 16h4",
        "M15 8h4M15 12h4M15 16h4",
        "M13 5v14",
      ],
      page: ["M6 2h9l5 5v15H6z", "M15 2v5h5"],
      folder: ["M3 7h6l2 2h10v10H3z", "M3 7V5h6l2 2"],
      folderPlus: [
        "M3 7h6l2 2h10v10H3z",
        "M3 7V5h6l2 2",
        "M12 12v5M9.5 14.5h5",
      ],
      up: "M12 19V5M6 11l6-6 6 6",
      labels: ["M4 5h9l7 7-7 7H4z", { circle: [8.5, 12, 1.3] }],
      chat: ["M4 5h16v11H9l-5 4z", "M8 9h8M8 12.5h5"],
      device: [{ rect: [4, 3.5, 16, 12, 1.5] }, "M2.5 19h19", "M10 15.5h4"],
      grid: [
        { rect: [4, 4, 6, 6, 1] },
        { rect: [14, 4, 6, 6, 1] },
        { rect: [4, 14, 6, 6, 1] },
        { rect: [14, 14, 6, 6, 1] },
      ],
      star: ["m12 4 2.4 5 5.6.8-4 3.9 1 5.5-5-2.7-5 2.7 1-5.5-4-3.9 5.6-.8z"],
      // Word's B/I/U/S and the sub/superscript marks are drawn, not typed: a
      // real path set keeps one 24-grid/stroke language across the ribbon.
      bold: [
        "M7 4.5h5.2a3.7 3.7 0 0 1 0 7.4H7z",
        "M7 11.9h5.9a3.9 3.9 0 0 1 0 7.8H7z",
      ],
      italic: ["M10 4.5h7", "M6.5 19.5h7", "M14.2 4.5 9.8 19.5"],
      underline: ["M6.8 4.5v5.8a5.2 5.2 0 0 0 10.4 0V4.5", "M6 19.5h12"],
      strike: [
        "M4.8 12h14.4",
        "M16.2 7.2A4.8 4.8 0 0 0 12 4.6c-2.5 0-4.3 1.3-4.3 3.2 0 1.5 1.1 2.4 3.1 3",
        "M7.8 16.6a4.9 4.9 0 0 0 4.2 2.8c2.6 0 4.4-1.4 4.4-3.3 0-1.4-.9-2.3-2.6-2.9",
      ],
      subscript: [
        "M4.5 4.5l7 8M11.5 4.5l-7 8",
        "M13.6 18.4c0-1.3 1.3-2.2 2.8-2.2s2.8.9 2.8 2.2c0 1.8-2.9 2.6-5.3 5.4h5.3",
      ],
      superscript: [
        "M4.5 4.5l7 8M11.5 4.5l-7 8",
        "M13.6 6.6c0-1.3 1.3-2.2 2.8-2.2s2.8.9 2.8 2.2c0 1.8-2.9 2.6-5.3 5.4h5.3",
      ],
    };

    /**
     * Line icons. They inherit `currentColor`, so a tool's hover/active/pressed
     * colour reaches the strokes, and they default to the ribbon's 22px size
     * with a 1.9 stroke: thick enough to read at 22px in both theme modes, and
     * still inside the 32px tool button with room for the label.
     */
    function Icon(props) {
      const size = typeof props.size === "number" ? props.size : 22;
      const parts = ICONS[props.name] || ICONS.page;
      const list = Array.isArray(parts) ? parts : [parts];
      return h(
        "svg",
        {
          className: props.className ? `dx-icon ${props.className}` : "dx-icon",
          viewBox: "0 0 24 24",
          width: size,
          height: size,
          fill: "none",
          stroke: "currentColor",
          strokeWidth: props.weight || 1.9,
          strokeLinecap: "round",
          strokeLinejoin: "round",
          "aria-hidden": "true",
          focusable: "false",
        },
        list.map((entry, index) => {
          if (typeof entry === "string")
            return h("path", { key: `p${index}`, d: entry });
          if (entry.rect)
            return h("rect", {
              key: `r${index}`,
              x: entry.rect[0],
              y: entry.rect[1],
              width: entry.rect[2],
              height: entry.rect[3],
              rx: entry.rect[4] || 0,
            });
          return h("circle", {
            key: `c${index}`,
            cx: entry.circle[0],
            cy: entry.circle[1],
            r: entry.circle[2],
          });
        }),
      );
    }

    /* --------------------------------------------------------------- CSS */

    /**
     * One stylesheet, rendered as a React `<style>` element by the panel so
     * unmounting the panel removes it. Every class is prefixed `.dx-` and all
     * colours inherit the DSH theme through `var(--dsh-*)` with fallbacks, so
     * the panel is correct in light and dark mode without a skin.
     */
    function styles() {
      return h(
        "style",
        null,
        `
.dx-root{
  /* Palette bridge. Every colour is read from a real DSH theme token first
     (--dsw-alias-*), then from the legacy --dsh-* name, then from a neutral
     fallback so the panel still renders outside the Harness. The --dsw-*
     tokens already switch with the app theme, so no colour below is chosen by
     prefers-color-scheme; only the paper sheet keeps a media query. */
  --dx-text:var(--dsw-alias-label-primary,var(--dsh-text,#1f2328));
  --dx-muted:var(--dsw-alias-label-secondary,var(--dsh-muted,#5f6b7a));
  --dx-border:var(--dsw-alias-border-l1,var(--dsh-border,#d0d7de));
  --dx-border-strong:var(--dsw-alias-border-l2,var(--dsh-border,#d0d7de));
  --dx-surface:var(--dsw-alias-bg-layer-1,var(--dsh-surface,#ffffff));
  --dx-surface-2:var(--dsw-alias-bg-layer-2,var(--dsh-surface,#ffffff));
  --dx-bg:var(--dsw-alias-bg-base,var(--dsh-bg,#f3f2f1));
  --dx-overlay:var(--dsw-alias-bg-overlay,var(--dsh-surface,#ffffff));
  --dx-sidebar:var(--dsw-specific-sidebar-fill,var(--dsh-surface,#ffffff));
  --dx-accent:var(--dsw-alias-brand-primary,var(--dsh-accent,#185abd));
  --dx-on-accent:var(--dsw-alias-label-primary-inverted,#ffffff);
  --dx-hover:var(--dsw-alias-interactive-bg-hover-solid,var(--dsh-hover,rgba(24,90,189,.12)));
  --dx-selected:var(--dsw-alias-interactive-bg-hover-solid,var(--dsh-selected,rgba(24,90,189,.16)));
  --dx-danger:var(--dsw-alias-state-error-primary,var(--dsh-danger,#b42318));
  --dx-success:var(--dsw-alias-state-success-primary,#1a7f37);
  --dx-warn:var(--dsw-alias-state-warn-primary,#9a6700);
  --dx-idle:var(--dsw-alias-state-idle-primary,#8c959f);
  --dx-paper:#ffffff;
  --dx-paper-text:#1b1b1b;
  --dx-paper-border:#d8d6d4;
  --dx-radius:2px;
  --dx-zoom:1;
  display:flex;flex-direction:column;height:100%;min-height:0;position:relative;
  color:var(--dx-text);background:var(--dx-bg);
  font-family:"Segoe UI",var(--dsh-font,system-ui,-apple-system,"Helvetica Neue",Arial,sans-serif);
  font-size:13px;line-height:1.35;overflow:hidden;
}
@media (prefers-color-scheme:dark){
  /* Paper only: the page is a sheet of paper in both themes. The rest of the
     panel follows the theme tokens above. */
  .dx-root{--dx-paper-border:#2a3038;}
}
.dx-root *{box-sizing:border-box;}
.dx-root button{font:inherit;color:inherit;}
.dx-sr{position:absolute;width:1px;height:1px;margin:-1px;padding:0;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap;border:0;}

/* ---------------------------------------------------------- title bar */
.dx-titlebar{
  display:flex;align-items:center;gap:8px;min-height:40px;padding:0 10px;
  background:var(--dx-accent);color:var(--dx-on-accent);flex:0 0 auto;
}
.dx-titlebar .dx-appdot{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600;opacity:.95;}
.dx-title{
  min-width:120px;max-width:32vw;padding:3px 7px;border:1px solid transparent;border-radius:var(--dx-radius);
  background:transparent;color:var(--dx-on-accent);font:inherit;font-size:13px;font-weight:600;text-overflow:ellipsis;
}
.dx-title:hover{border-color:var(--dx-on-accent);}
.dx-title:focus{border-color:var(--dx-on-accent);outline:none;}
.dx-title[disabled]{opacity:.9;cursor:default;}
.dx-titlebar-spacer{flex:1 1 auto;}
.dx-chip{
  display:inline-flex;align-items:center;gap:4px;padding:1px 7px;border-radius:999px;
  border:1px solid var(--dx-on-accent);color:var(--dx-on-accent);font-size:11px;white-space:nowrap;opacity:.92;
}
.dx-chip.dx-chip-warn{background:var(--dx-warn);border-color:transparent;color:var(--dx-on-accent);opacity:1;}
.dx-chip.dx-chip-saved{border-color:var(--dx-on-accent);}
.dx-titlebar .dx-zoom-mini{display:inline-flex;align-items:center;gap:4px;color:var(--dx-on-accent);font-size:11px;}
.dx-titlebar .dx-zoom-mini input[type=range]{width:82px;accent-color:var(--dx-on-accent);}
.dx-titlebar-btn{
  display:inline-flex;align-items:center;justify-content:center;width:30px;height:28px;border:1px solid var(--dx-on-accent);
  border-radius:var(--dx-radius);background:transparent;color:var(--dx-on-accent);cursor:pointer;
}
.dx-titlebar-btn:hover{opacity:.86;}

/* -------------------------------------------------------------- tabs */
.dx-tabrow{
  display:flex;align-items:stretch;gap:2px;flex:0 0 auto;min-height:32px;padding:0 6px;
  background:var(--dx-surface);border-bottom:1px solid var(--dx-border);
}
.dx-tabs{display:flex;align-items:stretch;gap:0;overflow-x:auto;scrollbar-width:none;}
.dx-tab{
  border:0;border-bottom:2px solid transparent;background:transparent;color:var(--dx-text);
  padding:0 11px;font-size:12.5px;white-space:nowrap;cursor:pointer;
}
.dx-tab:hover{background:var(--dx-hover);}
.dx-tab[aria-selected=true]{border-bottom-color:var(--dx-accent);color:var(--dx-accent);font-weight:600;}
.dx-tab:focus-visible{outline:2px solid var(--dx-accent);outline-offset:-2px;}
.dx-ribbon-toggle{
  display:inline-flex;align-items:center;justify-content:center;width:28px;margin-left:auto;
  border:0;background:transparent;color:var(--dx-muted);cursor:pointer;
}
.dx-ribbon-toggle:hover{background:var(--dx-hover);}
.dx-ribbon-file{
  align-self:stretch;padding:0 12px;border:0;background:transparent;color:var(--dx-text);font-size:12.5px;cursor:pointer;
}
.dx-ribbon-file[aria-selected=true]{color:var(--dx-accent);font-weight:600;border-bottom:2px solid var(--dx-accent);}

/* ------------------------------------------------------------ ribbon */
/* Resizable ribbon. The default is 1/10 of the viewport, clamped to a band
   that fits every tab's content; the user can drag it (or use the keyboard on
   the separator) between 64px and half the viewport, and the chosen height is
   stored in this browser. Groups stay on one row and scroll horizontally
   instead of growing taller, so switching tabs never shifts the document, and
   the bottom 6px strip is reserved for the drag handle so nothing hides under
   it (see test/ribbon.test.mjs and the Chromium probe). */
.dx-ribbon-dock{position:relative;display:flex;flex-direction:column;flex:0 0 auto;min-height:0;min-width:0;}
.dx-ribbon-dock[hidden]{display:none;}
.dx-ribbon{
  display:flex;align-items:stretch;gap:0;flex:0 0 auto;min-width:0;
  height:clamp(72px,calc(100vh / 10),240px);min-height:64px;max-height:50vh;padding:0 6px 6px;
  overflow-x:auto;overflow-y:hidden;scrollbar-width:thin;
  background:var(--dx-surface);border-bottom:1px solid var(--dx-border);
}
.dx-ribbon[hidden]{display:none;}
.dx-ribbon::-webkit-scrollbar{height:6px;}
.dx-ribbon::-webkit-scrollbar-thumb{background:var(--dx-border-strong);border-radius:3px;}
/* The strip itself is the separator: 6px tall, aligned with the padding the
   ribbon reserves, so it never covers a caption. */
.dx-ribbon-handle{
  position:absolute;left:0;right:0;bottom:0;z-index:3;display:flex;align-items:center;justify-content:center;
  height:6px;margin:0;padding:0;border:0;background:transparent;color:var(--dx-muted);
  cursor:row-resize;touch-action:none;
}
.dx-ribbon-handle:hover{background:var(--dx-hover);color:var(--dx-accent);}
.dx-ribbon-handle:focus-visible{outline:2px solid var(--dx-accent);outline-offset:-2px;}
.dx-ribbon-grip{display:block;width:30px;height:1px;border-radius:1px;background:currentColor;box-shadow:0 -3px 0 currentColor,0 3px 0 currentColor;opacity:.55;}
.dx-group{
  display:flex;flex-direction:column;justify-content:flex-start;align-items:stretch;
  height:100%;min-height:0;padding:2px 6px 1px;border-right:1px solid var(--dx-border);flex:0 0 auto;
}
.dx-group:last-child{border-right:0;}
.dx-group-items{
  display:flex;align-items:center;gap:2px;flex:1 1 auto;min-height:0;min-width:0;max-width:none;
  flex-wrap:nowrap;overflow-x:auto;overflow-y:hidden;scrollbar-width:thin;
}
.dx-group-items::-webkit-scrollbar{height:5px;}
.dx-group-items::-webkit-scrollbar-thumb{background:var(--dx-border-strong);border-radius:3px;}
.dx-group-items.dx-column{flex-direction:row;flex-wrap:nowrap;max-height:none;}
/* A caption is one 9px line: line-height:1.2 gives the glyph box (ascent plus
   descender) room inside the line box, so p/g descenders are never clipped at
   the 64px floor. */
.dx-group-label{
  margin:0;text-align:center;color:var(--dx-muted);font-size:9px;line-height:1.2;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:0 0 auto;
}
.dx-tool{
  display:inline-flex;align-items:center;justify-content:center;gap:2px;
  width:32px;height:32px;min-width:32px;padding:1px;border:1px solid transparent;border-radius:var(--dx-radius);
  background:transparent;color:var(--dx-text);cursor:pointer;font-size:10.5px;line-height:1.1;flex:0 0 auto;
}
.dx-tool .dx-tool-label{max-width:76px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.dx-tool:hover:not([disabled]){color:var(--dx-accent);background:var(--dx-hover);border-color:var(--dx-border-strong);}
.dx-tool:active:not([disabled]){color:var(--dx-accent);background:var(--dx-selected);border-color:var(--dx-accent);}
.dx-tool[aria-pressed=true]{background:var(--dx-selected);border-color:var(--dx-accent);color:var(--dx-accent);}
/* Disabled stays readable: the icon is dimmed, never washed out. */
.dx-tool[disabled]{opacity:.55;cursor:not-allowed;}
.dx-tool:focus-visible{outline:2px solid var(--dx-accent);outline-offset:-1px;}
.dx-tool.dx-tool-small{height:32px;min-width:32px;font-size:10.5px;}
.dx-tool.dx-tool-tiny{height:32px;min-width:32px;padding:1px;}
/* A colour tool stacks its swatch under the 22px icon inside the 32px button. */
.dx-tool.dx-tool-swatch{flex-direction:column;gap:0;}
/* View → Show labels restores the previous text labels inside the same height. */
.dx-show-labels .dx-tool{
  width:auto;min-width:34px;height:calc(100% - 2px);flex-direction:column;gap:0;padding:1px 4px;
}
.dx-show-labels .dx-tool .dx-tool-label{display:block;font-size:9.5px;line-height:1.2;}
.dx-show-labels .dx-tool.dx-tool-small,
.dx-show-labels .dx-tool.dx-tool-tiny{height:calc(100% - 2px);min-width:34px;flex-direction:row;gap:3px;font-size:10.5px;}
.dx-show-labels .dx-group{height:100%;min-height:0;}
.dx-icon{display:block;flex:0 0 auto;}
.dx-color-bar{display:block;width:18px;height:3px;margin-top:1px;border-radius:1px;flex:0 0 auto;}
.dx-stepper{display:inline-flex;align-items:stretch;height:26px;max-height:100%;border:1px solid var(--dx-border);border-radius:var(--dx-radius);overflow:hidden;background:var(--dx-surface);}
.dx-stepper button{border:0;background:transparent;color:var(--dx-text);width:20px;cursor:pointer;padding:0;}
.dx-stepper button:hover{background:var(--dx-hover);color:var(--dx-accent);}
.dx-stepper input{width:38px;border:0;border-left:1px solid var(--dx-border);border-right:1px solid var(--dx-border);background:transparent;color:var(--dx-text);font:inherit;font-size:12px;text-align:center;padding:0;}
.dx-select{
  height:26px;max-height:100%;max-width:150px;padding:0 4px;border:1px solid var(--dx-border);border-radius:var(--dx-radius);
  background:var(--dx-surface);color:var(--dx-text);font:inherit;font-size:12px;
}
.dx-select:hover{color:var(--dx-accent);border-color:var(--dx-border-strong);}
.dx-gallery{display:flex;align-items:center;gap:3px;align-self:stretch;max-width:330px;max-height:100%;overflow-x:auto;overflow-y:hidden;scrollbar-width:thin;}
.dx-gallery::-webkit-scrollbar{height:5px;}
.dx-gallery::-webkit-scrollbar-thumb{background:var(--dx-border-strong);border-radius:3px;}
.dx-style-card{
  display:inline-flex;align-items:center;justify-content:center;min-width:34px;max-width:52px;height:30px;max-height:100%;
  padding:2px 4px;border:1px solid var(--dx-border);border-radius:var(--dx-radius);
  background:var(--dx-surface);color:var(--dx-text);cursor:pointer;text-align:left;font-size:11px;overflow:hidden;flex:0 0 auto;
}
.dx-style-card:hover{color:var(--dx-accent);background:var(--dx-hover);border-color:var(--dx-border-strong);}
.dx-style-card[aria-pressed=true]{border-color:var(--dx-accent);background:var(--dx-selected);color:var(--dx-accent);}
.dx-style-card b{display:block;font-size:14px;line-height:1.1;font-weight:400;}
.dx-show-labels .dx-style-card{display:block;min-width:74px;max-width:120px;height:54px;max-height:100%;padding:5px 7px;}
.dx-show-labels .dx-style-card b{margin-bottom:2px;}
.dx-ribbon-note{
  align-self:center;max-width:330px;max-height:100%;overflow:hidden;margin:0 10px;color:var(--dx-muted);
  font-size:10.5px;line-height:1.25;
}
.dx-ribbon-note b{color:var(--dx-text);font-weight:600;}

/* ------------------------------------------------------------- layout */
.dx-work{display:flex;flex:1 1 auto;min-height:0;position:relative;}
.dx-rail{
  display:flex;flex-direction:column;gap:6px;width:236px;flex:0 0 236px;min-height:0;
  border-right:1px solid var(--dx-border);background:var(--dx-sidebar);padding:8px;overflow:auto;
}
.dx-rail[hidden]{display:none;}
.dx-rail-head{display:flex;align-items:center;gap:6px;}
.dx-rail-title{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:var(--dx-muted);flex:1 1 auto;}
.dx-rail-input,.dx-input{
  width:100%;padding:4px 7px;border:1px solid var(--dx-border);border-radius:var(--dx-radius);
  background:var(--dx-surface);color:var(--dx-text);font:inherit;font-size:12px;min-width:0;
}
.dx-rail-input:focus,.dx-input:focus{outline:2px solid var(--dx-accent);outline-offset:-1px;}
.dx-rail-row{display:flex;gap:4px;}
.dx-doclist{display:flex;flex-direction:column;gap:2px;overflow:auto;min-height:0;}
.dx-docrow{
  display:flex;flex-direction:column;gap:2px;width:100%;padding:6px 8px;border:1px solid transparent;
  border-radius:var(--dx-radius);background:transparent;color:var(--dx-text);text-align:left;cursor:pointer;font-size:12px;
}
.dx-docrow:hover{background:var(--dx-hover);}
.dx-docrow[aria-current=true]{background:var(--dx-selected);border-color:var(--dx-accent);}
.dx-docrow strong{font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.dx-docrow span{font-size:10.5px;color:var(--dx-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.dx-empty{padding:18px 10px;color:var(--dx-muted);font-size:12px;line-height:1.5;text-align:center;}
.dx-preview{display:flex;justify-content:center;padding:18px 18px 40px;}
.dx-preview-sheet{
  width:100%;max-width:760px;padding:24px 28px;background:var(--dx-paper);color:var(--dx-paper-text);
  border:1px solid var(--dx-paper-border);box-shadow:0 1px 4px rgba(0,0,0,.22);
}
.dx-preview-title{margin:0 0 8px;font-size:17px;font-weight:650;}
.dx-preview-outline{margin:2px 0 10px;padding:0;list-style:none;font-size:12px;color:#444;}
.dx-pre{
  margin:0;padding:10px 12px;max-height:420px;overflow:auto;white-space:pre-wrap;word-wrap:break-word;
  font-family:ui-monospace,Consolas,"Courier New",monospace;font-size:11.5px;line-height:1.5;
  border:1px solid #ddd;background:#fbfaf9;color:#222;
}

.dx-editor{
  position:relative;flex:1 1 auto;min-width:0;min-height:0;overflow:auto;background:var(--dx-bg);
  padding:0 0 42px;
}
.dx-ruler{
  position:sticky;top:0;z-index:4;display:flex;height:20px;background:var(--dx-surface);
  border-bottom:1px solid var(--dx-border);color:var(--dx-muted);font-size:9px;overflow:hidden;
}
.dx-ruler-tick{flex:1 1 0;border-right:1px solid var(--dx-border);text-align:center;line-height:18px;min-width:0;}
.dx-ruler-gutter{width:34px;flex:0 0 34px;border-right:1px solid var(--dx-border);}
.dx-page-stack{
  display:flex;flex-direction:column;align-items:center;gap:18px;padding:16px 18px 0;
  transform-origin:top center;transform:scale(var(--dx-zoom));transition:transform .12s ease-out;
}
.dx-page-wrap{width:100%;max-width:100%;display:flex;flex-direction:column;align-items:center;}
.dx-page-sep{display:flex;align-items:center;gap:8px;color:var(--dx-muted);font-size:10.5px;width:100%;max-width:820px;margin:0 auto 4px;}
.dx-page-sep::before,.dx-page-sep::after{content:"";flex:1 1 auto;border-top:1px dashed var(--dx-border);}
.dx-page{
  width:100%;max-width:760px;min-height:520px;padding:48px 56px 64px;background:var(--dx-paper);
  color:var(--dx-paper-text);border:1px solid var(--dx-paper-border);border-radius:0;
  box-shadow:0 1px 4px rgba(0,0,0,.22),0 8px 20px rgba(0,0,0,.12);
  display:flex;flex-direction:column;gap:2px;position:relative;
}
.dx-page:focus{outline:none;}
.dx-page-head,.dx-page-foot{
  color:#6a6a6a;font-size:9pt;border-bottom:1px solid #e2e0de;padding-bottom:4px;margin-bottom:10px;
  font-family:inherit;
}
.dx-page-foot{border-bottom:0;border-top:1px solid #e2e0de;margin-top:14px;padding-top:4px;}
.dx-page-head[hidden],.dx-page-foot[hidden]{display:none;}

/* ---------------------------------------------------- paragraph + runs */
.dx-para-row{display:flex;align-items:flex-start;gap:6px;position:relative;}
.dx-list-label{
  flex:0 0 auto;min-width:24px;text-align:right;color:var(--dx-paper-text);font-size:inherit;user-select:none;
  white-space:nowrap;
}
.dx-para{
  flex:1 1 auto;min-width:0;min-height:1.35em;white-space:pre-wrap;word-wrap:break-word;
  outline:none;tab-size:8;border:1px solid transparent;
}
.dx-para[contenteditable=true]{cursor:text;}
.dx-para[contenteditable=true]:focus{background:rgba(24,90,189,.03);}
.dx-para.dx-para-empty::before{content:"";display:inline-block;}
.dx-run{white-space:pre-wrap;}
.dx-run-deleted{text-decoration:line-through;color:#b42318;}
.dx-run-inserted{text-decoration:underline;text-decoration-style:dotted;color:#0a6b0a;}
.dx-atom{display:inline-block;vertical-align:baseline;user-select:none;}
/* A line break stays in the inline flow so the br element really ends the line. */
.dx-atom-break{display:inline;}
/* ...and its ↵ marker is only drawn with View → Formatting marks on. */
.dx-marks .dx-atom-break::before{content:"\\21B5";color:#9a9a9a;font-size:.85em;margin-right:1px;}
.dx-atom-pagebreak{display:block;border-top:1px dashed #b9b7b5;margin:6px 0;height:0;}
.dx-atom-tab{display:inline-block;width:2em;}
.dx-img{display:inline-block;max-width:100%;border:1px solid transparent;}
.dx-img:focus{outline:1px solid var(--dx-accent);}
.dx-table{border-collapse:collapse;width:100%;margin:6px 0;table-layout:auto;}
.dx-table td{border:1px solid #b9b7b5;padding:3px 5px;vertical-align:top;min-width:24px;}
.dx-table td .dx-para{min-height:1.2em;}
.dx-toc{border:1px solid #e2e0de;background:#fbfaf9;padding:8px 10px;margin:6px 0;font-size:11.5px;}
.dx-toc-title{font-weight:600;margin-bottom:4px;}
.dx-toc-entry{display:flex;gap:6px;padding:1px 0;color:#333;}
.dx-toc-entry span:last-child{margin-left:auto;color:#777;}
.dx-toc-note{font-size:10px;color:#8a8a8a;margin-top:4px;}

/* formatting marks */
.dx-marks .dx-para::after{content:"\\00B6";color:#b0aeac;font-size:.85em;margin-left:1px;}
.dx-marks .dx-space{position:relative;}
.dx-marks .dx-space::after{content:"\\00B7";position:absolute;left:0;top:-.15em;color:#c3c1bf;font-size:.8em;}
.dx-marks .dx-para:empty::after{content:"\\00B6";}

/* selection / find highlight */
.dx-find-hit{background:rgba(255,214,0,.45);border-radius:1px;}
.dx-find-hit.dx-find-active{background:rgba(255,140,0,.6);}
.dx-link{color:#185abd;text-decoration:underline;cursor:pointer;}

/* --------------------------------------------------------- status bar */
.dx-statusbar{
  display:flex;align-items:center;gap:10px;flex:0 0 auto;min-height:24px;padding:0 10px;
  overflow:hidden;background:var(--dx-surface);border-top:1px solid var(--dx-border);color:var(--dx-muted);
  font-size:11px;white-space:nowrap;
}
.dx-status-item{display:inline-flex;align-items:center;gap:4px;}
.dx-status-spacer{flex:1 1 auto;}
.dx-status-zoom{display:inline-flex;align-items:center;gap:5px;}
.dx-status-zoom input[type=range]{width:86px;accent-color:var(--dx-accent);}
.dx-status-chip{display:inline-flex;align-items:center;gap:4px;padding:0 6px;border-radius:999px;background:var(--dx-hover);}
.dx-status-btn{display:inline-flex;align-items:center;gap:4px;border:0;background:transparent;color:inherit;cursor:pointer;padding:0 3px;font-size:11px;}
.dx-status-btn:hover{color:var(--dx-text);}

/* ------------------------------------------------------ error + dialog */
.dx-error{
  display:flex;align-items:flex-start;gap:8px;flex:0 0 auto;margin:6px 8px;padding:7px 10px;
  border:1px solid var(--dx-danger);border-radius:var(--dx-radius);background:rgba(180,35,24,.10);
  color:var(--dx-danger);font-size:12px;line-height:1.4;
}
.dx-error code{font-family:ui-monospace,Consolas,monospace;font-size:11px;}
.dx-error-body{flex:1 1 auto;min-width:0;}
.dx-error-dismiss{border:0;background:transparent;color:inherit;cursor:pointer;padding:0 2px;}
.dx-notice{
  flex:0 0 auto;margin:6px 8px;padding:6px 10px;border:1px solid var(--dx-border);border-radius:var(--dx-radius);
  background:var(--dx-hover);color:var(--dx-muted);font-size:11.5px;line-height:1.4;
}
.dx-overlay{
  /* Above the File backstage (z-index 40): commands opened from the backstage
     must show their dialog in front of it. */
  position:absolute;inset:0;z-index:60;display:flex;align-items:flex-start;justify-content:center;
  padding-top:64px;background:rgba(20,20,20,.28);overflow:auto;
}
.dx-dialog{
  width:min(440px,92%);display:flex;flex-direction:column;gap:10px;padding:14px 16px 12px;
  background:var(--dx-surface);color:var(--dx-text);border:1px solid var(--dx-border);border-radius:6px;
  box-shadow:0 12px 40px rgba(0,0,0,.3);
}
.dx-dialog h3{margin:0;font-size:14px;font-weight:600;}
.dx-dialog p{margin:0;font-size:12px;color:var(--dx-muted);line-height:1.45;}
.dx-field{display:flex;flex-direction:column;gap:3px;font-size:11.5px;color:var(--dx-muted);}
.dx-field-row{display:flex;gap:8px;}
.dx-field-row .dx-field{flex:1 1 0;}
.dx-dialog-actions{display:flex;justify-content:flex-end;gap:8px;margin-top:2px;}
.dx-btn{
  display:inline-flex;align-items:center;gap:5px;padding:5px 11px;border:1px solid var(--dx-border);
  border-radius:var(--dx-radius);background:var(--dx-surface);color:var(--dx-text);cursor:pointer;font-size:12px;
}
.dx-btn:hover:not([disabled]){background:var(--dx-hover);}
.dx-btn[disabled]{opacity:.45;cursor:not-allowed;}
.dx-btn.dx-primary{background:var(--dx-accent);border-color:transparent;color:var(--dx-on-accent);}
.dx-btn.dx-danger{border-color:var(--dx-danger);color:var(--dx-danger);}
.dx-swatches{display:grid;grid-template-columns:repeat(8,1fr);gap:5px;}
.dx-swatch{width:100%;aspect-ratio:1;border:1px solid var(--dx-border);border-radius:2px;cursor:pointer;padding:0;}
.dx-swatch[aria-pressed=true]{outline:2px solid var(--dx-accent);outline-offset:1px;}
.dx-swatch.dx-swatch-none{background:transparent;color:var(--dx-muted);font-size:11px;}
.dx-find-list{max-height:190px;overflow:auto;display:flex;flex-direction:column;gap:2px;}
.dx-find-item{
  text-align:left;padding:4px 7px;border:1px solid transparent;border-radius:var(--dx-radius);
  background:transparent;color:var(--dx-text);cursor:pointer;font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
}
.dx-find-item:hover{background:var(--dx-hover);}
.dx-find-item[aria-current=true]{background:var(--dx-selected);border-color:var(--dx-accent);}

/* file tab panels */
.dx-file-panels{display:flex;gap:0;flex:1 1 auto;min-height:0;}
/* Compact File commands live in the constant-height ribbon. */
.dx-file-compact{display:flex;align-items:center;gap:4px;height:100%;width:100%;overflow-x:auto;overflow-y:hidden;padding:0 4px;scrollbar-width:thin;}
.dx-file-compact::-webkit-scrollbar{height:6px;}
.dx-file-compact::-webkit-scrollbar-thumb{background:var(--dx-border-strong);border-radius:3px;}
.dx-file-compact .dx-tool{width:32px;height:32px;min-width:32px;}
/* The rich File backstage is an overlay, so it never changes the ribbon height.
   Its top is measured from the ribbon's real bottom by a ResizeObserver and
   published as --dx-ribbon-bottom, so it follows a dragged ribbon exactly. */
.dx-backstage{
  position:absolute;left:0;right:0;top:var(--dx-ribbon-bottom,72px);bottom:0;z-index:40;display:flex;flex-direction:column;
  background:var(--dx-surface);border-top:1px solid var(--dx-border);overflow:auto;
  box-shadow:0 10px 30px rgba(0,0,0,.18);
}
.dx-backstage-head{
  display:flex;align-items:center;gap:8px;flex:0 0 auto;min-height:52px;padding:8px 14px;border-bottom:1px solid var(--dx-border);
}
.dx-backstage-head h3{margin:0;flex:1 1 auto;font-size:14px;font-weight:600;line-height:1.2;}
.dx-backstage .dx-file-panels{flex:1 1 auto;min-height:0;}
.dx-file-col{display:flex;flex-direction:column;gap:8px;padding:12px 14px;border-right:1px solid var(--dx-border);min-width:230px;}
.dx-file-col h4{margin:0;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:var(--dx-muted);}
.dx-file-actions{display:flex;flex-direction:column;gap:4px;}
.dx-file-action{
  display:flex;align-items:center;gap:8px;padding:5px 8px;border:1px solid transparent;border-radius:var(--dx-radius);
  background:transparent;color:var(--dx-text);cursor:pointer;font-size:12px;text-align:left;
}
.dx-file-action:hover:not([disabled]){background:var(--dx-hover);border-color:var(--dx-border);}
.dx-file-action[disabled]{opacity:.45;cursor:not-allowed;}
.dx-file-meta{font-size:11.5px;color:var(--dx-muted);line-height:1.5;}
.dx-file-meta b{color:var(--dx-text);font-weight:600;}
.dx-kv{display:grid;grid-template-columns:auto 1fr;gap:2px 10px;font-size:11.5px;color:var(--dx-muted);}
.dx-kv b{color:var(--dx-text);font-weight:500;}

/* workspace + chat session chip */
.dx-sessionbar{
  display:flex;align-items:center;gap:8px;flex:0 0 auto;padding:4px 10px;min-height:30px;
  background:var(--dx-hover);border-bottom:1px solid var(--dx-border);color:var(--dx-text);font-size:12px;
}
.dx-sessionbar[data-status=no-harness-workspace]{border-bottom-color:var(--dx-danger);}
.dx-session-icon{color:var(--dx-accent);}
.dx-sessionbar[data-status=no-harness-workspace] .dx-session-icon{color:var(--dx-danger);}
.dx-session-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:600;}
.dx-session-warn{color:var(--dx-danger);font-size:11.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.dx-session-spacer{flex:1 1 auto;}

/* File tab: advanced path disclosure + hidden native pickers */
.dx-file-input{display:none;}
.dx-advanced{margin-top:2px;font-size:11.5px;color:var(--dx-muted);}
.dx-advanced summary{cursor:pointer;padding:3px 0;font-size:11.5px;color:var(--dx-muted);}
.dx-advanced summary:hover{color:var(--dx-text);}
.dx-advanced-body{display:flex;flex-direction:column;gap:5px;padding-top:4px;}
.dx-roots{display:flex;flex-wrap:wrap;gap:4px;}
.dx-root-link{
  border:1px solid var(--dx-border);background:var(--dx-surface);color:var(--dx-text);
  border-radius:999px;padding:2px 8px;font-size:11px;cursor:pointer;max-width:160px;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
}
.dx-root-link:hover{background:var(--dx-hover);}

/* in-panel folder browser */
.dx-dialog.dx-dialog-wide{width:min(580px,94%);max-height:82vh;overflow:auto;}
.dx-browse{display:flex;flex-direction:column;gap:8px;min-height:0;}
.dx-browse-bar{display:flex;align-items:center;gap:6px;flex-wrap:wrap;}
.dx-breadcrumbs{
  display:flex;align-items:center;gap:1px;flex-wrap:wrap;flex:1 1 160px;min-width:0;
  font-size:11.5px;color:var(--dx-muted);
}
.dx-crumb{
  border:0;background:transparent;color:var(--dx-accent);cursor:pointer;padding:1px 3px;
  font-size:11.5px;border-radius:var(--dx-radius);max-width:160px;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
}
.dx-crumb:hover{background:var(--dx-hover);}
.dx-crumb[aria-current=page]{color:var(--dx-text);font-weight:600;}
.dx-crumb-sep{opacity:.55;}
.dx-browse-list{
  display:flex;flex-direction:column;gap:1px;max-height:250px;overflow:auto;
  border:1px solid var(--dx-border);border-radius:var(--dx-radius);padding:2px;
}
.dx-browse-item{
  display:flex;align-items:center;gap:7px;width:100%;padding:5px 7px;border:0;border-radius:var(--dx-radius);
  background:transparent;color:var(--dx-text);cursor:pointer;font-size:12px;text-align:left;
}
.dx-browse-item:hover:not([disabled]){background:var(--dx-hover);}
.dx-browse-item[disabled]{opacity:.5;cursor:not-allowed;}
.dx-browse-item:focus-visible{outline:2px solid var(--dx-accent);outline-offset:-1px;}
.dx-browse-item span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.dx-browse-meta{margin-left:auto;color:var(--dx-muted);font-size:10.5px;flex:0 0 auto;}
.dx-browse-empty{padding:10px;color:var(--dx-muted);font-size:11.5px;text-align:center;}
.dx-browse-error{margin:0;color:var(--dx-danger);font-size:11.5px;line-height:1.4;}
.dx-browse-toolbar{display:flex;gap:5px;align-items:center;flex-wrap:wrap;}
.dx-newfolder{display:flex;gap:4px;align-items:center;}

@media (max-width:820px){
  .dx-rail{display:none;}
  .dx-page{padding:28px 22px 40px;}
  .dx-page-stack{padding:10px 8px 0;}
  .dx-title{max-width:40vw;}
}
@media (prefers-reduced-motion:reduce){.dx-page-stack{transition:none;}}
`,
      );
    }

    /* --------------------------------------------------- pure helpers */

    /*
     * The helpers below mirror the editor API contract without importing the
     * Host package: the client module may only require React. They are pure and
     * are exported through `__test` so Node can test them directly.
     */

    const HIGHLIGHT_NAMES = {
      yellow: "#ffff00",
      green: "#00ff00",
      cyan: "#00ffff",
      magenta: "#ff00ff",
      blue: "#0000ff",
      red: "#ff0000",
      darkblue: "#000080",
      darkcyan: "#008080",
      darkgreen: "#008000",
      darkmagenta: "#800080",
      darkred: "#800000",
      darkyellow: "#808000",
      darkgray: "#808080",
      lightgray: "#c0c0c0",
      black: "#000000",
      white: "#ffffff",
    };

    /**
     * Plain-text projection of one inline in the editor's offset space.
     *
     * The Host resolves `{ blockId, offset }` with the engine's
     * `inlineTextLength`, which counts text runs even when they carry a tracked
     * deletion. The client must count the same characters or every offset would
     * drift on a document with pending tracked changes.
     */
    function inlineText(inline) {
      if (!inline || typeof inline !== "object") return "";
      switch (inline.type) {
        case "text":
          return typeof inline.text === "string" ? inline.text : "";
        case "tab":
          return "\t";
        case "break":
          return inline.kind === "page" ? "\f" : "\n";
        case "image":
          return "\uFFFC";
        case "field":
          return Array.isArray(inline.result)
            ? inline.result.map((child) => inlineText(child)).join("")
            : "";
        case "math":
          return typeof inline.linear === "string" ? inline.linear : "";
        case "unknown":
          return typeof inline.text === "string" ? inline.text : "";
        default:
          return "";
      }
    }

    /** Flattened paragraph text in the offset space used by every op range. */
    function paragraphText(paragraph) {
      if (!paragraph || !Array.isArray(paragraph.inlines)) return "";
      return paragraph.inlines.map((inline) => inlineText(inline)).join("");
    }

    /**
     * Final-mode text for statistics and the outline: tracked deletions and
     * inserted-revision paragraphs do not count, matching the engine's
     * `documentStats` projection.
     */
    function finalInlineText(inline) {
      if (!inline || typeof inline !== "object") return "";
      if (
        inline.type === "text" &&
        inline.props &&
        inline.props.revision &&
        inline.props.revision.kind === "delete"
      )
        return "";
      return inlineText(inline);
    }

    function finalParagraphText(paragraph) {
      if (!paragraph || !Array.isArray(paragraph.inlines)) return "";
      if (paragraph.revision && paragraph.revision.kind === "insert") return "";
      return paragraph.inlines
        .map((inline) => finalInlineText(inline))
        .join("");
    }

    /** Every inline with the half-open text range it occupies. */
    function flatInlines(paragraph) {
      const spans = [];
      if (!paragraph || !Array.isArray(paragraph.inlines)) return spans;
      let position = 0;
      paragraph.inlines.forEach((inline, index) => {
        const text = inlineText(inline);
        spans.push({
          index,
          inline,
          text,
          start: position,
          end: position + text.length,
        });
        position += text.length;
      });
      return spans;
    }

    function clampOffset(offset, length) {
      const value =
        typeof offset === "number" && Number.isFinite(offset) ? offset : 0;
      if (value <= 0) return 0;
      if (!Number.isFinite(length) || length <= 0) return 0;
      return value > length ? length : Math.floor(value);
    }

    /** The inline span containing `offset`, or the last span for the end offset. */
    function spanAtOffset(paragraph, offset) {
      const spans = flatInlines(paragraph);
      const total = spans.length ? spans[spans.length - 1].end : 0;
      const target = clampOffset(offset, total);
      let found = null;
      for (const span of spans) {
        if (span.start <= target && target <= span.end) found = span;
        if (span.start > target) break;
      }
      return found;
    }

    /** Direct run props at an offset, used when inserting new text. */
    function runPropsAtOffset(paragraph, offset) {
      const span = spanAtOffset(paragraph, offset);
      if (!span || !span.inline || !span.inline.props) return {};
      const props = span.inline.props;
      const copy = {};
      for (const key of Object.keys(props)) {
        if (key === "revision" || key === "hyperlink") continue;
        copy[key] = props[key];
      }
      return copy;
    }

    /**
     * Minimal diff between two paragraph texts: the replaced range in `next`'s
     * coordinate space on the old text plus the inserted text.
     * Returns `null` when the texts are identical.
     */
    function diffTextRange(previous, next) {
      const before = typeof previous === "string" ? previous : "";
      const after = typeof next === "string" ? next : "";
      if (before === after) return null;
      let prefix = 0;
      const limit = Math.min(before.length, after.length);
      while (prefix < limit && before[prefix] === after[prefix]) prefix += 1;
      let suffix = 0;
      const maxSuffix = Math.min(before.length, after.length) - prefix;
      while (
        suffix < maxSuffix &&
        before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
      )
        suffix += 1;
      return {
        start: prefix,
        end: before.length - suffix,
        insert: after.slice(prefix, after.length - suffix),
      };
    }

    function stableRunProps(props) {
      if (!props) return "";
      const keys = Object.keys(props).sort();
      return keys
        .map((key) => {
          const value = props[key];
          if (value && typeof value === "object")
            return `${key}:${JSON.stringify(value)}`;
          return `${key}:${String(value)}`;
        })
        .join("|");
    }

    function mergeTextRuns(inlines) {
      const merged = [];
      for (const inline of inlines) {
        if (!inline) continue;
        if (inline.type === "text" && !inline.text) continue;
        const previous = merged[merged.length - 1];
        if (
          previous &&
          previous.type === "text" &&
          inline.type === "text" &&
          stableRunProps(previous.props) === stableRunProps(inline.props)
        ) {
          merged[merged.length - 1] = {
            type: "text",
            text: previous.text + inline.text,
            props: previous.props || {},
          };
          continue;
        }
        merged.push(inline);
      }
      if (merged.length === 0)
        merged.push({ type: "text", text: "", props: {} });
      return merged;
    }

    /**
     * Apply a text replacement to a paragraph model in place, preserving run
     * formatting outside the replaced range and inheriting the formatting at the
     * insertion point for the new text. Returns the paragraph.
     */
    function applyTextReplace(paragraph, start, end, text) {
      if (!paragraph) return paragraph;
      const spans = flatInlines(paragraph);
      const total = spans.length ? spans[spans.length - 1].end : 0;
      const from = clampOffset(start, total);
      const to = Math.max(from, clampOffset(end, total));
      const inserted = typeof text === "string" ? text : "";
      const boundary = runPropsAtOffset(paragraph, from);
      const result = [];
      let placed = false;
      const place = () => {
        if (placed) return;
        placed = true;
        if (inserted)
          result.push({ type: "text", text: inserted, props: boundary });
      };
      for (const span of spans) {
        const inline = span.inline;
        if (!placed && span.start >= from) place();
        if (span.end <= from || span.start >= to) {
          result.push(inline);
          continue;
        }
        if (inline.type !== "text") continue;
        const localFrom = Math.max(0, from - span.start);
        const localTo = Math.max(
          localFrom,
          Math.min(inline.text.length, to - span.start),
        );
        const prefix = inline.text.slice(0, localFrom);
        const suffix = inline.text.slice(localTo);
        if (prefix)
          result.push({
            type: "text",
            text: prefix,
            props: inline.props || {},
          });
        place();
        if (suffix)
          result.push({
            type: "text",
            text: suffix,
            props: inline.props || {},
          });
      }
      place();
      paragraph.inlines = mergeTextRuns(result);
      return paragraph;
    }

    function findStyle(document, styleId) {
      const table =
        (document && document.styles && document.styles.styles) || {};
      if (!styleId) return null;
      if (table[styleId]) return table[styleId];
      const wanted = String(styleId).toLowerCase().replace(/\s+/g, "");
      for (const key of Object.keys(table)) {
        const style = table[key];
        if (!style) continue;
        if (String(key).toLowerCase() === String(styleId).toLowerCase())
          return style;
        if (
          style.name &&
          String(style.name).toLowerCase().replace(/\s+/g, "") === wanted
        )
          return style;
      }
      return null;
    }

    function styleKey(document, style) {
      const table =
        (document && document.styles && document.styles.styles) || {};
      for (const key of Object.keys(table)) {
        if (table[key] === style) return key;
      }
      return style && style.id ? style.id : "";
    }

    function styleChain(document, styleId) {
      const chain = [];
      let current = findStyle(document, styleId);
      let guard = 0;
      while (current && guard < 24) {
        chain.unshift(current);
        current = current.basedOn ? findStyle(document, current.basedOn) : null;
        guard += 1;
      }
      return chain;
    }

    function shallowMerge(base, patch) {
      if (!patch) return base || {};
      const out = Object.assign({}, base || {});
      for (const key of Object.keys(patch)) {
        const value = patch[key];
        if (value === undefined) continue;
        if (value && typeof value === "object" && !Array.isArray(value)) {
          out[key] = Object.assign({}, out[key] || {}, value);
        } else {
          out[key] = value;
        }
      }
      return out;
    }

    /** Effective run formatting for one inline (defaults → styles → direct). */
    function resolveRunProps(document, paragraph, inline) {
      const defaults =
        (document && document.styles && document.styles.defaults
          ? document.styles.defaults.run
          : null) || {};
      let merged = Object.assign({}, defaults);
      for (const style of styleChain(document, paragraph && paragraph.style))
        merged = shallowMerge(merged, style.run);
      const direct =
        (inline && inline.props) ||
        (paragraph && paragraph.props && paragraph.props.run) ||
        {};
      const characterStyle =
        direct.characterStyle && findStyle(document, direct.characterStyle);
      if (characterStyle) merged = shallowMerge(merged, characterStyle.run);
      merged = shallowMerge(
        merged,
        Object.assign({}, direct, {
          revision: undefined,
          characterStyle: undefined,
        }),
      );
      return merged;
    }

    /** Effective paragraph formatting for one paragraph (defaults → styles → direct). */
    function resolveParagraphProps(document, paragraph) {
      const defaults =
        (document && document.styles && document.styles.defaults
          ? document.styles.defaults.paragraph
          : null) || {};
      let merged = Object.assign({}, defaults);
      for (const style of styleChain(document, paragraph && paragraph.style))
        merged = shallowMerge(merged, style.paragraph);
      merged = shallowMerge(merged, paragraph && paragraph.props);
      return merged;
    }

    function round(value, places) {
      const factor = Math.pow(10, places === undefined ? 1 : places);
      return Math.round(value * factor) / factor;
    }

    function cssFamily(family) {
      const name = String(family).trim();
      if (!name) return undefined;
      if (/[",]/.test(name)) return name;
      return `"${name}", "${name.replace(/\s+/g, "")}", sans-serif`;
    }

    function cssHighlight(value) {
      if (!value) return undefined;
      const named = HIGHLIGHT_NAMES[String(value).toLowerCase()];
      return named || value;
    }

    /** CSS for one run's effective character formatting. */
    function runStyleToCss(props) {
      const css = {};
      if (!props) return css;
      if (props.fontFamily) css.fontFamily = cssFamily(props.fontFamily);
      let fontSize = typeof props.fontSize === "number" ? props.fontSize : null;
      if (
        props.vertAlign === "superscript" ||
        props.vertAlign === "subscript"
      ) {
        fontSize = fontSize ? fontSize * 0.72 : 0.72 * 11;
        css.verticalAlign = props.vertAlign === "superscript" ? "super" : "sub";
      }
      if (fontSize) css.fontSize = `${round(fontSize, 1)}pt`;
      if (props.bold) css.fontWeight = 700;
      if (props.italic) css.fontStyle = "italic";
      const decorations = [];
      if (props.underline && props.underline !== "none")
        decorations.push("underline");
      if (props.strike || props.doubleStrike) decorations.push("line-through");
      if (decorations.length) {
        css.textDecorationLine = decorations.join(" ");
        if (props.underlineColor)
          css.textDecorationColor = props.underlineColor;
      }
      if (props.color) css.color = props.color;
      if (props.highlight) css.backgroundColor = cssHighlight(props.highlight);
      if (props.background) css.backgroundColor = props.background;
      if (props.caps) css.textTransform = "uppercase";
      if (props.smallCaps) css.fontVariant = "small-caps";
      if (typeof props.charSpacing === "number" && props.charSpacing)
        css.letterSpacing = `${round(props.charSpacing, 1)}pt`;
      if (typeof props.charScale === "number" && props.charScale !== 100)
        css.transform = `scaleX(${round(props.charScale / 100, 3)})`;
      return css;
    }

    function borderCss(border) {
      if (!border || border.none || !border.style || border.style === "none")
        return null;
      const style =
        border.style === "double"
          ? "double"
          : border.style === "dotted"
            ? "dotted"
            : border.style === "dashed" ||
                border.style === "dotDash" ||
                border.style === "dotDotDash"
              ? "dashed"
              : border.style === "thick"
                ? "solid"
                : "solid";
      const width = border.style === "thick" ? 2 : 1;
      return `${width}px ${style} ${border.color || "#b9b7b5"}`;
    }

    /** CSS for one paragraph's effective paragraph formatting. */
    function paragraphStyleToCss(props) {
      const css = {};
      if (!props) return css;
      const alignment =
        props.alignment === "distribute" ? "justify" : props.alignment;
      if (alignment) css.textAlign = alignment;
      const spacing = props.lineSpacing;
      if (spacing && typeof spacing.value === "number") {
        css.lineHeight =
          spacing.rule === "exact" || spacing.rule === "atLeast"
            ? `${round(spacing.value, 1)}pt`
            : round(spacing.value / 240, 3);
      }
      if (typeof props.spaceBefore === "number")
        css.marginTop = `${round(props.spaceBefore, 1)}pt`;
      if (typeof props.spaceAfter === "number")
        css.marginBottom = `${round(props.spaceAfter, 1)}pt`;
      const indent = props.indent;
      if (indent) {
        if (typeof indent.left === "number")
          css.marginLeft = `${round(indent.left, 1)}pt`;
        if (typeof indent.right === "number")
          css.marginRight = `${round(indent.right, 1)}pt`;
        if (typeof indent.firstLine === "number")
          css.textIndent = `${round(indent.firstLine, 1)}pt`;
        if (typeof indent.hanging === "number")
          css.textIndent = `${round(-indent.hanging, 1)}pt`;
      }
      if (props.shading && props.shading.fill)
        css.backgroundColor = props.shading.fill;
      const borders = props.borders;
      if (borders) {
        const top = borderCss(borders.top);
        const bottom = borderCss(borders.bottom);
        const left = borderCss(borders.left);
        const right = borderCss(borders.right);
        if (top) css.borderTop = top;
        if (bottom) css.borderBottom = bottom;
        if (left) css.borderLeft = left;
        if (right) css.borderRight = right;
      }
      return css;
    }

    /** Allow only http, https and mailto URLs (via the WHATWG URL parser). */
    function sanitizeUrl(raw) {
      if (typeof raw !== "string") return null;
      const text = raw.trim();
      if (!text) return null;
      let url;
      try {
        url = new URL(text);
      } catch (_error) {
        return null;
      }
      const protocol = url.protocol.toLowerCase();
      if (
        protocol !== "http:" &&
        protocol !== "https:" &&
        protocol !== "mailto:"
      )
        return null;
      return url.href;
    }

    const SAFE_IMAGE_TYPES = [
      "image/png",
      "image/jpeg",
      "image/jpg",
      "image/gif",
      "image/webp",
      "image/bmp",
    ];

    /** Build a `data:` URL for a model media asset, refusing anything else. */
    function mediaUrl(asset) {
      if (!asset || typeof asset.data !== "string" || !asset.data) return null;
      const contentType = String(
        asset.contentType || "image/png",
      ).toLowerCase();
      if (SAFE_IMAGE_TYPES.indexOf(contentType) === -1) return null;
      if (!/^[A-Za-z0-9+/=\s]+$/.test(asset.data)) return null;
      return `data:${contentType};base64,${asset.data.replace(/\s+/g, "")}`;
    }

    function countWords(text) {
      if (!text) return 0;
      const matches = String(text).match(
        /[\p{L}\p{N}\p{M}]+(?:['\u2019-][\p{L}\p{N}\p{M}]+)*/gu,
      );
      return matches ? matches.length : 0;
    }

    /** Walk every block in document order, descending into table cells. */
    function walkDocumentBlocks(document, visit) {
      const sections = (document && document.sections) || [];
      const walkBlocks = (blocks, sectionIndex, depth) => {
        for (const block of blocks || []) {
          visit({ block, sectionIndex, depth });
          if (block && block.type === "table") {
            const rows = block.rows || [];
            for (const row of rows) {
              for (const cell of row.cells || [])
                walkBlocks(cell.blocks || [], sectionIndex, depth + 1);
            }
          }
        }
      };
      sections.forEach((section, sectionIndex) =>
        walkBlocks(section.blocks || [], sectionIndex, 0),
      );
    }

    /** Flat ordered block list plus id lookup (the editor's addressing space). */
    function collectBlocks(document) {
      const list = [];
      const byId = new Map();
      walkDocumentBlocks(document, (entry) => {
        list.push(entry);
        if (entry.block && entry.block.id) byId.set(entry.block.id, entry);
      });
      return { list, byId };
    }

    function computeStats(document) {
      const stats = {
        words: 0,
        characters: 0,
        charactersWithSpaces: 0,
        paragraphs: 0,
        tables: 0,
        images: 0,
      };
      walkDocumentBlocks(document, ({ block }) => {
        if (!block) return;
        if (block.type === "table") {
          stats.tables += 1;
          return;
        }
        if (block.type !== "paragraph") return;
        stats.paragraphs += 1;
        const text = finalParagraphText(block);
        stats.words += countWords(text);
        for (const character of text) {
          if (character === "\uFFFC") continue;
          stats.charactersWithSpaces += 1;
          if (!/\s/.test(character)) stats.characters += 1;
        }
        for (const inline of block.inlines || []) {
          if (inline && inline.type === "image") stats.images += 1;
        }
      });
      return stats;
    }

    function headingLevel(paragraph) {
      const props = paragraph && paragraph.props ? paragraph.props : {};
      if (typeof props.outlineLevel === "number" && props.outlineLevel >= 0)
        return Math.min(9, props.outlineLevel + 1);
      const style = String((paragraph && paragraph.style) || "")
        .toLowerCase()
        .replace(/\s+/g, "");
      if (/^title$/.test(style)) return 1;
      const match = /^heading([1-9])$/.exec(style);
      if (match) return Number(match[1]);
      return 0;
    }

    function computeOutline(document) {
      const outline = [];
      walkDocumentBlocks(document, ({ block }) => {
        if (!block || block.type !== "paragraph") return;
        const level = headingLevel(block);
        if (!level) return;
        outline.push({
          level,
          text: finalParagraphText(block).trim(),
          blockId: block.id,
          style: block.style,
        });
      });
      return outline;
    }

    /** All body blocks (sections in order), without table-cell children. */
    function documentBodyBlocks(document) {
      const sections = (document && document.sections) || [];
      const blocks = [];
      for (const section of sections) blocks.push(...(section.blocks || []));
      return blocks;
    }

    /**
     * Group body blocks into 1-based pages using the Host page map. When the map
     * is missing every block lands on page 1 (a single continuous sheet).
     */
    function groupBlocksIntoPages(blocks, blockPages) {
      const pages = [];
      let current = null;
      for (const block of blocks || []) {
        const mapped = blockPages && block ? blockPages[block.id] : null;
        const page =
          typeof mapped === "number" && mapped > 0
            ? mapped
            : current
              ? current.page
              : 1;
        if (!current || current.page !== page) {
          current = { page, blocks: [] };
          pages.push(current);
        }
        current.blocks.push(block);
      }
      if (pages.length === 0) pages.push({ page: 1, blocks: [] });
      return pages;
    }

    function padNumber(value) {
      return value < 10 ? `0${value}` : String(value);
    }

    function romanNumber(value) {
      const table = [
        [1000, "M"],
        [900, "CM"],
        [500, "D"],
        [400, "CD"],
        [100, "C"],
        [90, "XC"],
        [50, "L"],
        [40, "XL"],
        [10, "X"],
        [9, "IX"],
        [5, "V"],
        [4, "IV"],
        [1, "I"],
      ];
      let remaining = value;
      let out = "";
      for (const [amount, glyph] of table) {
        while (remaining >= amount) {
          out += glyph;
          remaining -= amount;
        }
      }
      return out;
    }

    function letterNumber(value) {
      let remaining = Math.max(1, value);
      let out = "";
      while (remaining > 0 && out.length < 3) {
        const digit = (remaining - 1) % 26;
        out = String.fromCharCode(65 + digit) + out;
        remaining = Math.floor((remaining - 1) / 26);
      }
      return out;
    }

    function formatCounter(value, numFmt) {
      switch (numFmt) {
        case "none":
          return "";
        case "decimalZero":
          return padNumber(value);
        case "lowerLetter":
          return letterNumber(value).toLowerCase();
        case "upperLetter":
          return letterNumber(value);
        case "lowerRoman":
          return romanNumber(value).toLowerCase();
        case "upperRoman":
          return romanNumber(value);
        case "decimalFullWidth":
          return String(value);
        default:
          return String(value);
      }
    }

    /**
     * List labels for every numbered paragraph, in document order. Counters are
     * approximated per numId/level, which is what the engine computes too.
     */
    function computeListLabels(document) {
      const labels = new Map();
      const numbering = (document && document.numbering) || {
        nums: {},
        abstract: {},
      };
      const counters = new Map();
      walkDocumentBlocks(document, ({ block }) => {
        if (!block || block.type !== "paragraph") return;
        const props = resolveParagraphProps(document, block);
        const ref = props.numbering;
        if (!ref || !ref.numId) return;
        const instance = numbering.nums ? numbering.nums[ref.numId] : null;
        const abstract =
          instance && numbering.abstract
            ? numbering.abstract[instance.abstractId]
            : null;
        const level = Math.max(0, Math.min(8, ref.level || 0));
        const levels = (abstract && abstract.levels) || [];
        const levelDef = levels.find((entry) => entry.level === level) || null;
        const state = counters.get(ref.numId) || [];
        const start =
          ref.startOverride ||
          (levelDef && typeof levelDef.start === "number" ? levelDef.start : 1);
        if (!state[level]) state[level] = Math.max(0, start - 1);
        state[level] += 1;
        for (let deeper = level + 1; deeper < 9; deeper += 1) state[deeper] = 0;
        counters.set(ref.numId, state);
        if (ref.labelOverride) {
          labels.set(block.id, { text: ref.labelOverride, level });
          return;
        }
        const lvlText =
          (levelDef && levelDef.lvlText) ||
          ref.startOverride ||
          `%${level + 1}.`;
        const numFmt =
          (levelDef && levelDef.numFmt) ||
          (level === 0 ? "decimal" : "decimal");
        if (numFmt === "bullet") {
          labels.set(block.id, { text: String(lvlText || "\u2022"), level });
          return;
        }
        const text = String(lvlText).replace(/%([1-9])/g, (_match, digit) => {
          const index = Number(digit) - 1;
          return formatCounter(state[index] || 0, numFmt);
        });
        labels.set(block.id, { text, level });
      });
      return labels;
    }

    function fontFamiliesFor(document) {
      const families = [];
      const add = (family) => {
        const name = typeof family === "string" ? family.trim() : "";
        if (name && families.indexOf(name) === -1) families.push(name);
      };
      FONT_FALLBACKS.forEach(add);
      const theme = document && document.theme ? document.theme.fonts : null;
      if (theme) {
        if (theme.major && theme.major.latin) add(theme.major.latin);
        if (theme.minor && theme.minor.latin) add(theme.minor.latin);
      }
      const styles =
        (document && document.styles && document.styles.styles) || {};
      for (const key of Object.keys(styles)) {
        const style = styles[key];
        if (style && style.run && style.run.fontFamily)
          add(style.run.fontFamily);
      }
      const defaults =
        document &&
        document.styles &&
        document.styles.defaults &&
        document.styles.defaults.run;
      if (defaults && defaults.fontFamily) add(defaults.fontFamily);
      return families;
    }

    /** The Word style gallery resolved against this document's style sheet. */
    function styleGalleryFor(document) {
      return STYLE_GALLERY.map((entry) => {
        const style = findStyle(document, entry.id);
        return {
          id: style ? styleKey(document, style) : entry.id,
          label: (style && style.name) || entry.label,
          available: Boolean(style),
        };
      });
    }

    /** Live page-size readout for the Layout tab, in points. */
    function pageSizeLabel(section) {
      const page = section && section.page;
      if (!page) return "Page size unknown";
      const orientation =
        page.orientation === "landscape" ? "landscape" : "portrait";
      return `${round(page.width, 1)} \u00D7 ${round(page.height, 1)} pt (${orientation})`;
    }

    function marginLabel(section) {
      const page = section && section.page;
      if (!page || !page.margins) return "Margins unknown";
      const { top, right, bottom, left } = page.margins;
      return `${round(top, 1)} / ${round(right, 1)} / ${round(bottom, 1)} / ${round(left, 1)} pt`;
    }

    function pageNumberForBlock(blockPages, blockId) {
      if (!blockId) return 1;
      if (blockPages && typeof blockPages[blockId] === "number")
        return blockPages[blockId];
      return 1;
    }

    function formatZoom(zoom) {
      return `${Math.round(zoom * 100)}%`;
    }

    /**
     * New block id created by a splitting op (`insertText` with "\n").
     *
     * The contract says the op result carries the created paragraph ids; the
     * Host currently answers `{ blockIds: ["p1", "p2"], start, end }`. Accept
     * every documented spelling and fall back to `null` so the caller can
     * refetch and find the block after the current one instead.
     */
    function newBlockIdFromOpResult(opResult, currentBlockId) {
      if (!opResult || typeof opResult !== "object") return null;
      const raw =
        opResult.result && typeof opResult.result === "object"
          ? opResult.result
          : opResult;
      const candidates = [];
      for (const key of ["paragraphIds", "blockIds", "createdBlocks"]) {
        if (Array.isArray(raw[key])) {
          for (const value of raw[key]) candidates.push(value);
        }
      }
      const usable = candidates.filter(
        (value) =>
          typeof value === "string" && value && value !== currentBlockId,
      );
      return usable.length ? usable[usable.length - 1] : null;
    }

    function normalizeSelection(selection, order) {
      if (!selection || !selection.start || !selection.end) return null;
      const indexOf = (blockId) =>
        order && order.has(blockId)
          ? order.get(blockId)
          : Number.MAX_SAFE_INTEGER;
      const startIndex = indexOf(selection.start.blockId);
      const endIndex = indexOf(selection.end.blockId);
      if (startIndex > endIndex)
        return { start: selection.end, end: selection.start };
      if (startIndex === endIndex) {
        const startOffset = selection.start.offset || 0;
        const endOffset = selection.end.offset || 0;
        if (startOffset > endOffset)
          return { start: selection.end, end: selection.start };
      }
      return { start: selection.start, end: selection.end };
    }

    function isCollapsed(selection) {
      return Boolean(
        selection &&
        selection.start &&
        selection.end &&
        selection.start.blockId === selection.end.blockId &&
        (selection.start.offset || 0) === (selection.end.offset || 0),
      );
    }

    /**
     * Per-paragraph model ranges covered by a non-collapsed selection, one range
     * per touched paragraph, clamped to that paragraph's text length.
     */
    function blockRangesForSelection(document, selection) {
      const order = selectionOrderOf(document);
      const normalized = normalizeSelection(selection, order);
      if (!normalized || isCollapsed(normalized)) return [];
      const { list } = collectBlocks(document);
      const startIndex = list.findIndex(
        (entry) => entry.block && entry.block.id === normalized.start.blockId,
      );
      const endIndex = list.findIndex(
        (entry) => entry.block && entry.block.id === normalized.end.blockId,
      );
      if (startIndex < 0 || endIndex < 0) return [];
      const ranges = [];
      for (let index = startIndex; index <= endIndex; index += 1) {
        const entry = list[index];
        if (!entry.block || entry.block.type !== "paragraph") continue;
        const text = paragraphText(entry.block);
        const from = index === startIndex ? normalized.start.offset || 0 : 0;
        const to =
          index === endIndex ? normalized.end.offset || 0 : text.length;
        const start = clampOffset(Math.min(from, to), text.length);
        const end = clampOffset(Math.max(from, to), text.length);
        if (start === end) continue;
        ranges.push({ blockId: entry.block.id, start, end });
      }
      return ranges;
    }

    /**
     * `formatText` ops for a character-formatting command.
     *
     * With a real selection this is one op per touched paragraph. With a
     * collapsed caret it is Word behaviour: the caret paragraph's whole text is
     * formatted, so the change is visible and text typed at the caret inherits
     * it. Returns an empty array when there is no usable caret.
     */
    function formatTextOpsForSelection(document, selection, props, options) {
      if (!document) return [];
      const normalized = normalizeSelection(
        selection,
        selectionOrderOf(document),
      );
      if (!normalized) return [];
      const clear = Boolean(options && options.clear === true);
      const toOp = (blockId, start, end) => ({
        op: "formatText",
        args: {
          range: {
            start: { blockId, offset: start },
            end: { blockId, offset: end },
          },
          props,
          clear,
        },
      });
      if (isCollapsed(normalized)) {
        const entry = collectBlocks(document).byId.get(
          normalized.start.blockId,
        );
        if (!entry || !entry.block || entry.block.type !== "paragraph")
          return [];
        return [toOp(entry.block.id, 0, paragraphText(entry.block).length)];
      }
      return blockRangesForSelection(document, normalized).map((range) =>
        toOp(range.blockId, range.start, range.end),
      );
    }

    /* ------------------------------------------------- selection helpers */

    /** Word boundary range around `offset`, for double-click selection. */
    function wordRangeAt(text, offset) {
      const value = typeof text === "string" ? text : "";
      const total = value.length;
      const target = clampOffset(offset, total);
      const isWord = (character) =>
        Boolean(character) && /[\p{L}\p{N}_'’-]/u.test(character);
      if (!isWord(value[target]) && !isWord(value[target - 1]))
        return { start: target, end: target };
      let start = target;
      let end = target;
      while (start > 0 && isWord(value[start - 1])) start -= 1;
      while (end < total && isWord(value[end])) end += 1;
      // A caret just after the word (double-click at its right edge) still
      // selects that word, which is the behaviour users expect.
      if (start === end && target > 0 && isWord(value[target - 1])) {
        let from = target;
        let to = target;
        while (from > 0 && isWord(value[from - 1])) from -= 1;
        while (to < total && isWord(value[to])) to += 1;
        return { start: from, end: to };
      }
      return { start, end };
    }

    /** Ordered paragraph ids with their text length, for caret arithmetic. */
    function paragraphFlow(document) {
      const { list } = collectBlocks(document || {});
      const flow = [];
      for (const entry of list) {
        if (!entry.block || entry.block.type !== "paragraph") continue;
        flow.push({
          blockId: entry.block.id,
          length: paragraphText(entry.block).length,
        });
      }
      return flow;
    }

    function clampPointToFlow(flow, point) {
      if (!point || flow.length === 0) return null;
      let index = flow.findIndex((entry) => entry.blockId === point.blockId);
      if (index < 0) index = 0;
      return {
        blockId: flow[index].blockId,
        offset: clampOffset(point.offset || 0, flow[index].length),
      };
    }

    /**
     * One arrow-key step in the document's text flow. Left/Right cross
     * paragraph boundaries (so Shift+Arrow extends across blocks); Up/Down keep
     * the column offset and move one paragraph.
     */
    function moveSelectionFocus(document, point, key) {
      const flow = paragraphFlow(document);
      if (flow.length === 0) return null;
      let index = flow.findIndex(
        (entry) => entry.blockId === (point && point.blockId),
      );
      if (index < 0) index = 0;
      let offset = clampOffset(
        (point && point.offset) || 0,
        flow[index].length,
      );
      if (key === "ArrowLeft") {
        if (offset > 0) offset -= 1;
        else if (index > 0) {
          index -= 1;
          offset = flow[index].length;
        }
      } else if (key === "ArrowRight") {
        if (offset < flow[index].length) offset += 1;
        else if (index < flow.length - 1) {
          index += 1;
          offset = 0;
        }
      } else if (key === "ArrowUp") {
        index = Math.max(0, index - 1);
        offset = clampOffset(offset, flow[index].length);
      } else if (key === "ArrowDown") {
        index = Math.min(flow.length - 1, index + 1);
        offset = clampOffset(offset, flow[index].length);
      } else if (key === "Home") {
        offset = 0;
      } else if (key === "End") {
        offset = flow[index].length;
      } else {
        return clampPointToFlow(flow, point);
      }
      return { blockId: flow[index].blockId, offset };
    }

    /**
     * New selection after Shift+Arrow/Home/End. `anchor` is the fixed end;
     * without one the current start is used (Word/Chrome semantics).
     */
    function selectionForKeyboardExtension(document, selection, anchor, key) {
      const order = selectionOrderOf(document);
      const normalized = normalizeSelection(selection, order);
      const startPoint = anchor || (normalized ? normalized.start : null);
      if (!startPoint) return null;
      const focus = normalized ? normalized.end : startPoint;
      const nextFocus = moveSelectionFocus(document, focus, key);
      if (!nextFocus) return null;
      return normalizeSelection({ start: startPoint, end: nextFocus }, order);
    }

    /**
     * The panel's public selection shape: `{ blocks, from, to }` plus the
     * collapsed flag and the plain text the selection covers.
     */
    function selectionStateOf(document, selection) {
      const order = selectionOrderOf(document || {});
      const normalized = normalizeSelection(selection, order);
      if (!normalized) return null;
      const { list } = collectBlocks(document || {});
      const startIndex = list.findIndex(
        (entry) => entry.block && entry.block.id === normalized.start.blockId,
      );
      const endIndex = list.findIndex(
        (entry) => entry.block && entry.block.id === normalized.end.blockId,
      );
      if (startIndex < 0 || endIndex < 0) return null;
      const blocks = [];
      for (
        let index = Math.min(startIndex, endIndex);
        index <= Math.max(startIndex, endIndex) && index < list.length;
        index += 1
      ) {
        const entry = list[index];
        if (entry.block && entry.block.type === "paragraph")
          blocks.push(entry.block.id);
      }
      const from = {
        blockId: normalized.start.blockId,
        offset: normalized.start.offset || 0,
      };
      const to = {
        blockId: normalized.end.blockId,
        offset: normalized.end.offset || 0,
      };
      return {
        collapsed: isCollapsed(normalized),
        blocks,
        from,
        to,
        text: selectionTextOf(document, normalized),
      };
    }

    /** Plain text covered by an already-normalised selection, "\n" joined. */
    function selectionTextOf(document, selection) {
      const normalized = selection;
      if (!normalized || isCollapsed(normalized)) return "";
      const { list } = collectBlocks(document || {});
      const startIndex = list.findIndex(
        (entry) => entry.block && entry.block.id === normalized.start.blockId,
      );
      const endIndex = list.findIndex(
        (entry) => entry.block && entry.block.id === normalized.end.blockId,
      );
      if (startIndex < 0 || endIndex < 0) return "";
      const pieces = [];
      for (
        let index = startIndex;
        index <= endIndex && index < list.length;
        index += 1
      ) {
        const entry = list[index];
        if (!entry.block || entry.block.type !== "paragraph") continue;
        const text = paragraphText(entry.block);
        const from =
          index === startIndex
            ? Math.min(normalized.start.offset, text.length)
            : 0;
        const to =
          index === endIndex
            ? Math.min(normalized.end.offset, text.length)
            : text.length;
        pieces.push(text.slice(Math.min(from, to), Math.max(from, to)));
      }
      return pieces.join("\n");
    }

    /* -------------------------------------------------- text flush plan */

    /**
     * Turn the pending per-block DOM text into the exact Host ops required to
     * make the model match, and drop every no-op.
     *
     * This is the duplication guard: an entry whose text already equals the
     * last text the Host confirmed produces no op at all, and each block
     * produces at most one op per flush. `skipped` explains why an entry was
     * not sent, which the tests assert on.
     */
    function planTextFlush(hostTexts, pendingEntries) {
      const ops = [];
      const skipped = [];
      const entries =
        pendingEntries instanceof Map
          ? Array.from(pendingEntries.entries())
          : Array.isArray(pendingEntries)
            ? pendingEntries
            : [];
      // Last write wins per block: a second entry for the same paragraph is
      // the same logical edit and must not become a second op.
      const latest = new Map();
      for (const pair of entries) {
        if (Array.isArray(pair) && typeof pair[0] === "string")
          latest.set(pair[0], pair[1]);
      }
      for (const [blockId, text] of latest) {
        if (typeof blockId !== "string" || !blockId) continue;
        if (typeof text !== "string") continue;
        const hostText =
          hostTexts instanceof Map
            ? hostTexts.get(blockId)
            : hostTexts && typeof hostTexts === "object"
              ? hostTexts[blockId]
              : undefined;
        if (hostText === undefined) {
          skipped.push({ blockId, reason: "unknown-block" });
          continue;
        }
        if (hostText === text) {
          skipped.push({ blockId, reason: "in-sync" });
          continue;
        }
        const diff = diffTextRange(hostText, text);
        if (!diff) {
          skipped.push({ blockId, reason: "no-diff" });
          continue;
        }
        ops.push({
          blockId,
          start: diff.start,
          end: diff.end,
          insert: diff.insert,
          text,
        });
      }
      return { ops, skipped };
    }

    /** Apply one planned op to a plain string (used by the differential tests). */
    function applyTextDiff(text, start, end, insert) {
      const value = typeof text === "string" ? text : "";
      const from = clampOffset(start, value.length);
      const to = Math.max(from, clampOffset(end, value.length));
      return (
        value.slice(0, from) +
        (typeof insert === "string" ? insert : "") +
        value.slice(to)
      );
    }

    /** True when two selections address the same two points. */
    function sameSelection(a, b) {
      if (!a || !b || !a.start || !a.end || !b.start || !b.end) return false;
      return (
        a.start.blockId === b.start.blockId &&
        (a.start.offset || 0) === (b.start.offset || 0) &&
        a.end.blockId === b.end.blockId &&
        (a.end.offset || 0) === (b.end.offset || 0)
      );
    }

    /**
     * Pointer-selection diagnostics. They make the drag/keyboard selection path
     * observable from the harness (and from the live v2 probe) without exposing
     * component internals: how many pointer events arrived, what point each one
     * resolved to, and why a selection application was refused.
     */
    const selectionRuntime = {
      pointerDown: 0,
      pointerMove: 0,
      applied: 0,
      failed: 0,
      lastDown: null,
      lastMove: null,
      lastPoint: null,
      lastApply: null,
      lastApplied: null,
      lastUpdate: null,
      patchCount: 0,
      lastError: null,
      lastKeyboard: null,
    };

    function selectionDiagnostics() {
      return {
        pointerDown: selectionRuntime.pointerDown,
        pointerMove: selectionRuntime.pointerMove,
        applied: selectionRuntime.applied,
        failed: selectionRuntime.failed,
        lastDown: selectionRuntime.lastDown,
        lastMove: selectionRuntime.lastMove,
        lastPoint: selectionRuntime.lastPoint,
        lastApply: selectionRuntime.lastApply,
        lastApplied: selectionRuntime.lastApplied,
        lastUpdate: selectionRuntime.lastUpdate,
        patchCount: selectionRuntime.patchCount,
        lastError: selectionRuntime.lastError,
        lastKeyboard: selectionRuntime.lastKeyboard,
        structuralDepth: selectionRuntime.structuralDepth || 0,
        structuralInputs: selectionRuntime.structuralInputs || 0,
        deferredChars: selectionRuntime.deferredChars || 0,
        lastPlacement: selectionRuntime.lastPlacement,
        lastSplit: selectionRuntime.lastSplit,
      };
    }

    /* ------------------------------------------------------- Host client */

    /**
     * One HTTP call against the Documents API. JSON failures become a
     * `HostError` that carries the Host code and message so the panel can show
     * both instead of a silent no-op.
     */
    function HostError(code, message, detail) {
      const error = new Error(message || code || "Document request failed");
      error.name = "HostError";
      error.code = code || "REQUEST_FAILED";
      error.detail = detail || null;
      return error;
    }

    function statusCodeOf(error) {
      if (error && typeof error.status === "number") return error.status;
      return 0;
    }

    function isMissingRoute(error) {
      const status = statusCodeOf(error);
      if (status === 404 || status === 405 || status === 501) return true;
      const code = String((error && error.code) || "");
      if (/^UNKNOWN_?(ROUTE|ENDPOINT|ACTION)$/i.test(code)) return true;
      if (/^NOT_?IMPLEMENTED$/i.test(code)) return true;
      const message = String((error && error.message) || "");
      return /unknown documents api route|not implemented|no such route/i.test(
        message,
      );
    }

    async function request(path, options) {
      const init = Object.assign({}, options || {});
      init.headers = Object.assign({}, HEADER, init.headers || {});
      let response;
      try {
        response = await fetch(`${API}${path}`, init);
      } catch (cause) {
        const wrapped = HostError(
          "NETWORK_ERROR",
          `Cannot reach the Documents API: ${cause && cause.message ? cause.message : cause}`,
        );
        wrapped.status = 0;
        throw wrapped;
      }
      const contentType = response.headers.get("content-type") || "";
      if (!contentType.includes("application/json")) {
        if (!response.ok) {
          const error = HostError(
            `HTTP_${response.status}`,
            `Request failed (${response.status} ${response.statusText || ""})`.trim(),
          );
          error.status = response.status;
          throw error;
        }
        return response;
      }
      let payload;
      try {
        payload = await response.json();
      } catch (_cause) {
        throw HostError("BAD_RESPONSE", "The Host returned invalid JSON.");
      }
      if (!payload || payload.ok === false) {
        const detail = (payload && payload.error) || {};
        const error = HostError(
          detail.code || `HTTP_${response.status}`,
          detail.message || "The document operation failed.",
          payload,
        );
        error.status = response.status;
        error.payload = payload;
        throw error;
      }
      return payload;
    }

    const apiClient = {
      list() {
        return request("/list");
      },
      bootstrap(documentId) {
        return request(`/bootstrap?id=${encodeURIComponent(documentId)}`);
      },
      document(documentId) {
        return request(`/document?id=${encodeURIComponent(documentId)}`);
      },
      op(documentId, op, args) {
        return request("/op", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ id: documentId, op, args: args || {} }),
        });
      },
      ops(documentId, ops) {
        return request("/ops", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ id: documentId, ops }),
        });
      },
      save(documentId, path) {
        return request("/save", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify(
            path ? { id: documentId, path } : { id: documentId },
          ),
        });
      },
      saveAs(documentId, path) {
        return request("/saveAs", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ id: documentId, path }),
        });
      },
      create(title) {
        return request("/create", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ title }),
        });
      },
      open(path) {
        return request("/open", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ path }),
        });
      },
      openUpload(file) {
        if (typeof FormData !== "function")
          return Promise.reject(
            HostError(
              "UPLOAD_UNSUPPORTED",
              "This browser cannot build a multipart upload.",
            ),
          );
        const body = new FormData();
        body.append("file", file, file.name || "document");
        // No content-type header: the browser adds the multipart boundary.
        return request("/open", { method: "POST", body });
      },
      openPath(path) {
        return request("/open-path", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ path }),
        });
      },
      browse(path) {
        const suffix = path ? `?path=${encodeURIComponent(String(path))}` : "";
        return request(`/browse${suffix}`);
      },
      workspace(path, create) {
        return request("/workspace", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ path, create: create === true }),
        });
      },
      remove(documentId) {
        return request("/delete", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ id: documentId }),
        });
      },
      search(documentId, query, options) {
        const settings = options || {};
        const params = [
          `id=${encodeURIComponent(documentId)}`,
          `q=${encodeURIComponent(query)}`,
          `regex=${settings.regex ? "1" : "0"}`,
          `caseSensitive=${settings.caseSensitive ? "1" : "0"}`,
        ];
        return request(`/search?${params.join("&")}`);
      },
      exportUrl(documentId, format) {
        return `${API}/export?id=${encodeURIComponent(documentId)}&format=${encodeURIComponent(format)}`;
      },
    };

    /** Download an export through the browser (no network assets, no storage). */
    function downloadExport(documentId, format, title) {
      if (typeof window === "undefined") return Promise.resolve();
      return fetch(apiClient.exportUrl(documentId, format), {
        headers: HEADER,
      }).then(async (response) => {
        if (!response.ok) {
          const text = await response.text().catch(() => "");
          throw HostError(
            `HTTP_${response.status}`,
            text || `Export failed (${response.status})`,
          );
        }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const anchor = window.document.createElement("a");
        anchor.href = url;
        anchor.download = `${(title || "document")
          .replace(/[^\w .-]+/g, "_")
          .slice(0, 80)}.${format}`;
        window.document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        URL.revokeObjectURL(url);
      });
    }

    /* ------------------------------------------- optional client services */

    /**
     * The Cordis Context bound by `apply`. The panel reads the optional Client
     * services (`uiWorkspace`, `layout`) through it instead of injecting them:
     * an older Host that lacks either service must still render the editor, so
     * every use is guarded and degrades to the advanced path input.
     */
    let clientContext = null;

    function bindClientContext(ctx) {
      clientContext = ctx || null;
    }

    /** One optional Client service, or undefined when the page has none. */
    function clientService(name) {
      const ctx = clientContext;
      if (!ctx) return undefined;
      let value;
      try {
        if (typeof ctx.get === "function") value = ctx.get(name);
      } catch (_error) {
        value = undefined;
      }
      if (value !== undefined) return value;
      return ctx[name];
    }

    function workspaceService() {
      const service = clientService("uiWorkspace");
      return service && typeof service === "object" ? service : undefined;
    }

    function layoutService() {
      const service = clientService("layout");
      return service && typeof service === "object" ? service : undefined;
    }

    /** Whether this Host can open the native OS folder dialog. */
    function canPickDirectory() {
      const service = workspaceService();
      return Boolean(service && typeof service.pickDirectory === "function");
    }

    /** Whether this Host can prepare a chat session in a Harness workspace. */
    function canStartSession() {
      const service = workspaceService();
      return Boolean(service && typeof service.startSession === "function");
    }

    /* ------------------------------------------------ ribbon label prefs */

    function readStoredLabels() {
      try {
        if (typeof localStorage === "undefined" || !localStorage) return false;
        return localStorage.getItem(LABELS_KEY) === "on";
      } catch (_error) {
        return false;
      }
    }

    function storeLabels(value) {
      try {
        if (typeof localStorage === "undefined" || !localStorage) return;
        if (value) localStorage.setItem(LABELS_KEY, "on");
        else localStorage.removeItem(LABELS_KEY);
      } catch (_error) {
        /* Storage may be unavailable (private mode): the toggle still works. */
      }
    }

    /* ------------------------------------------------------ path helpers */

    function pathSeparatorOf(path) {
      return String(path || "").indexOf("\\") >= 0 ? "\\" : "/";
    }

    function isRootPath(path) {
      const value = String(path || "");
      return value === "/" || /^[A-Za-z]:[\\/]$/.test(value);
    }

    function trimTrailingSeparator(path) {
      const value = String(path || "");
      if (!value || isRootPath(value)) return value;
      const trimmed = value.replace(/[\\/]+$/, "");
      return trimmed || pathSeparatorOf(value);
    }

    function baseNameOf(path) {
      const value = trimTrailingSeparator(path);
      if (!value) return "";
      const index = Math.max(value.lastIndexOf("/"), value.lastIndexOf("\\"));
      if (index === -1) return value;
      return value.slice(index + 1) || value;
    }

    function joinPath(directory, name) {
      const base = trimTrailingSeparator(directory);
      const child = String(name || "").replace(/^[\\/]+/, "");
      if (!base) return child;
      if (isRootPath(base)) return `${base}${child}`;
      return `${base}${pathSeparatorOf(base)}${child}`;
    }

    /**
     * Breadcrumb segments for one absolute directory, each carrying the path
     * that clicking it should browse.
     */
    function breadcrumbSegments(path) {
      const value = trimTrailingSeparator(path);
      if (!value) return [];
      const windows = /^[A-Za-z]:/.test(value);
      const sep = windows ? "\\" : "/";
      const parts = value.split(/[\\/]+/).filter(Boolean);
      const segments = [];
      let current = "";
      parts.forEach((part, index) => {
        if (index === 0 && windows) {
          current = `${part}${sep}`;
          segments.push({ label: part, path: current });
          return;
        }
        current = isRootPath(current)
          ? `${current}${part}`
          : current
            ? `${current}${sep}${part}`
            : `${sep}${part}`;
        segments.push({ label: part, path: current });
      });
      if (!windows && segments.length === 0) return [{ label: "/", path: "/" }];
      return segments;
    }

    function formatBytes(size) {
      const value = Number(size);
      if (!Number.isFinite(value) || value < 0) return "";
      if (value < 1024) return `${value} B`;
      if (value < 1024 * 1024) return `${Math.round(value / 1024)} KB`;
      return `${(value / (1024 * 1024)).toFixed(1)} MB`;
    }

    /* ----------------------------------- workspace switching + sessions */

    /**
     * Page-session registry of prepared chat sessions, keyed by workspace path.
     * Opening the same folder twice reuses the prepared session instead of
     * starting a second one.
     */
    const PREPARED_SESSIONS = new Map();

    function clearPreparedSessions() {
      PREPARED_SESSIONS.clear();
    }

    /**
     * Switch the panel's document workspace to `path` through
     * `POST /workspace`, then prepare one chat session there.
     *
     * Returns a chip description the panel header renders:
     * - `chat-ready`: the Host returned `dshWorkspaceId` and `startSession`
     *   exists (or the session was already prepared for this path);
     * - `workspace-only`: the workspace is ready, but this Host has no
     *   workspace service, so no chat session could be started;
     * - `no-harness-workspace`: the Host returned no `dshWorkspaceId`.
     *
     * A failed `POST /workspace` throws (the caller reports it); a failed
     * `startSession` is reported in `warning` and never blocks the editor.
     */
    async function prepareWorkspaceSession(options) {
      const input = options || {};
      const services = input.services || {};
      const uiWorkspace =
        input.uiWorkspace ||
        services.uiWorkspace ||
        (input.useBoundServices === false ? undefined : workspaceService());
      const callWorkspace =
        input.callWorkspace ||
        function requestWorkspace(path, create) {
          return apiClient.workspace(path, create);
        };
      const prepared = input.prepared || PREPARED_SESSIONS;
      const payload = await callWorkspace(input.path, input.create === true);
      const path = trimTrailingSeparator(
        (payload && payload.path) || input.path || "",
      );
      const chip = {
        path,
        folderName: baseNameOf(path) || path,
        dshWorkspaceId:
          payload && typeof payload.dshWorkspaceId === "string"
            ? payload.dshWorkspaceId
            : "",
        created: Boolean(payload && payload.created === true),
        status: "no-harness-workspace",
        sessionStarted: false,
        reused: false,
        warning: null,
      };
      if (!chip.dshWorkspaceId) {
        chip.warning =
          "The Host did not return a Harness workspace id, so a chat session could not be prepared.";
        return chip;
      }
      const previous = prepared.get(path);
      if (previous && previous.dshWorkspaceId === chip.dshWorkspaceId) {
        chip.status = "chat-ready";
        chip.reused = true;
        return chip;
      }
      if (!uiWorkspace || typeof uiWorkspace.startSession !== "function") {
        const canConnect =
          uiWorkspace && typeof uiWorkspace.connectWorkspace === "function";
        if (!canConnect) {
          chip.status = "workspace-only";
          chip.warning =
            "The workspace is ready, but a chat session could not be started automatically: this Host has no workspace service.";
          return chip;
        }
      }
      try {
        if (typeof uiWorkspace.connectWorkspace === "function") {
          // connectWorkspace resolves (and reuses) the blank Session for this
          // Workspace WITHOUT navigating away from the Documents panel, so the
          // document the user just opened stays on screen. "Switch to chat"
          // opens it on demand.
          const sessionId = await uiWorkspace.connectWorkspace(
            chip.dshWorkspaceId,
          );
          if (sessionId) chip.sessionId = String(sessionId);
        } else {
          uiWorkspace.startSession(chip.dshWorkspaceId);
        }
      } catch (cause) {
        chip.status = "workspace-only";
        chip.warning = `The workspace is ready, but preparing the chat session failed: ${
          cause && cause.message ? cause.message : cause
        }`;
        return chip;
      }
      prepared.set(path, {
        dshWorkspaceId: chip.dshWorkspaceId,
        sessionId: chip.sessionId,
        startedAt: Date.now(),
      });
      chip.status = "chat-ready";
      chip.sessionStarted = true;
      return chip;
    }

    /** Header text for one prepared-workspace chip. */
    function sessionChipText(chip) {
      if (!chip) return "";
      if (chip.status === "chat-ready")
        return `Chat ready · ${chip.folderName || chip.path}`;
      if (chip.status === "workspace-only")
        return `Workspace ready · ${chip.path}`;
      return `Workspace · ${chip.path}`;
    }

    /**
     * Show the Conversation again. The layout service accepts the registered
     * main key (`conversation`); null is the same panel on the current Host, so
     * both are attempted without forcing navigation anywhere else.
     */
    function switchToConversation(deps) {
      const input = deps || {};
      const services = input.services || {};
      const layout =
        input.layout ||
        services.layout ||
        (input.useBoundServices === false ? undefined : layoutService());
      if (!layout || typeof layout.selectPanel !== "function") {
        return {
          ok: false,
          message:
            "This Host cannot switch the main panel from the editor, so open the chat from the sidebar.",
        };
      }
      try {
        layout.selectPanel("conversation");
        return { ok: true, message: "" };
      } catch (cause) {
        return {
          ok: false,
          message: `The conversation panel could not be opened (${
            cause && cause.message ? cause.message : cause
          }); use the sidebar instead.`,
        };
      }
    }

    /**
     * File → Choose folder…: native OS directory picker, then switch the
     * document workspace and prepare a chat session there.
     *
     * Throws a `HostError` with `fallback: "path"` when this Host has no
     * picker (or the picker fails); the panel then explains and reveals the
     * advanced path input instead of blocking the editor.
     */
    async function runChooseFolderFlow(options) {
      const input = options || {};
      const services = input.services || {};
      const uiWorkspace =
        input.uiWorkspace ||
        services.uiWorkspace ||
        (input.useBoundServices === false ? undefined : workspaceService());
      if (!uiWorkspace || typeof uiWorkspace.pickDirectory !== "function") {
        const error = HostError(
          "PICKER_UNAVAILABLE",
          "This Host does not expose the native folder picker. Use the “Advanced: path” input in the File tab instead.",
        );
        error.fallback = "path";
        throw error;
      }
      let directory;
      try {
        directory = await uiWorkspace.pickDirectory();
      } catch (cause) {
        const error = HostError(
          "PICKER_FAILED",
          `The folder picker failed (${
            cause && cause.message ? cause.message : cause
          }). Use the “Advanced: path” input in the File tab instead.`,
        );
        error.fallback = "path";
        throw error;
      }
      if (!directory)
        return { ok: false, cancelled: true, chip: null, path: "" };
      const chip = await prepareWorkspaceSession({
        path: directory,
        create: false,
        uiWorkspace,
        callWorkspace: input.callWorkspace,
        prepared: input.prepared,
        useBoundServices: false,
      });
      return { ok: true, cancelled: false, path: chip.path, chip };
    }

    /**
     * Open a document the folder browser selected (`POST /open-path`), then
     * switch the document workspace to its folder and prepare a chat session.
     */
    async function runOpenDocumentPathFlow(path, options) {
      const input = options || {};
      const services = input.services || {};
      const uiWorkspace =
        input.uiWorkspace ||
        services.uiWorkspace ||
        (input.useBoundServices === false ? undefined : workspaceService());
      const openPath =
        input.callOpenPath ||
        function requestOpenPath(target) {
          return apiClient.openPath(target);
        };
      const opened = await openPath(path);
      let chip = null;
      let workspaceError = null;
      if (opened && opened.directory) {
        try {
          chip = await prepareWorkspaceSession({
            path: opened.directory,
            create: false,
            uiWorkspace,
            callWorkspace: input.callWorkspace,
            prepared: input.prepared,
            useBoundServices: false,
          });
        } catch (cause) {
          // The document opened already: report the workspace failure without
          // discarding it.
          workspaceError = cause;
        }
      }
      return {
        ok: true,
        opened,
        chip,
        workspaceError,
        documentId: (opened && opened.documentId) || "",
        directory: (opened && opened.directory) || "",
      };
    }

    /**
     * Upload one document through `POST /open` (multipart), switch to the
     * folder the Host reports and prepare a chat session there.
     */
    async function runOpenUploadFlow(file, options) {
      const input = options || {};
      const services = input.services || {};
      const uiWorkspace =
        input.uiWorkspace ||
        services.uiWorkspace ||
        (input.useBoundServices === false ? undefined : workspaceService());
      const openUpload =
        input.callOpenUpload ||
        function requestOpenUpload(picked) {
          return apiClient.openUpload(picked);
        };
      const opened = await openUpload(file);
      let chip = null;
      let workspaceError = null;
      if (opened && opened.directory) {
        try {
          chip = await prepareWorkspaceSession({
            path: opened.directory,
            create: false,
            uiWorkspace,
            callWorkspace: input.callWorkspace,
            prepared: input.prepared,
            useBoundServices: false,
          });
        } catch (cause) {
          workspaceError = cause;
        }
      }
      return {
        ok: true,
        opened,
        chip,
        workspaceError,
        documentId: (opened && opened.documentId) || "",
        directory: (opened && opened.directory) || "",
        title: (opened && opened.title) || "",
      };
    }

    /* ------------------------------------------- folder browser helpers */

    function normalizeBrowseEntry(entry) {
      const data = entry || {};
      const kind = data.kind === "directory" ? "directory" : "document";
      return {
        name: String(data.name || ""),
        path: String(data.path || ""),
        kind,
        importable: kind === "document" ? data.importable === true : false,
        size: typeof data.size === "number" ? data.size : null,
        modified: data.modified || null,
      };
    }

    /** Turn one `GET /browse?path=` payload into folder-browser dialog state. */
    function browseDialogState(payload, options) {
      const data = payload || {};
      const entries = (Array.isArray(data.entries) ? data.entries : [])
        .map(normalizeBrowseEntry)
        .filter((entry) => entry.name && entry.path);
      entries.sort(function order(left, right) {
        if (left.kind !== right.kind) return left.kind === "directory" ? -1 : 1;
        return left.name.localeCompare(right.name);
      });
      return Object.assign(
        {
          kind: "browse",
          path: String(data.path || ""),
          parent: data.parent || null,
          roots: (Array.isArray(data.roots) ? data.roots : [])
            .map((root) => ({
              label: String((root && root.label) || ""),
              path: String((root && root.path) || ""),
            }))
            .filter((root) => root.label && root.path),
          entries,
          truncated: data.truncated === true,
          loading: false,
          error: null,
          newFolder: null,
        },
        options || {},
      );
    }

    function describeBrowseFailure(cause) {
      if (isMissingRoute(cause)) {
        return "This Host does not implement /browse, so the folder browser is unavailable. Use the “Advanced: path” input in the File tab instead.";
      }
      return (
        (cause && cause.message) ||
        "The folder could not be listed by the Host."
      );
    }

    /* -------------------------------------------------------- panel state */

    function initialState() {
      return {
        documents: [],
        workspace: "",
        bootstrap: null,
        selectedId: null,
        tab: "home",
        ribbonCollapsed: false,
        zoom: 1,
        fitWidth: false,
        showRuler: true,
        showMarks: false,
        showPageSeparators: true,
        focusMode: false,
        showRail: true,
        readOnlyOverride: null,
        workspaceReadOnly: false,
        fallback: null,
        saveState: "saved",
        saving: false,
        loading: false,
        error: null,
        notice: "",
        noticeDismissible: false,
        boldReserved: false,
        offline: false,
        dialog: null,
        selection: null,
        canUndo: true,
        canRedo: true,
        dialogReturnFocus: null,
        showLabels: readStoredLabels(),
        ribbonHeight: readStoredRibbonHeight(),
        advancedOpen: false,
        advancedPath: "",
        browseRoots: [],
        sessionChip: null,
      };
    }

    function normalizeBootstrap(payload) {
      const document = payload.document || { sections: [] };
      const stats = Object.assign(computeStats(document), payload.stats || {});
      return {
        documentId: payload.documentId || "document",
        title:
          payload.title ||
          (document.metadata && document.metadata.title) ||
          "Untitled document",
        dirty: payload.dirty === true,
        revision: payload.revision || 1,
        sourceFormat: payload.sourceFormat || "docx",
        workspace: payload.workspace || "",
        readOnly: payload.readOnly === true,
        formats: payload.formats || { import: [], export: [] },
        document,
        blockPages: payload.blockPages || null,
        pageCount: payload.pageCount || 1,
        stats,
        outline: payload.outline || computeOutline(document),
        capabilities: payload.capabilities || {},
        selection: {
          start: {
            blockId:
              (documentBodyBlocks(document)[0] &&
                documentBodyBlocks(document)[0].id) ||
              "",
            offset: 0,
          },
          end: {
            blockId:
              (documentBodyBlocks(document)[0] &&
                documentBodyBlocks(document)[0].id) ||
              "",
            offset: 0,
          },
        },
      };
    }

    /** Page map used when the Host does not send one (single-page estimate). */
    function defaultPageMap(document) {
      const map = {};
      let page = 1;
      for (const block of documentBodyBlocks(document)) {
        if (!block || !block.id) continue;
        map[block.id] = page;
        if (block.type !== "paragraph") continue;
        for (const inline of block.inlines || []) {
          if (
            inline &&
            inline.type === "break" &&
            (inline.kind === "page" || inline.kind === "column")
          )
            page += 1;
        }
      }
      return map;
    }

    /**
     * Seed the whole panel state without any Host call. Used by the Node tests
     * (and by nothing else): `renderForTest` renders from this state.
     */
    function seedState(model, options) {
      const opts = options || {};
      const document = model || { sections: [] };
      const bodyBlocks = documentBodyBlocks(document);
      const firstId = (bodyBlocks[0] && bodyBlocks[0].id) || "p1";
      const blockPages = opts.blockPages || defaultPageMap(document);
      const pageCount =
        opts.pageCount ||
        Math.max(
          1,
          Object.keys(blockPages).reduce(
            (max, key) => Math.max(max, blockPages[key] || 1),
            1,
          ),
        );
      const stats = computeStats(document);
      const bootstrap = {
        documentId: opts.documentId || "doc-test",
        title:
          opts.title ||
          (document.metadata && document.metadata.title) ||
          "Untitled document",
        dirty: opts.dirty === true,
        revision: opts.revision || 1,
        sourceFormat: opts.sourceFormat || "docx",
        workspace: opts.workspace || "C:/Users/editor/.dsh-documents",
        readOnly: opts.readOnly === true,
        formats: opts.formats || {
          import: ["docx", "doc"],
          export: ["docx", "pdf", "html", "md"],
        },
        document,
        blockPages,
        pageCount,
        stats: Object.assign({}, stats, opts.stats || {}),
        outline: opts.outline || computeOutline(document),
        capabilities: opts.capabilities || {
          ai: false,
          formats: true,
          convert: true,
          renderer: true,
        },
        selection: {
          start: { blockId: firstId, offset: 0 },
          end: { blockId: firstId, offset: 0 },
        },
      };
      const state = initialState();
      state.documents = opts.documents || [
        {
          id: bootstrap.documentId,
          title: bootstrap.title,
          format: bootstrap.sourceFormat,
          dirty: bootstrap.dirty,
          words: stats.words,
        },
      ];
      state.workspace = bootstrap.workspace;
      state.bootstrap = opts.noDocument === true ? null : bootstrap;
      state.selectedId = opts.noDocument === true ? null : bootstrap.documentId;
      state.selection =
        opts.noDocument === true ? null : opts.selection || bootstrap.selection;
      if (opts.selection) bootstrap.selection = opts.selection;
      state.tab = opts.tab || "home";
      state.ribbonCollapsed = opts.ribbonCollapsed === true;
      if (typeof opts.ribbonHeight === "number")
        state.ribbonHeight = opts.ribbonHeight;
      state.zoom = typeof opts.zoom === "number" ? opts.zoom : 1;
      state.readOnlyOverride =
        typeof opts.readOnly === "boolean" ? opts.readOnly : null;
      state.saveState = opts.saveState || "saved";
      state.error = opts.error || null;
      state.notice = opts.notice || "";
      state.noticeDismissible = opts.noticeDismissible === true;
      state.boldReserved = opts.boldReserved === true;
      state.offline = opts.offline === true;
      state.dialog = opts.dialog || null;
      state.fallback = opts.fallback || null;
      state.showMarks = opts.showMarks === true;
      state.showRuler = opts.showRuler !== false;
      state.focusMode = opts.focusMode === true;
      state.showPageSeparators = opts.showPageSeparators !== false;
      state.showRail = opts.showRail !== false;
      state.showLabels = opts.showLabels === true;
      state.advancedOpen = opts.advancedOpen === true;
      state.advancedPath = opts.advancedPath || "";
      state.browseRoots = opts.browseRoots || [];
      state.sessionChip = opts.sessionChip || null;
      return state;
    }

    function selectionOrderOf(document) {
      const order = new Map();
      const { list } = collectBlocks(document);
      list.forEach((entry, index) => {
        if (entry.block && entry.block.id) order.set(entry.block.id, index);
      });
      return order;
    }

    /** Ribbon/status-bar projection of the current selection. */
    function readFormatting(document, selection, order) {
      const fallback = {
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        vertAlign: "baseline",
        color: "#1b1b1b",
        highlight: "",
        fontFamily: "",
        fontSize: 11,
        alignment: "left",
        styleId: "",
        styleName: "Normal",
        hasRange: false,
        inList: false,
        indentLeft: 0,
        lineSpacing: 0,
        spaceBefore: 0,
        spaceAfter: 0,
        link: "",
        blockId: "",
      };
      if (!document) return fallback;
      const normalized = normalizeSelection(selection, order);
      const anchor = normalized ? normalized.start : null;
      const blockId = anchor ? anchor.blockId : "";
      const { byId } = collectBlocks(document);
      const entry = blockId ? byId.get(blockId) : null;
      const paragraph =
        entry && entry.block && entry.block.type === "paragraph"
          ? entry.block
          : null;
      if (!paragraph) return Object.assign({}, fallback, { blockId });
      const offset = anchor ? anchor.offset || 0 : 0;
      const span = spanAtOffset(paragraph, offset);
      const runProps = resolveRunProps(
        document,
        paragraph,
        span ? span.inline : null,
      );
      const paragraphProps = resolveParagraphProps(document, paragraph);
      const style = findStyle(document, paragraph.style);
      const defaults =
        (document.styles && document.styles.defaults
          ? document.styles.defaults.run
          : null) || {};
      return {
        bold: Boolean(runProps.bold),
        italic: Boolean(runProps.italic),
        underline: Boolean(runProps.underline && runProps.underline !== "none"),
        strike: Boolean(runProps.strike),
        vertAlign: runProps.vertAlign || "baseline",
        color: runProps.color || "#1b1b1b",
        highlight: runProps.highlight || "",
        fontFamily: runProps.fontFamily || defaults.fontFamily || "",
        fontSize:
          typeof runProps.fontSize === "number"
            ? runProps.fontSize
            : typeof defaults.fontSize === "number"
              ? defaults.fontSize
              : 11,
        alignment: paragraphProps.alignment || "left",
        styleId: style ? styleKey(document, style) : paragraph.style || "",
        styleName: (style && style.name) || paragraph.style || "Normal",
        hasRange: Boolean(normalized && !isCollapsed(normalized)),
        inList: Boolean(
          paragraphProps.numbering && paragraphProps.numbering.numId,
        ),
        indentLeft:
          paragraphProps.indent &&
          typeof paragraphProps.indent.left === "number"
            ? paragraphProps.indent.left
            : 0,
        lineSpacing:
          paragraphProps.lineSpacing && paragraphProps.lineSpacing.value
            ? paragraphProps.lineSpacing.value / 240
            : 0,
        spaceBefore:
          typeof paragraphProps.spaceBefore === "number"
            ? paragraphProps.spaceBefore
            : 0,
        spaceAfter:
          typeof paragraphProps.spaceAfter === "number"
            ? paragraphProps.spaceAfter
            : 0,
        link: (runProps.hyperlink && runProps.hyperlink.href) || "",
        blockId,
      };
    }

    /** Build the immutable props bundle the presentational panel renders from. */
    function buildViewProps(state, api) {
      const bootstrap = state.bootstrap;
      const document = bootstrap ? bootstrap.document : null;
      const order = selectionOrderOf(document);
      const formatting = readFormatting(document, state.selection, order);
      const sections = (document && document.sections) || [];
      const media = new Map();
      for (const asset of (document && document.media) || []) {
        if (asset && asset.id) media.set(asset.id, asset);
      }
      const listLabels = document ? computeListLabels(document) : new Map();
      const readOnly =
        state.readOnlyOverride === null
          ? Boolean(bootstrap && bootstrap.readOnly) ||
            state.workspaceReadOnly === true
          : state.readOnlyOverride === true;
      const settings = (document && document.settings) || {};
      const defaultsRun =
        (document &&
          document.styles &&
          document.styles.defaults &&
          document.styles.defaults.run) ||
        {};
      const language =
        settings.defaultLanguage || defaultsRun.language || "en-US";
      const hostStats = bootstrap ? bootstrap.stats : null;
      // Live counts come from the in-memory model so the status bar follows
      // typing even before the debounced Host op lands (and while degraded).
      const stats = document ? computeStats(document) : hostStats;
      const paragraphsByPage = bootstrap
        ? groupBlocksIntoPages(
            documentBodyBlocks(document),
            bootstrap.blockPages,
          )
        : [];
      const entry = formatting.blockId
        ? collectBlocks(document).byId.get(formatting.blockId)
        : null;
      return {
        state,
        api,
        bootstrap,
        document,
        documents: state.documents,
        workspace: state.workspace,
        title: bootstrap ? bootstrap.title : "",
        selectedId: state.selectedId,
        formats: bootstrap ? bootstrap.formats : { import: [], export: [] },
        language,
        dirty: bootstrap ? bootstrap.dirty : false,
        revision: bootstrap ? bootstrap.revision : 0,
        sourceFormat: bootstrap ? bootstrap.sourceFormat : "",
        readOnly,
        saving: state.saving,
        saveState: state.saveState,
        error: state.error,
        notice: state.notice,
        fallback: state.fallback,
        offline: state.offline,
        tab: state.tab,
        ribbonCollapsed: state.ribbonCollapsed,
        // A stored (or dragged) height wins; otherwise the stylesheet's
        // clamp(72px, 10vh, 240px) default applies and ribbonHeight stays null.
        ribbonHeight:
          typeof state.ribbonHeight === "number" ? state.ribbonHeight : null,
        ribbonHeightValue:
          typeof state.ribbonHeight === "number"
            ? clampRibbonHeight(state.ribbonHeight, viewportHeight())
            : defaultRibbonHeight(viewportHeight()),
        ribbonMaxHeight: ribbonMaxHeight(viewportHeight()),
        zoom: state.zoom,
        fitWidth: state.fitWidth,
        showRuler: state.showRuler,
        showMarks: state.showMarks,
        showPageSeparators: state.showPageSeparators,
        focusMode: state.focusMode,
        showRail: state.showRail && !state.focusMode,
        showLabels: state.showLabels === true,
        advancedOpen: state.advancedOpen === true,
        advancedPath: state.advancedPath || "",
        browseRoots: state.browseRoots || [],
        sessionChip: state.sessionChip || null,
        canPickDirectory: canPickDirectory(),
        dialog: state.dialog,
        selection: state.selection,
        boldReserved: state.boldReserved === true,
        shortcutOverrides:
          state.boldReserved === true ? { bold: "Ctrl+Shift+B" } : {},
        noticeDismissible: state.noticeDismissible === true,
        surfaceEpoch: state.surfaceEpoch || 0,
        formatting,
        canUndo: state.canUndo,
        canRedo: state.canRedo,
        stats,
        pageCount: bootstrap ? bootstrap.pageCount : 0,
        currentPage: pageNumberForBlock(
          bootstrap ? bootstrap.blockPages : null,
          formatting.blockId,
        ),
        outline: bootstrap ? bootstrap.outline : [],
        sections,
        activeSection:
          entry && typeof entry.sectionIndex === "number"
            ? entry.sectionIndex
            : 0,
        media,
        listLabels,
        paragraphsByPage,
        fontFamilies: fontFamiliesFor(document || {}),
        styleGallery: styleGalleryFor(document || {}),
        fontSizes: FONT_SIZES,
        zoomSteps: ZOOM_STEPS,
        lineSpacingOptions: LINE_SPACING,
        spacingPresets: SPACING_PRESETS,
        marginPresets: MARGIN_PRESETS,
        pageSizes: PAGE_SIZES,
        themeColors: THEME_COLORS,
        reasons: REASONS,
        tabs: TABS,
        blocksById: collectBlocks(document || {}).byId,
      };
    }

    /* ---------------------------------------------------- DOM selection */

    function blockElementFor(node) {
      let element =
        node && node.nodeType === 3 ? node.parentElement : node || null;
      if (!element || typeof element.closest !== "function") return null;
      return element.closest("[data-block-id]");
    }

    /** UTF-16 offset of a DOM point inside a `[data-block-id]` element. */
    function offsetInBlockElement(blockElement, container, offset) {
      if (!blockElement || typeof blockElement.ownerDocument === "undefined")
        return 0;
      const ownerDocument = blockElement.ownerDocument;
      const range = ownerDocument.createRange();
      range.selectNodeContents(blockElement);
      try {
        range.setEnd(container, offset);
      } catch (_error) {
        return 0;
      }
      let length = range.toString().length;
      const atoms = blockElement.querySelectorAll("[data-dx-atom]");
      for (const atom of atoms) {
        try {
          if (range.intersectsNode(atom)) length += 1;
        } catch (_error) {
          /* intersect failures are ignored: the atom is not counted */
        }
      }
      return length;
    }

    function domSelectionToModel(rootElement) {
      if (typeof window === "undefined" || !rootElement) return null;
      const selection = window.getSelection();
      if (!selection || selection.rangeCount === 0) return null;
      const range = selection.getRangeAt(0);
      if (
        !rootElement.contains(range.startContainer) ||
        !rootElement.contains(range.endContainer)
      )
        return null;
      const startElement = blockElementFor(range.startContainer);
      const endElement = blockElementFor(range.endContainer);
      if (!startElement || !endElement) return null;
      return {
        start: {
          blockId: startElement.getAttribute("data-block-id") || "",
          offset: offsetInBlockElement(
            startElement,
            range.startContainer,
            range.startOffset,
          ),
        },
        end: {
          blockId: endElement.getAttribute("data-block-id") || "",
          offset: offsetInBlockElement(
            endElement,
            range.endContainer,
            range.endOffset,
          ),
        },
      };
    }

    /** Flat text of a rendered paragraph, including atom characters. */
    function domTextOfBlock(blockElement) {
      let text = "";
      const walk = (node) => {
        if (node.nodeType === 3) {
          text += node.data;
          return;
        }
        if (node.nodeType !== 1) return;
        if (node.hasAttribute("data-dx-atom")) {
          text += node.getAttribute("data-dx-atom") || "";
          return;
        }
        for (const child of node.childNodes) walk(child);
      };
      for (const child of blockElement.childNodes) walk(child);
      return text;
    }

    function findBlockElement(rootElement, blockId) {
      if (!rootElement || !blockId) return null;
      const candidates = rootElement.querySelectorAll("[data-block-id]");
      for (const candidate of candidates) {
        if (candidate.getAttribute("data-block-id") === blockId)
          return candidate;
      }
      return null;
    }

    function textPointAtOffset(blockElement, offset) {
      const total = (() => {
        let length = 0;
        for (const child of blockElement.childNodes) {
          if (child.nodeType === 3) length += child.data.length;
          else if (child.nodeType === 1) {
            if (child.hasAttribute("data-dx-atom")) length += 1;
            else length += (child.textContent || "").length;
          }
        }
        return length;
      })();
      const target = clampOffset(offset, total);
      let remaining = target;
      let point = null;
      const walk = (node) => {
        if (point) return;
        if (node.nodeType === 3) {
          if (remaining <= node.data.length) {
            point = { node, offset: remaining };
            return;
          }
          remaining -= node.data.length;
          return;
        }
        if (node.nodeType !== 1) return;
        if (node.hasAttribute("data-dx-atom")) {
          const parent = node.parentNode;
          const atomIndex = Array.prototype.indexOf.call(
            parent.childNodes,
            node,
          );
          if (remaining === 0) {
            point = { node: parent, offset: atomIndex };
            return;
          }
          remaining -= 1;
          if (remaining === 0) {
            point = { node: parent, offset: atomIndex + 1 };
            return;
          }
          return;
        }
        for (const child of node.childNodes) walk(child);
      };
      for (const child of blockElement.childNodes) {
        walk(child);
        if (point) break;
      }
      return (
        point || { node: blockElement, offset: blockElement.childNodes.length }
      );
    }

    /** Restore a model selection in the rendered DOM. */
    function placeSelection(rootElement, selection) {
      if (typeof window === "undefined" || !rootElement || !selection)
        return false;
      const startElement = findBlockElement(
        rootElement,
        selection.start.blockId,
      );
      const endElement = findBlockElement(
        rootElement,
        selection.end ? selection.end.blockId : selection.start.blockId,
      );
      if (!startElement || !endElement) return false;
      beginSelectionWrite();
      try {
        if (typeof startElement.focus === "function") {
          try {
            startElement.focus({ preventScroll: true });
          } catch (_error) {
            startElement.focus();
          }
        }
      } finally {
        endSelectionWrite();
      }
      const ownerDocument = rootElement.ownerDocument;
      const range = ownerDocument.createRange();
      const startPoint = textPointAtOffset(
        startElement,
        selection.start.offset || 0,
      );
      const endPoint = textPointAtOffset(
        endElement,
        (selection.end ? selection.end.offset : selection.start.offset) || 0,
      );
      try {
        range.setStart(startPoint.node, startPoint.offset);
        range.setEnd(endPoint.node, endPoint.offset);
      } catch (_error) {
        return false;
      }
      const domSelection = ownerDocument.getSelection
        ? ownerDocument.getSelection()
        : window.getSelection();
      if (!domSelection) return false;
      beginSelectionWrite();
      try {
        domSelection.removeAllRanges();
        domSelection.addRange(range);
      } finally {
        endSelectionWrite();
      }
      return true;
    }

    /* ------------------------------------------------------- components */

    /**
     * Model point under a viewport coordinate.
     *
     * `caretRangeFromPoint` is Chromium's; the fallback walks every rendered
     * block and maps the click to the nearest text position. Returning null
     * means the pointer is not over the document at all.
     */
    function selectionPointAt(rootElement, clientX, clientY) {
      if (
        typeof document === "undefined" ||
        !rootElement ||
        typeof rootElement.querySelectorAll !== "function"
      )
        return null;
      const candidates = Array.from(
        rootElement.querySelectorAll("[data-block-id]"),
      ).filter(
        (element) =>
          element.tagName !== "TABLE" &&
          element.hasAttribute("contenteditable"),
      );
      if (candidates.length === 0) return null;
      let blockElement = null;
      if (typeof document.elementFromPoint === "function") {
        const hit = document.elementFromPoint(clientX, clientY);
        const closest =
          hit && typeof hit.closest === "function"
            ? hit.closest("[data-block-id]")
            : null;
        if (closest && rootElement.contains(closest)) blockElement = closest;
      }
      if (!blockElement) {
        let best = null;
        for (const candidate of candidates) {
          const rect = candidate.getBoundingClientRect();
          if (!rect.width && !rect.height) continue;
          const vertical =
            clientY < rect.top
              ? rect.top - clientY
              : clientY > rect.bottom
                ? clientY - rect.bottom
                : 0;
          if (!best || vertical < best.distance)
            best = { element: candidate, distance: vertical };
        }
        blockElement = best ? best.element : null;
      }
      if (!blockElement) return null;
      const rect = blockElement.getBoundingClientRect();
      const x = Math.min(Math.max(clientX, rect.left + 1), rect.right - 1);
      const y = Math.min(Math.max(clientY, rect.top + 1), rect.bottom - 1);
      let offset = 0;
      if (typeof document.caretRangeFromPoint === "function") {
        const range = document.caretRangeFromPoint(x, y);
        if (range && blockElement.contains(range.startContainer))
          offset = offsetInBlockElement(
            blockElement,
            range.startContainer,
            range.startOffset,
          );
      } else if (typeof document.caretPositionFromPoint === "function") {
        const caret = document.caretPositionFromPoint(x, y);
        if (caret && blockElement.contains(caret.offsetNode))
          offset = offsetInBlockElement(
            blockElement,
            caret.offsetNode,
            caret.offset,
          );
      }
      return {
        blockId: blockElement.getAttribute("data-block-id") || "",
        offset,
      };
    }

    /** Build a DOM Range between two model points (either order). */
    function domRangeForPoints(rootElement, fromPoint, toPoint) {
      if (typeof document === "undefined" || !rootElement) return null;
      const fromElement = findBlockElement(rootElement, fromPoint.blockId);
      const toElement = findBlockElement(rootElement, toPoint.blockId);
      if (!fromElement || !toElement) return null;
      const start = textPointAtOffset(fromElement, fromPoint.offset || 0);
      const end = textPointAtOffset(toElement, toPoint.offset || 0);
      const range = rootElement.ownerDocument.createRange();
      try {
        range.setStart(start.node, start.offset);
        range.setEnd(end.node, end.offset);
      } catch (_error) {
        return null;
      }
      return range;
    }

    /**
     * Depth counter for programmatic selection writes. Focusing a paragraph
     * fires a focus event, and the panel's focus handler reads the *old* DOM
     * selection; without this guard that read overwrites the selection the
     * caller is in the middle of writing (Ctrl+A appeared to do nothing).
     */
    let selectionWriteDepth = 0;

    function beginSelectionWrite() {
      selectionWriteDepth += 1;
    }

    function endSelectionWrite() {
      selectionWriteDepth = Math.max(0, selectionWriteDepth - 1);
    }

    /** Replace the browser selection with a model range, without stealing scroll. */
    function selectDomRange(rootElement, range) {
      if (!range || typeof window === "undefined" || !window.getSelection)
        return false;
      const domSelection = window.getSelection();
      if (!domSelection) return false;
      domSelection.removeAllRanges();
      domSelection.addRange(range);
      return true;
    }

    const PT_PER_PX = 96 / 72;

    function pxFromPt(value) {
      return Math.max(0, Math.round(value * PT_PER_PX));
    }

    function blocksToText(blocks) {
      return (blocks || [])
        .map((block) =>
          block && block.type === "paragraph" ? paragraphText(block) : "",
        )
        .filter((text, index, all) => text || index < all.length - 1)
        .join("\n");
    }

    function headerFooterText(section, part) {
      if (!section) return "";
      const set = part === "header" ? section.headers : section.footers;
      if (!set) return "";
      const variant = set.default || set.first || set.even || null;
      return variant ? blocksToText(variant.blocks) : "";
    }

    function itemGate(item, ctx) {
      if (item.reason) return item.reason;
      if (typeof item.disabled === "function") {
        const reason = item.disabled(ctx);
        if (reason) return reason;
      }
      if (item.needs === "document" && !ctx.document) return REASONS.noDocument;
      if (item.editing && ctx.readOnly) return REASONS.readOnly;
      if (item.editing && ctx.offline) return REASONS.offline;
      if (item.editing && !ctx.document) return REASONS.noDocument;
      if (item.remote && ctx.offline) return REASONS.offline;
      if (item.needs === "selection" && !ctx.formatting.hasRange)
        return REASONS.selectionRequired;
      if (item.needs === "caret" && !ctx.formatting.blockId)
        return REASONS.caretRequired;
      if (item.needs === "block" && !ctx.formatting.blockId)
        return REASONS.noDocument;
      return null;
    }

    /**
     * Tooltip for one ribbon tool: the reason when disabled, otherwise the
     * explicit title or label with the editor's keyboard shortcut appended.
     */
    function ribbonToolTitle(item, reason, label, ctx) {
      if (reason) return reason;
      const custom =
        typeof item.title === "function" ? item.title(ctx) : item.title;
      const base = custom || label;
      const overrides = (ctx && ctx.shortcutOverrides) || {};
      const shortcut = overrides[item.run] || SHORTCUTS[item.run];
      if (shortcut && String(base).indexOf(shortcut) === -1)
        return `${base} (${shortcut})`;
      return base;
    }

    function RibbonTool(props) {
      const item = props.item;
      const ctx = props.ctx;
      const api = props.api;
      const reason = itemGate(item, ctx);
      const label =
        typeof item.label === "function" ? item.label(ctx) : item.label;
      const pressed =
        typeof item.pressed === "function"
          ? Boolean(item.pressed(ctx))
          : undefined;
      const title = ribbonToolTitle(item, reason, label, ctx);
      const showLabel = props.showLabels === true && item.size !== "tiny";
      const className = [
        "dx-tool",
        item.size === "small" ? "dx-tool-small" : "",
        item.size === "tiny" ? "dx-tool-tiny" : "",
        item.swatch ? "dx-tool-swatch" : "",
      ]
        .filter(Boolean)
        .join(" ");
      const children = [];
      children.push(
        h(Icon, { key: "icon", name: item.icon || "page", size: 22 }),
      );
      if (item.swatch)
        children.push(
          h("span", {
            key: "swatch",
            className: "dx-color-bar",
            style: { background: item.swatch(ctx) || "transparent" },
          }),
        );
      // Icon-only by default: the label moves into `title`/`aria-label` and is
      // rendered as visible text only after View → Show labels.
      if (showLabel)
        children.push(
          h("span", { key: "label", className: "dx-tool-label" }, label),
        );
      return h(
        "button",
        {
          type: "button",
          className,
          title,
          "aria-label": label,
          "aria-pressed": pressed,
          disabled: Boolean(reason),
          onClick: () => {
            if (reason) return;
            api.onAction(item.run, item.arg);
          },
        },
        children,
      );
    }

    function RibbonSelect(props) {
      const item = props.item;
      const ctx = props.ctx;
      const api = props.api;
      const reason = itemGate(item, ctx);
      const options =
        typeof item.options === "function"
          ? item.options(ctx)
          : item.options || [];
      const value =
        typeof item.value === "function" ? item.value(ctx) : item.value;
      const title =
        reason ||
        (typeof item.title === "function" ? item.title(ctx) : item.title) ||
        item.label;
      return h(
        "label",
        { className: "dx-field", title, style: { gap: 2 } },
        h(
          "select",
          {
            className: "dx-select",
            value: value === undefined || value === null ? "" : String(value),
            "aria-label": item.label,
            title,
            disabled: Boolean(reason),
            style: item.width ? { maxWidth: item.width } : undefined,
            onChange: (event) => api.onAction(item.run, event.target.value),
          },
          options.map((option) =>
            h(
              "option",
              {
                key: String(option.value === undefined ? option : option.value),
                value: String(
                  option.value === undefined ? option : option.value,
                ),
              },
              String(option.label === undefined ? option : option.label),
            ),
          ),
        ),
      );
    }

    function RibbonStepper(props) {
      const item = props.item;
      const ctx = props.ctx;
      const api = props.api;
      const reason = itemGate(item, ctx);
      const value =
        typeof item.value === "function" ? item.value(ctx) : item.value;
      const step = item.step || 1;
      const commit = (next) => {
        const bounded = Math.max(
          item.min || 1,
          Math.min(item.max || 400, next),
        );
        api.onAction(item.run, String(bounded));
      };
      return h(
        "span",
        {
          className: "dx-stepper",
          title: reason || item.title || item.label,
        },
        h(
          "button",
          {
            type: "button",
            "aria-label": `Decrease ${item.label}`,
            disabled: Boolean(reason),
            onClick: () => commit(Number(value || item.min || 1) - step),
          },
          h(Icon, { name: "minus", size: 12 }),
        ),
        h("input", {
          key: `size-${value === undefined ? "" : String(value)}`,
          type: "text",
          inputMode: "decimal",
          "aria-label": item.label,
          defaultValue: value === undefined ? "" : String(value),
          disabled: Boolean(reason),
          // Commit on blur/Enter only: committing on every keystroke would post
          // an op and refetch for each digit, stealing focus mid-entry.
          onBlur: (event) => {
            const parsed = Number(String(event.target.value).replace(",", "."));
            if (Number.isFinite(parsed) && parsed > 0)
              api.onAction(item.run, String(parsed));
          },
          onKeyDown: (event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              event.currentTarget.blur();
            }
          },
        }),
        h(
          "button",
          {
            type: "button",
            "aria-label": `Increase ${item.label}`,
            disabled: Boolean(reason),
            onClick: () => commit(Number(value || 0) + step),
          },
          h(Icon, { name: "plus", size: 12 }),
        ),
      );
    }

    function StyleGallery(props) {
      const ctx = props.ctx;
      const api = props.api;
      const showLabels = ctx.showLabels === true;
      return h(
        "div",
        { className: "dx-gallery", role: "group", "aria-label": "Styles" },
        ctx.styleGallery.map((style) => {
          const reason = ctx.readOnly
            ? REASONS.readOnly
            : ctx.offline
              ? REASONS.offline
              : style.available
                ? null
                : `This document has no style named "${style.label}".`;
          // The "Aa" preview stays; the style name moves into the tooltip
          // unless the user restores labels.
          const children = [h("b", { key: "preview" }, "Aa")];
          if (showLabels) children.push(style.label);
          return h(
            "button",
            {
              key: style.id,
              type: "button",
              className: "dx-style-card",
              title: reason || `Apply the ${style.label} style (Aa preview)`,
              "aria-label": `Apply the ${style.label} style`,
              "aria-pressed": ctx.formatting.styleId === style.id,
              disabled: Boolean(reason),
              onClick: () => api.onAction("applyStyle", style.id),
            },
            children,
          );
        }),
      );
    }

    function RibbonGroup(props) {
      const group = props.group;
      const ctx = props.ctx;
      const api = props.api;
      const items = group.items || [];
      const isColumn = group.column === true;
      const children = items.map((item, index) => {
        if (item.kind === "select")
          return h(RibbonSelect, {
            key: item.id || `select-${index}`,
            item,
            ctx,
            api,
          });
        if (item.kind === "stepper")
          return h(RibbonStepper, {
            key: item.id || `stepper-${index}`,
            item,
            ctx,
            api,
          });
        if (item.kind === "gallery")
          return h(StyleGallery, { key: "gallery", ctx, api });
        if (item.kind === "note")
          return h(
            "p",
            { key: item.id || `note-${index}`, className: "dx-ribbon-note" },
            typeof item.text === "function" ? item.text(ctx) : item.text,
          );
        if (item.kind === "readout")
          return h(
            "div",
            {
              key: item.id || `readout-${index}`,
              className: "dx-file-meta",
              style: { alignSelf: "center", maxWidth: 220 },
            },
            typeof item.text === "function" ? item.text(ctx) : item.text,
          );
        return h(RibbonTool, {
          key: item.id || item.run || `tool-${index}`,
          item,
          ctx,
          api,
          showLabels: ctx.showLabels === true,
        });
      });
      return h(
        "div",
        { className: "dx-group", role: "group", "aria-label": group.label },
        h(
          "div",
          {
            className: isColumn ? "dx-group-items dx-column" : "dx-group-items",
          },
          children,
        ),
        h("span", { className: "dx-group-label" }, group.label),
      );
    }

    function tool(run, icon, label, extra) {
      return Object.assign({ kind: "tool", run, icon, label }, extra || {});
    }

    function homeGroups(ctx) {
      const formatting = ctx.formatting;
      return [
        {
          id: "undo",
          label: "Undo",
          items: [
            tool("undo", "undo", "Undo", {
              title: "Undo (Ctrl+Z)",
              disabled: () => (ctx.canUndo ? null : "Nothing to undo."),
              remote: true,
            }),
            tool("redo", "redo", "Redo", {
              title: "Redo (Ctrl+Y)",
              disabled: () => (ctx.canRedo ? null : "Nothing to redo."),
              remote: true,
            }),
          ],
        },
        {
          id: "clipboard",
          label: "Clipboard",
          items: [
            tool("cut", "cut", "Cut", { needs: "selection", editing: true }),
            tool("copy", "copy", "Copy", { needs: "selection" }),
            tool("paste", "paste", "Paste", { editing: true }),
          ],
        },
        {
          id: "font",
          label: "Font",
          items: [
            {
              kind: "select",
              id: "fontFamily",
              label: "Font family",
              title:
                "Font family for the selection, or for the whole paragraph when the caret is collapsed",
              run: "fontFamily",
              width: 150,
              editing: true,
              needs: "caret",
              value: () => formatting.fontFamily || "",
              options: () =>
                ctx.fontFamilies.map((family) => ({
                  value: family,
                  label: family,
                })),
            },
            {
              kind: "stepper",
              id: "fontSize",
              label: "Font size",
              run: "fontSize",
              min: 1,
              max: 400,
              step: 1,
              editing: true,
              needs: "caret",
              value: () => formatting.fontSize,
            },
            tool("bold", "bold", "Bold", {
              pressed: () => formatting.bold,
              editing: true,
              needs: "caret",
              title: (toolCtx) =>
                `Bold (${toolCtx.boldReserved ? "Ctrl+Shift+B" : "Ctrl+B"}): applies to the selection, or to the current paragraph when the caret is collapsed`,
            }),
            tool("italic", "italic", "Italic", {
              pressed: () => formatting.italic,
              editing: true,
              needs: "caret",
              title: "Italic (Ctrl+I)",
            }),
            tool("underline", "underline", "Underline", {
              pressed: () => formatting.underline,
              editing: true,
              needs: "caret",
              title: "Underline (Ctrl+U)",
            }),
            tool("strike", "strike", "Strikethrough", {
              pressed: () => formatting.strike,
              editing: true,
              needs: "caret",
            }),
            tool("subscript", "subscript", "Subscript", {
              pressed: () => formatting.vertAlign === "subscript",
              editing: true,
              needs: "caret",
            }),
            tool("superscript", "superscript", "Superscript", {
              pressed: () => formatting.vertAlign === "superscript",
              editing: true,
              needs: "caret",
            }),
            tool("textColor", "fontSet", "Text colour", {
              swatch: () => formatting.color,
              editing: true,
              needs: "caret",
              title: "Text colour for the selected text",
            }),
            tool("highlight", "shading", "Highlight", {
              swatch: () => formatting.highlight,
              editing: true,
              needs: "caret",
              title: "Highlight colour for the selected text",
            }),
            tool("clearFormatting", "clearFormat", "Clear", {
              editing: true,
              needs: "caret",
              title: "Clear direct formatting from the selection",
            }),
          ],
        },
        {
          id: "paragraph",
          label: "Paragraph",
          items: [
            tool("bullets", "bulletList", "Bullets", {
              editing: true,
              needs: "block",
              title: "Insert a new bulleted list at the caret",
            }),
            tool("numbering", "numberList", "Numbering", {
              editing: true,
              needs: "block",
              title: "Insert a new numbered list at the caret",
            }),
            tool("multilevel", "multilevelList", "Multilevel", {
              editing: true,
              needs: "block",
              title: "Insert a new multilevel list at the caret",
            }),
            tool("decreaseIndent", "indentDecrease", "Decrease indent", {
              editing: true,
              needs: "block",
            }),
            tool("increaseIndent", "indentIncrease", "Increase indent", {
              editing: true,
              needs: "block",
            }),
            tool("alignLeft", "alignLeft", "Align left", {
              pressed: () => formatting.alignment === "left",
              editing: true,
              needs: "block",
            }),
            tool("alignCenter", "alignCenter", "Centre", {
              pressed: () => formatting.alignment === "center",
              editing: true,
              needs: "block",
            }),
            tool("alignRight", "alignRight", "Align right", {
              pressed: () => formatting.alignment === "right",
              editing: true,
              needs: "block",
            }),
            tool("alignJustify", "alignJustify", "Justify", {
              pressed: () => formatting.alignment === "justify",
              editing: true,
              needs: "block",
            }),
            {
              kind: "select",
              id: "lineSpacing",
              label: "Line spacing",
              run: "lineSpacing",
              width: 74,
              editing: true,
              needs: "block",
              value: () => String(formatting.lineSpacing || 1),
              options: ctx.lineSpacingOptions.map((option) => ({
                value: String(option.value),
                label: option.label,
              })),
            },
            {
              kind: "select",
              id: "spaceBefore",
              label: "Space before",
              run: "spaceBefore",
              width: 74,
              editing: true,
              needs: "block",
              value: () => String(formatting.spaceBefore || 0),
              options: ["0", "4", "6", "10", "12", "18", "24"].map((value) => ({
                value,
                label: `Before ${value}pt`,
              })),
            },
            {
              kind: "select",
              id: "spaceAfter",
              label: "Space after",
              run: "spaceAfter",
              width: 74,
              editing: true,
              needs: "block",
              value: () => String(formatting.spaceAfter || 0),
              options: ["0", "4", "6", "10", "12", "18", "24"].map((value) => ({
                value,
                label: `After ${value}pt`,
              })),
            },
            tool("paragraphBorders", "borders", "Borders", {
              reason: REASONS.paragraphBorders,
            }),
            tool("paragraphShading", "shading", "Shading", {
              reason: REASONS.paragraphBorders,
            }),
          ],
        },
        {
          id: "styles",
          label: "Styles",
          items: [{ kind: "gallery" }],
        },
        {
          id: "editing",
          label: "Editing",
          items: [
            tool("find", "find", "Find", {
              title: "Find in the document (Ctrl+F)",
            }),
            tool("replace", "replace", "Replace", {
              title: "Find and replace",
            }),
            tool("selectAll", "selectAll", "Select all", {
              title: "Select the whole document (Ctrl+A)",
            }),
          ],
        },
      ];
    }

    function insertGroups() {
      return [
        {
          id: "pages",
          label: "Pages",
          items: [
            tool("pageBreak", "pageBreak", "Page break", {
              editing: true,
              needs: "block",
            }),
            tool("sectionBreak", "sectionBreak", "Section break", {
              editing: true,
              needs: "block",
            }),
          ],
        },
        {
          id: "tables",
          label: "Tables",
          items: [
            tool("insertTable", "table", "Table", {
              editing: true,
              needs: "block",
              size: "small",
            }),
          ],
        },
        {
          id: "illustrations",
          label: "Illustrations",
          items: [
            tool("insertImage", "image", "Pictures", {
              editing: true,
              needs: "block",
              size: "small",
            }),
            tool("shape", "shape", "Shapes", { reason: REASONS.shapes }),
          ],
        },
        {
          id: "links",
          label: "Links",
          items: [
            tool("hyperlink", "link", "Link", {
              editing: true,
              needs: "selection",
            }),
          ],
        },
        {
          id: "headerFooter",
          label: "Header & Footer",
          items: [
            tool("editHeader", "header", "Header", {
              editing: true,
            }),
            tool("editFooter", "footer", "Footer", {
              editing: true,
            }),
            tool("pageNumber", "pageNumber", "Page number", {
              reason: REASONS.pageNumber,
            }),
          ],
        },
        {
          id: "text",
          label: "Text",
          items: [
            tool("insertToc", "toc", "Table of Contents", {
              editing: true,
            }),
            tool("insertFootnote", "footnote", "Footnote", {
              editing: true,
              needs: "block",
            }),
            tool("insertEquation", "equation", "Equation", {
              editing: true,
              needs: "block",
            }),
          ],
        },
        {
          id: "symbols",
          label: "Symbols",
          items: [
            tool("symbol", "symbol", "Symbol", { reason: REASONS.symbols }),
          ],
        },
      ];
    }

    function designGroups(ctx) {
      return [
        {
          id: "documentFormatting",
          label: "Document Formatting",
          items: [
            tool("pageColor", "pageColor", "Page colour", {
              reason: REASONS.pageColor,
            }),
            {
              kind: "select",
              id: "paragraphSpacing",
              label: "Paragraph spacing",
              run: "paragraphSpacingPreset",
              width: 140,
              editing: true,
              needs: "block",
              value: () => "",
              options: [{ value: "", label: "Paragraph spacing…" }].concat(
                ctx.spacingPresets.map((preset) => ({
                  value: preset.id,
                  label: preset.label,
                })),
              ),
            },
            {
              kind: "select",
              id: "designFont",
              label: "Font set",
              run: "designFont",
              width: 150,
              editing: true,
              needs: "block",
              value: () => "",
              options: [{ value: "", label: "Font set…" }].concat(
                ctx.fontFamilies
                  .slice(0, 12)
                  .map((family) => ({ value: family, label: family })),
              ),
            },
            tool("pageBorders", "pageBorders", "Page borders", {
              reason: REASONS.pageBorders,
            }),
            {
              kind: "note",
              id: "designNote",
              text: "Themes and document backgrounds come from the stored document; page colour and page borders are not exposed by the engine.",
            },
          ],
        },
      ];
    }

    function layoutGroups(ctx) {
      const section =
        (ctx.sections || [])[ctx.activeSection || 0] || ctx.sections[0] || null;
      const page = (section && section.page) || null;
      const currentMargins = page && page.margins ? page.margins : null;
      const matchingMargin = currentMargins
        ? MARGIN_PRESETS.find(
            (preset) =>
              Math.abs(preset.top - currentMargins.top) < 1 &&
              Math.abs(preset.right - currentMargins.right) < 1 &&
              Math.abs(preset.bottom - currentMargins.bottom) < 1 &&
              Math.abs(preset.left - currentMargins.left) < 1,
          )
        : null;
      const matchingSize = page
        ? PAGE_SIZES.find(
            (preset) =>
              Math.abs(preset.width - page.width) < 1 &&
              Math.abs(preset.height - page.height) < 1,
          ) ||
          PAGE_SIZES.find(
            (preset) =>
              Math.abs(preset.width - page.width) < 1 ||
              Math.abs(preset.height - page.width) < 1,
          )
        : null;
      return [
        {
          id: "pageSetup",
          label: "Page Setup",
          items: [
            {
              kind: "select",
              id: "margins",
              label: "Margins",
              run: "margins",
              width: 110,
              editing: true,
              value: () => (matchingMargin ? matchingMargin.id : "custom"),
              options: MARGIN_PRESETS.map((preset) => ({
                value: preset.id,
                label: preset.label,
              })).concat([{ value: "custom", label: "Custom…" }]),
            },
            {
              kind: "select",
              id: "orientation",
              label: "Orientation",
              run: "orientation",
              width: 96,
              editing: true,
              value: () =>
                page && page.orientation === "landscape"
                  ? "landscape"
                  : "portrait",
              options: [
                { value: "portrait", label: "Portrait" },
                { value: "landscape", label: "Landscape" },
              ],
            },
            {
              kind: "select",
              id: "pageSize",
              label: "Size",
              run: "pageSize",
              width: 84,
              editing: true,
              value: () => (matchingSize ? matchingSize.id : ""),
              options: PAGE_SIZES.map((preset) => ({
                value: preset.id,
                label: preset.label,
              })),
            },
            {
              kind: "select",
              id: "columns",
              label: "Columns",
              run: "columns",
              width: 84,
              editing: true,
              value: () =>
                section && section.columns && section.columns.count
                  ? String(section.columns.count)
                  : "1",
              options: [
                { value: "1", label: "One" },
                { value: "2", label: "Two" },
                { value: "3", label: "Three" },
              ],
            },
          ],
        },
        {
          id: "breaks",
          label: "Breaks",
          items: [
            tool("pageBreak", "pageBreak", "Page", {
              editing: true,
              needs: "block",
            }),
            tool("columnBreak", "columnBreak", "Column", {
              editing: true,
              needs: "block",
            }),
            tool("sectionBreak", "sectionBreak", "Section", {
              editing: true,
              needs: "block",
            }),
          ],
        },
        {
          id: "readout",
          label: "Page Size",
          items: [
            {
              kind: "readout",
              id: "pageSizeReadout",
              text: () =>
                `${pageSizeLabel(section)} — margins ${marginLabel(section)}`,
            },
          ],
        },
      ];
    }

    function referenceGroups() {
      return [
        {
          id: "toc",
          label: "Table of Contents",
          items: [
            tool("insertToc", "toc", "Table of Contents", {
              editing: true,
            }),
            tool("updateToc", "refresh", "Update TOC", {
              title: "Refetch the document so the contents list is rebuilt",
            }),
          ],
        },
        {
          id: "notes",
          label: "Footnotes",
          items: [
            tool("insertFootnote", "footnote", "Footnote", {
              editing: true,
              needs: "block",
            }),
            tool("insertEndnote", "endnote", "Endnote", {
              editing: true,
              needs: "block",
            }),
          ],
        },
        {
          id: "citations",
          label: "Citations & Bibliography",
          items: [
            tool("bookmark", "bookmark", "Bookmark", {
              reason: REASONS.bookmark,
            }),
            tool("caption", "caption", "Caption", { reason: REASONS.caption }),
            tool("crossReference", "crossReference", "Cross-reference", {
              reason: REASONS.crossReference,
            }),
          ],
        },
      ];
    }

    function viewGroups(ctx) {
      const groups = [
        {
          id: "views",
          label: "Views",
          items: [
            tool("toggleReadOnly", "lock", "Read only", {
              pressed: () => ctx.readOnly,
            }),
            tool("toggleFocus", "focus", "Focus", {
              pressed: () => ctx.focusMode,
              title: "Word-like focus mode (hides the document rail)",
            }),
            tool("toggleRuler", "ruler", "Ruler", {
              pressed: () => ctx.showRuler,
            }),
            tool("toggleMarks", "marks", "Formatting marks", {
              pressed: () => ctx.showMarks,
            }),
            tool("togglePageSeparators", "pageSeparators", "Page separators", {
              pressed: () => ctx.showPageSeparators,
            }),
            tool("toggleRail", "sidebar", "Document rail", {
              pressed: () => ctx.showRail,
            }),
            tool("toggleLabels", "labels", "Show labels", {
              pressed: () => ctx.showLabels,
              title:
                "Show text labels on ribbon buttons (saved in this browser); off by default",
            }),
          ],
        },
        {
          id: "zoom",
          label: "Zoom",
          items: [
            tool("zoomOut", "zoomOut", "Zoom out", { size: "small" }),
            tool("zoomIn", "zoomIn", "Zoom in", { size: "small" }),
            tool("fitWidth", "fitWidth", "Fit width", { size: "small" }),
            {
              kind: "select",
              id: "zoomSelect",
              label: "Zoom level",
              run: "setZoom",
              width: 78,
              value: () => String(ctx.zoom),
              options: ZOOM_STEPS.map((step) => ({
                value: String(step),
                label: formatZoom(step),
              })),
            },
          ],
        },
      ];
      if (ctx.boldReserved) {
        groups.push({
          id: "shortcuts",
          label: "Keyboard",
          items: [
            {
              kind: "readout",
              id: "boldReservation",
              text: () => BOLD_RESERVED_NOTE,
              title: BOLD_RESERVED_NOTE,
            },
          ],
        });
      }
      return groups;
    }

    function ribbonGroups(tab, ctx) {
      switch (tab) {
        case "home":
          return homeGroups(ctx);
        case "insert":
          return insertGroups(ctx);
        case "design":
          return designGroups(ctx);
        case "layout":
          return layoutGroups(ctx);
        case "references":
          return referenceGroups(ctx);
        case "view":
          return viewGroups(ctx);
        default:
          return [];
      }
    }

    function TitleBar(props) {
      const ctx = props;
      const api = props.api;
      const dirty = props.dirty;
      const saveLabel =
        props.saveState === "saving"
          ? "Saving…"
          : props.saveState === "unsaved"
            ? "Unsaved changes"
            : props.saveState === "error"
              ? "Save failed"
              : "Saved";
      return h(
        "header",
        { className: "dx-titlebar", "aria-label": "Document title bar" },
        h(
          "span",
          { className: "dx-appdot" },
          h(Icon, { name: "grid", size: 15 }),
          "Documents",
        ),
        h("input", {
          className: "dx-title",
          value: props.title || "",
          "aria-label": "Document title",
          title: ctx.readOnly
            ? "The document is read-only."
            : "Edit the document title",
          disabled: !props.bootstrap || ctx.readOnly,
          onChange: (event) => api.onTitleChange(event.target.value),
          onBlur: () => api.onTitleCommit(),
          onKeyDown: (event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              api.onTitleCommit();
              event.currentTarget.blur();
            }
          },
        }),
        dirty
          ? h("span", { className: "dx-chip dx-chip-warn" }, "● Unsaved")
          : h("span", { className: "dx-chip dx-chip-saved" }, "✓ Saved"),
        h("span", { className: "dx-chip" }, saveLabel),
        ctx.readOnly ? h("span", { className: "dx-chip" }, "Read-only") : null,
        h("span", { className: "dx-titlebar-spacer" }),
        h(
          "label",
          { className: "dx-zoom-mini", title: "Zoom" },
          h(Icon, { name: "zoomOut", size: 13 }),
          h("input", {
            type: "range",
            min: 50,
            max: 200,
            step: 10,
            value: Math.round(props.zoom * 100),
            "aria-label": "Zoom",
            onChange: (event) => api.onZoom(Number(event.target.value) / 100),
          }),
          h("span", null, formatZoom(props.zoom)),
        ),
        h(
          "button",
          {
            type: "button",
            className: "dx-titlebar-btn",
            title: "Save (Ctrl+S)",
            "aria-label": "Save document",
            disabled: !props.bootstrap || ctx.readOnly || props.saving,
            onClick: () => api.onAction("save"),
          },
          h(Icon, { name: "save", size: 15 }),
        ),
        h(
          "button",
          {
            type: "button",
            className: "dx-titlebar-btn",
            title: "Focus mode",
            "aria-label": "Toggle focus mode",
            onClick: () => api.onAction("toggleFocus"),
          },
          h(Icon, { name: "focus", size: 15 }),
        ),
      );
    }

    function TabRow(props) {
      const api = props.api;
      const tabRefs = React.useRef({});
      const move = (event, index) => {
        const keys = ["ArrowRight", "ArrowLeft", "Home", "End"];
        if (keys.indexOf(event.key) === -1) return;
        event.preventDefault();
        let next = index;
        if (event.key === "ArrowRight") next = (index + 1) % TABS.length;
        if (event.key === "ArrowLeft")
          next = (index - 1 + TABS.length) % TABS.length;
        if (event.key === "Home") next = 0;
        if (event.key === "End") next = TABS.length - 1;
        api.onTab(TABS[next].id);
        const target = tabRefs.current[TABS[next].id];
        if (target && typeof target.focus === "function") target.focus();
      };
      return h(
        "div",
        { className: "dx-tabrow" },
        h(
          "div",
          {
            className: "dx-tabs",
            role: "tablist",
            "aria-label": "Ribbon tabs",
          },
          TABS.map((entry, index) =>
            h(
              "button",
              {
                key: entry.id,
                type: "button",
                role: "tab",
                id: `dx-tab-${entry.id}`,
                className: "dx-tab",
                "aria-selected": props.tab === entry.id ? true : false,
                "aria-controls": "dx-ribbon",
                tabIndex: props.tab === entry.id ? 0 : -1,
                ref: (node) => {
                  tabRefs.current[entry.id] = node;
                },
                onClick: () => api.onTab(entry.id),
                onKeyDown: (event) => move(event, index),
              },
              entry.label,
            ),
          ),
        ),
        h(
          "button",
          {
            type: "button",
            className: "dx-ribbon-toggle",
            "aria-label": props.ribbonCollapsed
              ? "Expand the ribbon"
              : "Collapse the ribbon",
            "aria-expanded": props.ribbonCollapsed ? false : true,
            title: props.ribbonCollapsed
              ? "Expand the ribbon"
              : "Collapse the ribbon",
            onClick: () => api.onToggleRibbon(),
          },
          h(Icon, {
            name: props.ribbonCollapsed ? "chevronDown" : "chevronUp",
            size: 14,
          }),
        ),
      );
    }

    /**
     * The ribbon's drag handle: a 6px strip along the bottom edge of the dock.
     * It is a real separator widget: pointer drags resize live, the arrow keys
     * move by 16px, Home/End jump to the bounds, and Enter (or a double click)
     * returns to the 1/10-of-viewport default.
     */
    function RibbonResizer(props) {
      const api = props.api;
      const dragRef = React.useRef(null);
      const min = RIBBON_MIN_HEIGHT;
      const max =
        typeof props.max === "number" ? props.max : ribbonMaxHeight(0);
      const value =
        typeof props.height === "number"
          ? props.height
          : defaultRibbonHeight(viewportHeight());
      const currentHeight = () => {
        const node = api.ribbonRef && api.ribbonRef.current;
        if (node && typeof node.getBoundingClientRect === "function") {
          const height = node.getBoundingClientRect().height;
          if (height > 0) return height;
        }
        return typeof props.height === "number"
          ? props.height
          : defaultRibbonHeight(viewportHeight());
      };
      const apply = (value, persist) => {
        if (typeof api.onRibbonHeight === "function")
          api.onRibbonHeight(value, persist === true);
      };
      const onPointerDown = (event) => {
        if (typeof event.button === "number" && event.button !== 0) return;
        event.preventDefault();
        const startHeight = currentHeight();
        dragRef.current = {
          pointerId: event.pointerId,
          startY: event.clientY,
          startHeight,
          value: startHeight,
        };
        const node = event.currentTarget;
        if (
          node &&
          typeof node.setPointerCapture === "function" &&
          typeof event.pointerId === "number"
        ) {
          try {
            node.setPointerCapture(event.pointerId);
          } catch (_error) {
            /* A synthetic pointer may not be capturable; dragging still works. */
          }
        }
      };
      const onPointerMove = (event) => {
        const drag = dragRef.current;
        if (!drag) return;
        event.preventDefault();
        const next = drag.startHeight + (event.clientY - drag.startY);
        drag.value = clampRibbonHeight(next, viewportHeight());
        apply(drag.value, false);
      };
      const finishDrag = (event) => {
        const drag = dragRef.current;
        if (!drag) return;
        dragRef.current = null;
        const node = event.currentTarget;
        if (
          node &&
          typeof node.releasePointerCapture === "function" &&
          typeof drag.pointerId === "number"
        ) {
          try {
            node.releasePointerCapture(drag.pointerId);
          } catch (_error) {
            /* Already released, or never captured. */
          }
        }
        apply(drag.value, true);
      };
      const onKeyDown = (event) => {
        const value = currentHeight();
        let next = null;
        let reset = false;
        if (event.key === "ArrowDown") next = value + RIBBON_KEYBOARD_STEP;
        else if (event.key === "ArrowUp") next = value - RIBBON_KEYBOARD_STEP;
        else if (event.key === "Home") next = min;
        else if (event.key === "End") next = max;
        else if (
          event.key === "Enter" ||
          event.key === " " ||
          event.key === "Spacebar"
        )
          reset = true;
        else return;
        event.preventDefault();
        event.stopPropagation();
        apply(reset ? null : clampRibbonHeight(next, viewportHeight()), true);
      };
      return h(
        "div",
        {
          className: "dx-ribbon-handle",
          role: "separator",
          "aria-orientation": "horizontal",
          "aria-label": "Ribbon height",
          "aria-valuenow": Math.round(value),
          "aria-valuemin": min,
          "aria-valuemax": max,
          "aria-valuetext": `${Math.round(value)} pixels`,
          title: "Drag to resize the ribbon; double-click to reset",
          tabIndex: 0,
          onPointerDown,
          onPointerMove,
          onPointerUp: finishDrag,
          onPointerCancel: finishDrag,
          onDoubleClick: () => apply(null, true),
          onKeyDown,
        },
        h("span", { className: "dx-ribbon-grip", "aria-hidden": "true" }),
      );
    }

    /** Inline height for the ribbon: nothing when the CSS default applies. */
    function ribbonStyle(ctx) {
      return typeof ctx.ribbonHeight === "number"
        ? { height: `${ctx.ribbonHeight}px` }
        : undefined;
    }

    function RibbonDock(props) {
      return h(
        "div",
        { className: "dx-ribbon-dock", hidden: props.hidden === true },
        props.children,
      );
    }

    function Ribbon(props) {
      const ctx = props;
      const api = props.api;
      if (props.ribbonCollapsed)
        return h(
          RibbonDock,
          { hidden: true },
          h("div", {
            className: "dx-ribbon",
            id: "dx-ribbon",
            role: "tabpanel",
            "aria-label": `${props.tab} ribbon (collapsed)`,
            hidden: true,
          }),
        );
      if (props.tab === "file") return h(FileRibbon, { ctx, api });
      const groups = ribbonGroups(props.tab, ctx);
      return h(
        RibbonDock,
        null,
        h(
          "div",
          {
            className: "dx-ribbon",
            id: "dx-ribbon",
            role: "tabpanel",
            "aria-label": `${props.tab} ribbon`,
            style: ribbonStyle(props),
            ref: api.ribbonRef,
          },
          groups.map((group) =>
            h(RibbonGroup, { key: group.id, group, ctx, api }),
          ),
        ),
        h(RibbonResizer, {
          api,
          height: props.ribbonHeightValue,
          max: props.ribbonMaxHeight,
        }),
      );
    }

    function FileActions(props) {
      const ctx = props.ctx;
      const api = props.api;
      const disabled = ctx.readOnly ? REASONS.readOnly : null;
      const remoteDisabled = ctx.offline ? REASONS.offline : null;
      const action = (run, label, icon, extra, ariaLabel) => {
        const reason =
          (extra && extra.reason) ||
          (extra && extra.editing ? disabled : null) ||
          (extra && extra.remote ? remoteDisabled : null) ||
          null;
        return h(
          "button",
          {
            key: run + label,
            type: "button",
            className: "dx-file-action",
            title: reason || label,
            "aria-label": ariaLabel || label,
            disabled: Boolean(reason),
            onClick: () => api.onAction(run),
          },
          h(Icon, { name: icon, size: 22 }),
          label,
        );
      };
      const roots = ctx.browseRoots || [];
      return h(
        "div",
        { className: "dx-file-col", style: { overflowY: "auto" } },
        h("h4", null, "Document"),
        h(
          "div",
          { className: "dx-file-actions" },
          action("newDocument", "New document…", "newDoc", {
            editing: true,
          }),
          // The in-panel chooser is the primary "choose file" path: it needs no
          // OS file dialog and lists folders and importable documents.
          action("openFileChooser", "Open file…", "open", {
            remote: true,
          }),
          action("openDocument", "Open from this device…", "device", {
            remote: true,
          }),
          action("chooseFolder", "Choose folder…", "folder", {
            remote: true,
          }),
          // A real (hidden) file input: the button above opens the native
          // file dialog and the Host receives the file through POST /open.
          h("input", {
            key: "file-input",
            ref: api.fileInputRef,
            className: "dx-file-input",
            type: "file",
            accept: DOCUMENT_ACCEPT,
            "aria-label": "Open a document from this device",
            onChange: api.onFilePicked,
          }),
          action("save", "Save", "save", {
            editing: true,
            remote: true,
          }),
          action("saveAs", "Save as…", "saveAs", {
            editing: true,
            remote: true,
          }),
          action("refresh", "Refresh list", "refresh", { remote: true }),
          action("deleteDocument", "Delete document", "trash", {
            editing: true,
            remote: true,
          }),
        ),
        h(
          "details",
          {
            className: "dx-advanced",
            open: ctx.advancedOpen === true,
            onToggle: (event) =>
              api.onToggleAdvanced(Boolean(event.currentTarget.open)),
          },
          h("summary", null, "Advanced: path"),
          h(
            "div",
            { className: "dx-advanced-body" },
            h(
              "label",
              { className: "dx-field" },
              "Document path",
              h("input", {
                className: "dx-input",
                type: "text",
                value: ctx.advancedPath || "",
                placeholder: "documents/report.docx",
                "aria-label": "Document path",
                onChange: (event) => api.onAdvancedPath(event.target.value),
              }),
            ),
            h(
              "button",
              {
                type: "button",
                className: "dx-btn",
                title: "Open the path typed above",
                "aria-label": "Open by path",
                onClick: () => api.onAction("openAdvancedPath"),
              },
              "Open by path…",
            ),
            roots.length
              ? h(
                  "div",
                  {
                    className: "dx-roots",
                    role: "group",
                    "aria-label": "Browse roots",
                  },
                  roots.map((root, index) =>
                    h(
                      "button",
                      {
                        key: `${root.path}-${index}`,
                        type: "button",
                        className: "dx-root-link",
                        title: `Browse ${root.path}`,
                        onClick: () => api.onAction("browseRoot", root.path),
                      },
                      root.label,
                    ),
                  ),
                )
              : h(
                  "button",
                  {
                    type: "button",
                    className: "dx-btn",
                    title: "Ask the Host for the browse roots",
                    onClick: () => api.onAction("loadBrowseRoots"),
                  },
                  "Load roots",
                ),
          ),
        ),
      );
    }

    function FileRibbon(props) {
      const ctx = props.ctx;
      const api = props.api;
      // The ribbon keeps its (default or user-chosen) height; the rich File
      // "backstage" opens directly below it as an overlay so it cannot change
      // that height. Escape or Close returns to the document.
      const [backstage, setBackstage] = React.useState(true);
      const closeBackstage = React.useCallback(() => {
        setBackstage(false);
        // Return focus to the File tab so keyboard users stay oriented.
        if (typeof document !== "undefined") {
          const tabButton = document.getElementById("dx-tab-file");
          if (tabButton && typeof tabButton.focus === "function")
            tabButton.focus();
        }
      }, []);
      React.useEffect(() => {
        if (!backstage) return undefined;
        // The backstage is a disclosure, not a focus trap: Escape must close it
        // wherever focus is (Word closes the Backstage the same way), including
        // while focus is still on the File tab button.
        const onKeyDown = (event) => {
          if (event.key !== "Escape") return;
          // An open dialog owns Escape: it closes itself and the backstage
          // stays behind it. Only close the backstage when no dialog is up.
          const target = event.target;
          if (
            target &&
            typeof target.closest === "function" &&
            target.closest(".dx-overlay")
          )
            return;
          if (
            typeof document !== "undefined" &&
            document.querySelector(".dx-dialog")
          )
            return;
          event.preventDefault();
          event.stopPropagation();
          closeBackstage();
        };
        if (typeof window !== "undefined" && window.addEventListener)
          window.addEventListener("keydown", onKeyDown, true);
        return () => {
          if (typeof window !== "undefined" && window.removeEventListener)
            window.removeEventListener("keydown", onKeyDown, true);
        };
      }, [backstage, closeBackstage]);
      const disabled = ctx.readOnly ? REASONS.readOnly : null;
      const remoteDisabled = ctx.offline ? REASONS.offline : null;
      const compact = (run, label, icon, extra) => {
        const reason =
          (extra && extra.reason) ||
          (extra && extra.editing ? disabled : null) ||
          (extra && extra.remote ? remoteDisabled : null) ||
          null;
        return h(
          "button",
          {
            key: run + label,
            type: "button",
            className: "dx-tool dx-tool-small",
            title: reason || label,
            "aria-label": label,
            disabled: Boolean(reason),
            onClick: () => api.onAction(run),
          },
          h(Icon, { name: icon, size: 22 }),
        );
      };
      return h(
        React.Fragment,
        null,
        h(
          RibbonDock,
          null,
          h(
            "div",
            {
              className: "dx-ribbon",
              id: "dx-ribbon",
              role: "tabpanel",
              "aria-label": "file ribbon",
              style: ribbonStyle(ctx),
              ref: api.ribbonRef,
            },
            h(
              "div",
              { className: "dx-file-compact" },
              compact("newDocument", "New document…", "newDoc", {
                editing: true,
              }),
              // Primary "choose file" entry point: the in-panel chooser, no OS
              // dialog involved.
              compact("openFileChooser", "Open file…", "open", {
                remote: true,
              }),
              compact("openDocument", "Open from this device…", "device", {
                remote: true,
              }),
              compact("chooseFolder", "Choose folder…", "folder", {
                remote: true,
              }),
              compact("save", "Save", "save", { editing: true, remote: true }),
              compact("saveAs", "Save as…", "saveAs", {
                editing: true,
                remote: true,
              }),
              compact("export", "Export PDF", "export", { remote: true }),
              h(
                "button",
                {
                  key: "more",
                  type: "button",
                  className: "dx-tool dx-tool-small",
                  title: backstage
                    ? "Hide File details"
                    : "More File commands…",
                  "aria-label": "More File commands",
                  "aria-expanded": backstage,
                  onClick: () => setBackstage((value) => !value),
                },
                h(Icon, { name: "chevronDown", size: 22 }),
              ),
            ),
          ),
          h(RibbonResizer, {
            api,
            height: ctx.ribbonHeightValue,
            max: ctx.ribbonMaxHeight,
          }),
        ),
        backstage
          ? h(
              "div",
              {
                className: "dx-backstage",
                role: "dialog",
                "aria-modal": "false",
                "aria-label": "File",
              },
              h(
                "div",
                { className: "dx-backstage-head" },
                h("h3", null, "File"),
                h(
                  "button",
                  {
                    type: "button",
                    className: "dx-tool dx-tool-small",
                    title: "Close File",
                    "aria-label": "Close File",
                    onClick: closeBackstage,
                  },
                  h(Icon, { name: "close", size: 22 }),
                ),
              ),
              h(
                "div",
                { className: "dx-file-panels" },
                h(FileActions, { ctx, api }),
                h(
                  "div",
                  { className: "dx-file-col" },
                  h("h4", null, "Export"),
                  h(
                    "div",
                    { className: "dx-file-actions" },
                    // Prefer the formats the Host's engine actually reports.
                    (ctx.formats &&
                    Array.isArray(ctx.formats.export) &&
                    ctx.formats.export.length
                      ? ctx.formats.export
                      : ["docx", "pdf", "html", "md"]
                    ).map((format) =>
                      h(
                        "button",
                        {
                          key: format,
                          type: "button",
                          className: "dx-file-action",
                          title: `Export ${String(format).toUpperCase()}`,
                          disabled: !ctx.bootstrap,
                          onClick: () => api.onAction("export", format),
                        },
                        h(Icon, { name: "export", size: 22 }),
                        String(format).toUpperCase(),
                      ),
                    ),
                  ),
                ),
                h(
                  "div",
                  { className: "dx-file-col", style: { minWidth: 210 } },
                  h("h4", null, "Information"),
                  h(
                    "div",
                    { className: "dx-file-meta" },
                    h("b", null, ctx.title || "No document"),
                    h("br"),
                    ctx.bootstrap
                      ? `Format ${ctx.sourceFormat || "docx"} · revision ${ctx.revision}`
                      : "Open or create a document to start editing.",
                  ),
                  h(
                    "div",
                    { className: "dx-kv" },
                    h("span", null, "Workspace"),
                    h("b", null, ctx.workspace || "-"),
                    h("span", null, "Pages"),
                    h("b", null, String(ctx.pageCount || 0)),
                    h("span", null, "Words"),
                    h("b", null, String((ctx.stats && ctx.stats.words) || 0)),
                    h("span", null, "Tables"),
                    h("b", null, String((ctx.stats && ctx.stats.tables) || 0)),
                    h("span", null, "Images"),
                    h("b", null, String((ctx.stats && ctx.stats.images) || 0)),
                    h("span", null, "Read-only"),
                    h("b", null, ctx.readOnly ? "yes" : "no"),
                  ),
                  ctx.offline
                    ? h(
                        "p",
                        {
                          className: "dx-file-meta",
                          style: { color: "var(--dx-danger)" },
                        },
                        "The Host editor API is unavailable; editing stays local.",
                      )
                    : null,
                  ctx.canPickDirectory === false
                    ? h(
                        "p",
                        { className: "dx-file-meta" },
                        "This Host exposes no native folder picker; use “Advanced: path”.",
                      )
                    : null,
                ),
                h(
                  "div",
                  { className: "dx-file-col", style: { flex: "1 1 auto" } },
                  h("h4", null, "Open documents"),
                  h(DocumentList, { ctx, api }),
                ),
              ),
            )
          : null,
      );
    }

    function DocumentList(props) {
      const ctx = props.ctx;
      const api = props.api;
      if (!ctx.documents || ctx.documents.length === 0)
        return h(
          "p",
          { className: "dx-empty" },
          "No open documents. Create one, or open a DOCX/Pages/DOC file by path.",
        );
      return h(
        "div",
        { className: "dx-doclist", role: "list" },
        ctx.documents.map((item) =>
          h(
            "button",
            {
              key: item.id,
              type: "button",
              className: "dx-docrow",
              role: "listitem",
              "aria-current": item.id === ctx.selectedId ? true : undefined,
              onClick: () => api.onOpenDocument(item.id),
            },
            h("strong", null, item.title || "Untitled document"),
            h(
              "span",
              null,
              [
                item.format,
                item.words ? `${item.words} words` : null,
                item.dirty ? "unsaved" : null,
              ]
                .filter(Boolean)
                .join(" · "),
            ),
          ),
        ),
      );
    }

    function DocumentRail(props) {
      const ctx = props;
      const api = props.api;
      return h(
        "aside",
        { className: "dx-rail", "aria-label": "Documents" },
        h(
          "div",
          { className: "dx-rail-head" },
          h("span", { className: "dx-rail-title" }, "Documents"),
          h(
            "button",
            {
              type: "button",
              className: "dx-status-btn",
              title: "Refresh the document list",
              onClick: () => api.onAction("refresh"),
            },
            h(Icon, { name: "refresh", size: 13 }),
          ),
          h(
            "button",
            {
              type: "button",
              className: "dx-status-btn",
              title: "Hide the document rail",
              onClick: () => api.onAction("toggleRail"),
            },
            h(Icon, { name: "close", size: 13 }),
          ),
        ),
        h(
          "div",
          { className: "dx-rail-row" },
          h(
            "button",
            {
              type: "button",
              className: "dx-btn",
              disabled: Boolean(ctx.readOnly),
              title: ctx.readOnly ? REASONS.readOnly : "Create a document",
              onClick: () => api.onAction("newDocument"),
            },
            h(Icon, { name: "newDoc", size: 14 }),
            "New",
          ),
          h(
            "button",
            {
              type: "button",
              className: "dx-btn",
              disabled: Boolean(ctx.readOnly),
              title: ctx.readOnly ? REASONS.readOnly : "Open a file by path",
              onClick: () => api.onAction("openByPath"),
            },
            h(Icon, { name: "open", size: 14 }),
            "Open",
          ),
        ),
        h(DocumentList, { ctx, api }),
        h(
          "div",
          { className: "dx-file-meta" },
          ctx.workspace ? `Workspace: ${ctx.workspace}` : "Workspace unknown",
          h("br"),
          ctx.formats && ctx.formats.export
            ? `Export: ${(ctx.formats.export || []).join(", ")}`
            : "",
        ),
      );
    }

    function Ruler(props) {
      const section = props.section;
      const width = section && section.page ? section.page.width : 595;
      const ticks = 8;
      const children = [
        h("div", { key: "gutter", className: "dx-ruler-gutter" }),
      ];
      for (let index = 0; index < ticks; index += 1) {
        children.push(
          h(
            "div",
            { key: `tick-${index}`, className: "dx-ruler-tick" },
            String(Math.round((width / ticks) * (index + 1))),
          ),
        );
      }
      return h(
        "div",
        { className: "dx-ruler", "aria-hidden": "true" },
        children,
      );
    }

    function renderTextWithMarks(text, showMarks, keyPrefix) {
      if (!showMarks || text.indexOf(" ") === -1) return text;
      const parts = text.split(" ");
      const nodes = [];
      parts.forEach((part, index) => {
        if (part) nodes.push(part);
        if (index < parts.length - 1)
          nodes.push(
            h(
              "span",
              { key: `${keyPrefix}-space-${index}`, className: "dx-space" },
              " ",
            ),
          );
      });
      return nodes;
    }

    function RunView(props) {
      const inline = props.inline;
      const index = props.index;
      const ctx = props.ctx;
      const text = inlineText(inline);
      const key = index;
      if (inline.type === "text") {
        const resolved = resolveRunProps(ctx.document, ctx.paragraph, inline);
        const link = resolved.hyperlink || {};
        const href = link.href ? sanitizeUrl(link.href) : null;
        const revision =
          inline.props && inline.props.revision ? inline.props.revision : null;
        const deleted = Boolean(revision && revision.kind === "delete");
        const inserted = Boolean(revision && revision.kind === "insert");
        const classes = ["dx-run"];
        if (href) classes.push("dx-link");
        if (deleted) classes.push("dx-run-deleted");
        if (inserted) classes.push("dx-run-inserted");
        const runProps = {
          key,
          className: classes.join(" "),
          "data-dx-inline": index,
          style: runStyleToCss(resolved),
        };
        if (href) {
          runProps.title = link.tooltip || href;
          runProps.onClick = (event) => {
            if (!event.ctrlKey && !event.metaKey) return;
            event.preventDefault();
            ctx.api.onOpenUrl(href);
          };
        } else if (revision) {
          runProps.title = deleted
            ? "Tracked deletion (still stored in the document)"
            : "Tracked insertion";
        }
        return h(
          "span",
          runProps,
          renderTextWithMarks(text, ctx.showMarks, `r${index}`),
        );
      }
      if (inline.type === "tab")
        return h("span", {
          key,
          className: "dx-atom dx-atom-tab",
          "data-dx-atom": "\t",
          "data-dx-inline": index,
          "aria-hidden": "true",
          title: "Tab",
        });
      if (inline.type === "break") {
        if (inline.kind === "page")
          return h("span", {
            key,
            className: "dx-atom dx-atom-pagebreak",
            "data-dx-atom": "\f",
            "data-dx-inline": index,
            title: "Page break",
          });
        // A hard line break (DOCX `<w:br/>`, Word's Shift+Enter) must be a real
        // line break, not an inline atom: a `br` inside the atom wrapper keeps
        // the one-character offset space (`data-dx-atom` counts it) while the
        // following text starts on a new line. The ↵ marker is drawn by CSS
        // only under View → Formatting marks.
        return h(
          "span",
          {
            key,
            className: "dx-atom dx-atom-break",
            "data-dx-atom": "\n",
            "data-dx-inline": index,
            title: "Line break",
          },
          h("br", null),
        );
      }
      if (inline.type === "image") {
        const asset = ctx.media.get(inline.mediaId);
        const url = asset ? mediaUrl(asset) : null;
        if (!url)
          return h(
            "span",
            {
              key,
              className: "dx-atom",
              "data-dx-atom": "\uFFFC",
              "data-dx-inline": index,
              title: "Image data is not available in this document",
            },
            h(
              "span",
              {
                className: "dx-img",
                style: { color: "#9a9a9a", fontSize: "10pt" },
              },
              "[image]",
            ),
          );
        return h(
          "span",
          {
            key,
            className: "dx-atom",
            "data-dx-atom": "\uFFFC",
            "data-dx-inline": index,
          },
          h("img", {
            className: "dx-img",
            src: url,
            alt: inline.alt || "Image",
            title: inline.title || inline.alt || "Image",
            draggable: false,
            style: {
              width: inline.width ? `${pxFromPt(inline.width)}px` : undefined,
              height: inline.height
                ? `${pxFromPt(inline.height)}px`
                : undefined,
            },
          }),
        );
      }
      if (inline.type === "math")
        return h(
          "span",
          {
            key,
            className: "dx-run",
            "data-dx-inline": index,
            style: { fontStyle: "italic", fontFamily: "Cambria Math, serif" },
            title: "Equation (linear form)",
          },
          text,
        );
      if (inline.type === "field")
        return h(
          "span",
          {
            key,
            className: "dx-run",
            "data-dx-inline": index,
            style: { color: "#555" },
            title: `Field ${inline.instr || ""}`,
          },
          text,
        );
      if (inline.type === "noteRef")
        return h(
          "span",
          {
            key,
            className: "dx-atom",
            "data-dx-atom": "",
            "data-dx-inline": index,
            style: { verticalAlign: "super", fontSize: "0.7em" },
            title:
              inline.noteType === "endnote"
                ? "Endnote reference"
                : "Footnote reference",
          },
          "*",
        );
      if (inline.type === "unknown")
        return h(
          "span",
          { key, className: "dx-run", "data-dx-inline": index },
          text,
        );
      return null;
    }

    function ParagraphView(props) {
      const ctx = props.ctx;
      const paragraph = props.paragraph;
      const css = paragraphStyleToCss(
        resolveParagraphProps(ctx.document, paragraph),
      );
      const label = ctx.listLabels.get(paragraph.id);
      const attrs = {
        className: "dx-para",
        "data-block-id": paragraph.id,
        style: css,
        role: "textbox",
        "aria-multiline": "true",
        "aria-label": `Paragraph ${paragraph.id}`,
        suppressContentEditableWarning: true,
      };
      if (ctx.readOnly) {
        // `contentEditable` is the React prop; the browser DOM attribute is the
        // case-insensitive `contenteditable`, which is what the tests and the
        // accessibility tree look for.
        attrs.contentEditable = false;
        attrs["aria-readonly"] = "true";
      } else {
        attrs.contentEditable = true;
        attrs.spellCheck = true;
      }
      const children = (paragraph.inlines || []).map((inline, index) =>
        h(RunView, {
          key: index,
          inline,
          index,
          ctx: Object.assign({}, ctx, { paragraph }),
        }),
      );
      if (children.length === 0) children.push(h("br", { key: "br" }));
      return h(
        "div",
        { className: "dx-para-row" },
        label
          ? h(
              "span",
              {
                className: "dx-list-label",
                "aria-hidden": "true",
                style: {
                  paddingLeft: `${Math.round((label.level || 0) * 14)}px`,
                },
              },
              label.text,
            )
          : null,
        h("div", attrs, children),
      );
    }

    function TableView(props) {
      const ctx = props.ctx;
      const table = props.table;
      const grid = (table.grid || []).length ? table.grid : null;
      return h(
        "table",
        { className: "dx-table", "data-block-id": table.id },
        grid
          ? h(
              "colgroup",
              null,
              grid.map((width, index) =>
                h("col", {
                  key: `col-${index}`,
                  style: { width: `${pxFromPt(width)}px` },
                }),
              ),
            )
          : null,
        h(
          "tbody",
          null,
          (table.rows || []).map((row) =>
            h(
              "tr",
              { key: row.id },
              (row.cells || []).map((cell) => {
                const merged =
                  cell.props &&
                  (cell.props.vMerge === "continue" ||
                    cell.props.gridSpan === 0);
                if (merged) return null;
                const cellStyle = {};
                if (cell.props && cell.props.width && cell.props.width.value)
                  cellStyle.width = `${pxFromPt(cell.props.width.value)}px`;
                if (
                  cell.props &&
                  cell.props.background &&
                  cell.props.background.fill
                )
                  cellStyle.backgroundColor = cell.props.background.fill;
                if (cell.props && cell.props.vAlign)
                  cellStyle.verticalAlign = cell.props.vAlign;
                return h(
                  "td",
                  {
                    key: cell.id,
                    colSpan:
                      cell.props && cell.props.gridSpan > 1
                        ? cell.props.gridSpan
                        : undefined,
                    style: cellStyle,
                  },
                  (cell.blocks || []).map((block) =>
                    renderBlock(
                      block,
                      Object.assign({}, ctx, { key: block.id }),
                    ),
                  ),
                );
              }),
            ),
          ),
        ),
      );
    }

    function TocView(props) {
      const ctx = props.ctx;
      const toc = props.toc;
      const entries = (ctx.outline || []).filter(
        (entry) =>
          entry.level >= (toc.minLevel || 1) &&
          entry.level <= (toc.maxLevel || 3),
      );
      return h(
        "div",
        { className: "dx-toc", "data-block-id": toc.id },
        h(
          "div",
          { className: "dx-toc-title" },
          toc.caption || "Table of Contents",
        ),
        entries.map((entry, index) =>
          h(
            "div",
            {
              key: `${entry.blockId || index}`,
              className: "dx-toc-entry",
              style: { paddingLeft: `${Math.max(0, entry.level - 1) * 14}px` },
            },
            h("span", null, entry.text || "Untitled heading"),
            h(
              "span",
              null,
              pageNumberForBlock(ctx.blockPages, entry.blockId) > 0
                ? String(pageNumberForBlock(ctx.blockPages, entry.blockId))
                : "",
            ),
          ),
        ),
        entries.length === 0
          ? h(
              "div",
              { className: "dx-toc-note" },
              "No headings with this outline level exist in the document.",
            )
          : null,
        h(
          "div",
          { className: "dx-toc-note" },
          "Generated from the document outline; update it from the References tab.",
        ),
      );
    }

    function renderBlock(block, ctx) {
      if (!block) return null;
      // Including the refetch epoch in the key remounts the paragraph after a
      // reload. The DOM is edited directly by typing, while React's last
      // committed text props still hold the pre-edit string; without the key
      // change an undo to that same string is a no-op for React and the stale
      // typed text stays on screen.
      const key = block.id ? `${block.id}@${ctx.epoch || 0}` : ctx.key;
      if (block.type === "paragraph")
        return h(ParagraphView, { key, ctx, paragraph: block });
      if (block.type === "table")
        return h(TableView, { key, ctx, table: block });
      if (block.type === "toc") return h(TocView, { key, ctx, toc: block });
      return null;
    }

    function EditorSurface(props) {
      const ctx = props.ctx;
      const api = props.api;
      const document = ctx.document;
      const pages = groupBlocksIntoPages(
        documentBodyBlocks(document),
        ctx.blockPages,
      );
      const sectionIndexByBlock = new Map();
      (document.sections || []).forEach((section, sectionIndex) => {
        for (const block of section.blocks || [])
          if (block && block.id)
            sectionIndexByBlock.set(block.id, sectionIndex);
      });
      const surfaceCtx = Object.assign({}, ctx, {
        sectionIndexByBlock,
      });
      const stackStyle = { "--dx-zoom": String(ctx.zoom || 1) };
      const surfaceAttrs = {
        className: ctx.showMarks ? "dx-editor dx-marks" : "dx-editor",
        ref: api.surfaceRef,
        role: "document",
        "aria-label": "Document editing surface",
        onInput: api.onInput,
        onPaste: api.onPaste,
        onBlur: api.onBlur,
        onKeyUp: api.onSelectionChange,
        onMouseUp: api.onSelectionChange,
        onMouseDown: api.onSurfaceMouseDown,
        onDoubleClick: api.onSurfaceDoubleClick,
        onCopy: api.onSurfaceCopy,
        onCut: api.onSurfaceCut,
        onBeforeInput: api.onBeforeInput,
        onCompositionStart: api.onCompositionStart,
        onCompositionEnd: api.onCompositionEnd,
        onFocus: api.onSelectionChange,
      };
      return h(
        "div",
        surfaceAttrs,
        ctx.showRuler
          ? h(Ruler, {
              section:
                (ctx.sections || [])[ctx.activeSection || 0] || ctx.sections[0],
            })
          : null,
        h(
          "div",
          { className: "dx-page-stack", style: stackStyle },
          pages.map((page) => {
            const sectionIndex =
              sectionIndexByBlock.get(page.blocks[0] && page.blocks[0].id) || 0;
            const section = (document.sections || [])[sectionIndex] || null;
            const headerText = headerFooterText(section, "header");
            const footerText = headerFooterText(section, "footer");
            const pageStyle = {};
            if (section && section.page) {
              pageStyle.width = `${pxFromPt(section.page.width)}px`;
              pageStyle.minWidth = `${Math.min(320, pxFromPt(section.page.width))}px`;
              pageStyle.minHeight = `${pxFromPt(
                Math.max(
                  240,
                  section.page.height -
                    section.page.margins.top -
                    section.page.margins.bottom,
                ),
              )}px`;
            }
            return h(
              "div",
              { key: `page-${page.page}`, className: "dx-page-wrap" },
              ctx.showPageSeparators && page.page > 1
                ? h("div", { className: "dx-page-sep" }, `Page ${page.page}`)
                : null,
              h(
                "div",
                {
                  className: "dx-page",
                  "data-page": String(page.page),
                  "aria-label": `Page ${page.page}`,
                  style: pageStyle,
                },
                h(
                  "div",
                  {
                    className: "dx-page-head",
                    hidden: !headerText,
                  },
                  headerText,
                ),
                page.blocks.map((block) =>
                  renderBlock(
                    block,
                    Object.assign({}, surfaceCtx, { key: block.id }),
                  ),
                ),
                h(
                  "div",
                  {
                    className: "dx-page-foot",
                    hidden: !footerText,
                  },
                  footerText,
                ),
              ),
            );
          }),
        ),
      );
    }

    const MemoSurface = React.memo(EditorSurface, (previous, next) => {
      return (
        previous.ctx.document === next.ctx.document &&
        previous.ctx.epoch === next.ctx.epoch &&
        previous.ctx.blockPages === next.ctx.blockPages &&
        previous.ctx.readOnly === next.ctx.readOnly &&
        previous.ctx.showMarks === next.ctx.showMarks &&
        previous.ctx.showRuler === next.ctx.showRuler &&
        previous.ctx.showPageSeparators === next.ctx.showPageSeparators &&
        previous.ctx.zoom === next.ctx.zoom &&
        previous.ctx.activeSection === next.ctx.activeSection &&
        previous.api === next.api
      );
    });

    function StatusBar(props) {
      const ctx = props;
      const api = props.api;
      const stats = ctx.stats || { words: 0, characters: 0 };
      const saveLabel =
        ctx.saveState === "saving"
          ? "Saving…"
          : ctx.saveState === "error"
            ? "Save failed"
            : ctx.dirty
              ? "Unsaved changes"
              : "Saved";
      const language =
        (ctx.document &&
          ctx.document.settings &&
          ctx.document.settings.defaultLanguage) ||
        (ctx.document &&
          ctx.document.styles &&
          ctx.document.styles.defaults &&
          ctx.document.styles.defaults.run &&
          ctx.document.styles.defaults.run.language) ||
        "en-US";
      return h(
        "footer",
        { className: "dx-statusbar", role: "status" },
        h(
          "span",
          { className: "dx-status-item" },
          ctx.bootstrap
            ? `Page ${ctx.currentPage || 1} of ${ctx.pageCount || 1}`
            : "No document",
        ),
        h("span", { className: "dx-status-item" }, `${stats.words || 0} words`),
        h(
          "span",
          { className: "dx-status-item" },
          `${stats.characters || 0} characters`,
        ),
        h("span", { className: "dx-status-item" }, language),
        h("span", { className: "dx-status-spacer" }),
        h(
          "span",
          {
            className: "dx-status-chip",
            title: ctx.readOnly
              ? REASONS.readOnly
              : "The document is editable.",
          },
          ctx.readOnly ? "Read-only" : "Editing",
        ),
        h("span", { className: "dx-status-item" }, saveLabel),
        h(
          "span",
          { className: "dx-status-zoom" },
          h(Icon, { name: "zoomOut", size: 12 }),
          h("input", {
            type: "range",
            min: 50,
            max: 200,
            step: 10,
            value: Math.round((ctx.zoom || 1) * 100),
            "aria-label": "Zoom slider",
            onChange: (event) => api.onZoom(Number(event.target.value) / 100),
          }),
          h("span", null, formatZoom(ctx.zoom || 1)),
        ),
      );
    }

    function ErrorBar(props) {
      const error = props.error;
      const api = props.api;
      const code = (error && error.code) || "ERROR";
      const message = (error && error.message) || "The operation failed.";
      return h(
        "div",
        { className: "dx-error", role: "alert" },
        h(Icon, { name: "info", size: 15 }),
        h(
          "div",
          { className: "dx-error-body" },
          h("code", null, code),
          " — ",
          message,
        ),
        h(
          "button",
          {
            type: "button",
            className: "dx-error-dismiss",
            "aria-label": "Dismiss error",
            title: "Dismiss",
            onClick: () => api.onAction("dismissError"),
          },
          h(Icon, { name: "close", size: 13 }),
        ),
      );
    }

    const FOCUSABLE_SELECTOR =
      "input:not([disabled]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), [href], [tabindex]:not([tabindex='-1'])";

    /** Keep Tab inside the dialog so keyboard focus never escapes the modal. */
    function trapDialogFocus(node, event) {
      if (!node || event.key !== "Tab") return;
      const focusables = Array.from(node.querySelectorAll(FOCUSABLE_SELECTOR));
      if (focusables.length === 0) return;
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      if (event.shiftKey && event.target === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && event.target === last) {
        event.preventDefault();
        first.focus();
      }
    }

    /**
     * Panel-header chip for the workspace/chat prepared by the last open or
     * folder choice. It never navigates on its own: "Switch to chat" is the
     * only navigation and it is user-initiated.
     */
    function SessionChip(props) {
      const chip = props.chip;
      if (!chip) return null;
      const status = chip.status || "workspace-only";
      return h(
        "div",
        {
          className: "dx-sessionbar",
          role: "status",
          "aria-live": "polite",
          "data-status": status,
        },
        h(Icon, {
          name: status === "chat-ready" ? "chat" : "info",
          size: 14,
          className: "dx-session-icon",
        }),
        h("span", { className: "dx-session-text" }, sessionChipText(chip)),
        chip.warning
          ? h("span", { className: "dx-session-warn" }, chip.warning)
          : null,
        h("span", { className: "dx-session-spacer" }),
        status === "chat-ready"
          ? h(
              "button",
              {
                type: "button",
                className: "dx-btn dx-primary",
                title: `Open the chat session for ${chip.folderName || chip.path}`,
                "aria-label": "Switch to chat",
                onClick: () => props.onSwitch && props.onSwitch(chip),
              },
              "Switch to chat",
            )
          : null,
        h(
          "button",
          {
            type: "button",
            className: "dx-btn",
            title: "Dismiss this notice",
            "aria-label": "Dismiss workspace notice",
            onClick: () => props.onDismiss && props.onDismiss(),
          },
          h(Icon, { name: "close", size: 12 }),
        ),
      );
    }

    function DialogShell(props) {
      const api = props.api;
      const nodeRef = React.useRef(null);
      React.useEffect(() => {
        const node = nodeRef.current;
        if (!node) return undefined;
        const focusable = node.querySelector(FOCUSABLE_SELECTOR);
        if (focusable && typeof focusable.focus === "function")
          focusable.focus();
        return undefined;
      }, [props.dialog && props.dialog.kind]);
      return h(
        "div",
        {
          className: "dx-overlay",
          onKeyDown: (event) => {
            if (event.key === "Escape") {
              event.preventDefault();
              event.stopPropagation();
              api.onDialogClose();
              return;
            }
            trapDialogFocus(nodeRef.current, event);
          },
        },
        h(
          "div",
          {
            className: props.wide ? "dx-dialog dx-dialog-wide" : "dx-dialog",
            role: "dialog",
            "aria-modal": "true",
            "aria-label": props.title,
            // DSH's shortcut dispatcher reads this when deciding which modal
            // owns a key; it also lets the panel's own commands be allowed in
            // this dialog through the `shortcuts` service.
            "data-shortcut-modal": "dx-dialog",
            ref: nodeRef,
          },
          h("h3", null, props.title),
          props.description ? h("p", null, props.description) : null,
          props.children,
          h(
            "div",
            { className: "dx-dialog-actions" },
            h(
              "button",
              {
                type: "button",
                className: "dx-btn",
                onClick: () => api.onDialogClose(),
              },
              "Cancel",
            ),
            props.onSubmit
              ? h(
                  "button",
                  {
                    type: "button",
                    className: "dx-btn dx-primary",
                    disabled: Boolean(props.submitDisabled),
                    title: props.submitTitle || props.submitLabel,
                    onClick: () => api.onDialogSubmit(),
                  },
                  props.submitLabel || "Insert",
                )
              : null,
          ),
        ),
      );
    }

    function DialogField(props) {
      return h(
        "label",
        { className: "dx-field" },
        props.label,
        h("input", {
          className: "dx-input",
          type: props.type || "text",
          value:
            props.value === undefined || props.value === null
              ? ""
              : String(props.value),
          placeholder: props.placeholder,
          disabled: props.disabled,
          min: props.min,
          max: props.max,
          "aria-label": props.label,
          onChange: (event) => props.onChange(event.target.value),
        }),
      );
    }

    function DialogTextArea(props) {
      return h(
        "label",
        { className: "dx-field" },
        props.label,
        h("textarea", {
          className: "dx-input",
          rows: props.rows || 3,
          value:
            props.value === undefined || props.value === null
              ? ""
              : String(props.value),
          placeholder: props.placeholder,
          "aria-label": props.label,
          onChange: (event) => props.onChange(event.target.value),
        }),
      );
    }

    function colorSwatches(ctx, api, field, current, includeNone) {
      const children = ctx.themeColors.map((color) =>
        h(
          "button",
          {
            key: color,
            type: "button",
            className: "dx-swatch",
            style: { background: color },
            title: color,
            "aria-label": `Use ${color}`,
            "aria-pressed": current === color,
            onClick: () => api.updateDialog({ [field]: color }),
          },
          null,
        ),
      );
      if (includeNone)
        children.unshift(
          h(
            "button",
            {
              key: "none",
              type: "button",
              className: "dx-swatch dx-swatch-none",
              title: "No colour",
              "aria-label": "No colour",
              "aria-pressed": !current,
              onClick: () => api.updateDialog({ [field]: "" }),
            },
            "None",
          ),
        );
      return h("div", { className: "dx-swatches" }, children);
    }

    function browseEntryLabel(entry) {
      return entry.kind === "directory"
        ? `Open folder ${entry.name}`
        : `Open ${entry.name}`;
    }

    function browseEntryMeta(entry) {
      if (entry.kind === "directory") return "folder";
      if (!entry.importable) return "not importable";
      return formatBytes(entry.size) || "document";
    }

    /**
     * In-panel folder browser driven by `GET /browse?path=`. Directories
     * navigate, importable documents open through `POST /open-path`; the
     * toolbar offers Up, breadcrumbs, the Host's roots and New folder.
     */
    function BrowseDialog(props) {
      const api = props.api;
      const dialog = props.dialog || {};
      // The native directory picker is optional: the Host may not expose
      // `uiWorkspace.pickDirectory`, and the chooser still works without it.
      const canPickDirectory = props.canPickDirectory !== false;
      const entries = dialog.entries || [];
      const roots = dialog.roots || [];
      const crumbs = breadcrumbSegments(dialog.path || "");
      const newFolder = dialog.newFolder;
      const newFolderName =
        typeof newFolder === "string" ? newFolder.trim() : "";
      const list = dialog.loading
        ? h(
            "p",
            { className: "dx-browse-empty" },
            "Listing the folder through the Host…",
          )
        : entries.length === 0
          ? h(
              "p",
              { className: "dx-browse-empty" },
              "This folder has no directories or importable documents.",
            )
          : h(
              "div",
              {
                className: "dx-browse-list",
                role: "list",
                "aria-label": "Folder contents",
              },
              entries.map((entry) =>
                h(
                  "button",
                  {
                    key: entry.path,
                    type: "button",
                    className: "dx-browse-item",
                    role: "listitem",
                    "aria-label": browseEntryLabel(entry),
                    title:
                      entry.kind === "document" && !entry.importable
                        ? "The Host cannot import this file type."
                        : entry.path,
                    disabled: entry.kind === "document" && !entry.importable,
                    onClick: () =>
                      api.onAction(
                        entry.kind === "directory"
                          ? "browsePath"
                          : "openBrowseDocument",
                        entry.path,
                      ),
                  },
                  h(Icon, {
                    name: entry.kind === "directory" ? "folder" : "page",
                    size: 14,
                  }),
                  h("span", null, entry.name),
                  h(
                    "span",
                    { className: "dx-browse-meta" },
                    browseEntryMeta(entry),
                  ),
                ),
              ),
            );
      return h(DialogShell, {
        api,
        wide: true,
        title: "Open a document",
        description:
          "Pick a folder or a document the Host can import. Browsing is read-only; only names and sizes are listed.",
        submitLabel: "Open this folder",
        submitDisabled: !dialog.path,
        submitTitle:
          "Switch the document workspace to this folder without opening a file",
        onSubmit: () => api.onAction("openCurrentFolder"),
        children: h(
          "div",
          { className: "dx-browse" },
          roots.length
            ? h(
                "div",
                {
                  className: "dx-roots",
                  role: "group",
                  "aria-label": "Browse roots",
                },
                roots.map((root, index) =>
                  h(
                    "button",
                    {
                      key: `${root.path}-${index}`,
                      type: "button",
                      className: "dx-root-link",
                      title: `Browse ${root.path}`,
                      onClick: () => api.onAction("browseRoot", root.path),
                    },
                    root.label,
                  ),
                ),
              )
            : null,
          h(
            "div",
            { className: "dx-browse-bar" },
            h(
              "button",
              {
                type: "button",
                className: "dx-btn",
                title: dialog.parent
                  ? `Go up to ${dialog.parent}`
                  : "Already at the top of the browsable roots",
                "aria-label": "Go up one folder",
                disabled: !dialog.parent,
                onClick: () => api.onAction("browseUp"),
              },
              h(Icon, { name: "up", size: 13 }),
              "Up",
            ),
            h(
              "div",
              { className: "dx-breadcrumbs", "aria-label": "Folder path" },
              crumbs.map((crumb, index) =>
                h(
                  "span",
                  { key: crumb.path, style: { display: "inline-flex" } },
                  index > 0
                    ? h(
                        "span",
                        { className: "dx-crumb-sep", "aria-hidden": "true" },
                        "›",
                      )
                    : null,
                  h(
                    "button",
                    {
                      key: "crumb",
                      type: "button",
                      className: "dx-crumb",
                      title: crumb.path,
                      "aria-current":
                        index === crumbs.length - 1 ? "page" : undefined,
                      onClick: () => api.onAction("browsePath", crumb.path),
                    },
                    crumb.label,
                  ),
                ),
              ),
            ),
            h(
              "button",
              {
                type: "button",
                className: "dx-btn",
                title: "Create a new folder here and browse into it",
                "aria-label": "New folder",
                onClick: () => api.onAction("newFolderStart"),
              },
              h(Icon, { name: "folderPlus", size: 13 }),
              "New folder",
            ),
            h(
              "button",
              {
                type: "button",
                className: "dx-btn",
                title: canPickDirectory
                  ? "Use the system folder picker, then list that folder here"
                  : "This Host exposes no native folder picker; browse with Up and the folders above instead.",
                "aria-label": "Choose folder with the system dialog",
                disabled: !canPickDirectory,
                onClick: () => api.onAction("chooseFolder"),
              },
              h(Icon, { name: "folder", size: 13 }),
              "Choose folder…",
            ),
          ),
          typeof newFolder === "string"
            ? h(
                "div",
                { className: "dx-newfolder" },
                h("input", {
                  className: "dx-input",
                  type: "text",
                  value: newFolder,
                  placeholder: "New folder name",
                  "aria-label": "New folder name",
                  onChange: (event) =>
                    api.updateDialog({ newFolder: event.target.value }),
                }),
                h(
                  "button",
                  {
                    type: "button",
                    className: "dx-btn dx-primary",
                    disabled: !newFolderName,
                    title: "Create this folder in the Host workspace",
                    onClick: () => api.onAction("newFolderConfirm"),
                  },
                  "Create",
                ),
              )
            : null,
          dialog.error
            ? h(
                "p",
                { className: "dx-browse-error", role: "alert" },
                dialog.error,
              )
            : null,
          list,
          dialog.truncated
            ? h(
                "p",
                { className: "dx-file-meta" },
                "The Host truncated this listing; open a subfolder to narrow it.",
              )
            : null,
        ),
      });
    }

    function DialogContent(props) {
      const ctx = props;
      const api = props.api;
      const dialog = props.dialog || {};
      switch (dialog.kind) {
        case "new":
          return h(DialogShell, {
            api,
            title: "New document",
            description:
              "The document is created in the workspace and opened in the editor.",
            submitLabel: "Create",
            submitDisabled: !String(dialog.title || "").trim(),
            onSubmit: () => api.onAction("createDocument"),
            children: h(DialogField, {
              label: "Title",
              value: dialog.title || "",
              placeholder: "Quarterly report",
              onChange: (value) => api.updateDialog({ title: value }),
            }),
          });
        case "open":
          return h(DialogShell, {
            api,
            title: "Open a document by path",
            description:
              "The path must resolve inside the Host workspace; DOCX, DOC, ODT, RTF, Pages, HTML and Markdown are supported by the engine.",
            submitLabel: "Open",
            submitDisabled: !String(dialog.path || "").trim(),
            onSubmit: () => api.onAction("openDocumentByPath"),
            children: h(DialogField, {
              label: "Path",
              value: dialog.path || "",
              placeholder: "documents/report.docx",
              onChange: (value) => api.updateDialog({ path: value }),
            }),
          });
        case "browse":
          return h(BrowseDialog, {
            api,
            dialog,
            canPickDirectory: props.canPickDirectory !== false,
          });
        case "saveAs":
          return h(DialogShell, {
            api,
            title: "Save as",
            description:
              "The extension selects the format. The path must stay inside the workspace.",
            submitLabel: "Save",
            submitDisabled: !String(dialog.path || "").trim(),
            onSubmit: () => api.onAction("saveDocumentAs"),
            children: h(DialogField, {
              label: "Path",
              value: dialog.path || "",
              placeholder: "documents/report.docx",
              onChange: (value) => api.updateDialog({ path: value }),
            }),
          });
        case "table":
          return h(DialogShell, {
            api,
            title: "Insert table",
            description: "Rows and columns are validated by the Host.",
            submitLabel: "Insert",
            onSubmit: () => api.onAction("insertTableConfirm"),
            children: h(
              "div",
              { className: "dx-field-row" },
              h(DialogField, {
                label: "Rows",
                type: "number",
                min: 1,
                max: 200,
                value: dialog.rows === undefined ? 3 : dialog.rows,
                onChange: (value) => api.updateDialog({ rows: value }),
              }),
              h(DialogField, {
                label: "Columns",
                type: "number",
                min: 1,
                max: 40,
                value: dialog.columns === undefined ? 3 : dialog.columns,
                onChange: (value) => api.updateDialog({ columns: value }),
              }),
            ),
          });
        case "link":
          return h(DialogShell, {
            api,
            title: "Insert hyperlink",
            description:
              "Only http, https and mailto destinations are accepted; anything else is rejected before it reaches the model.",
            submitLabel: "Insert",
            submitDisabled: !sanitizeUrl(dialog.href || ""),
            submitTitle: sanitizeUrl(dialog.href || "")
              ? "Insert the link"
              : "Enter a valid http, https or mailto URL",
            onSubmit: () => api.onAction("insertLinkConfirm"),
            children: h(
              "div",
              null,
              h(DialogField, {
                label: "Address",
                value: dialog.href || "",
                placeholder: "https://example.com",
                onChange: (value) => api.updateDialog({ href: value }),
              }),
              h(DialogField, {
                label: "Tooltip (optional)",
                value: dialog.tooltip || "",
                onChange: (value) => api.updateDialog({ tooltip: value }),
              }),
            ),
          });
        case "image":
          return h(DialogShell, {
            api,
            title: "Insert picture",
            description:
              "The image is read locally, sent to the Host as base64 and stored in the document model. Nothing is uploaded to a network.",
            submitLabel: "Insert",
            submitDisabled: !dialog.dataBase64,
            onSubmit: () => api.onAction("insertImageConfirm"),
            children: h(
              "label",
              { className: "dx-field" },
              "Image file",
              h("input", {
                className: "dx-input",
                type: "file",
                accept: "image/png,image/jpeg,image/gif,image/webp,image/bmp",
                "aria-label": "Image file",
                onChange: (event) => api.onImagePicked(event),
              }),
            ),
          });
        case "toc":
          return h(DialogShell, {
            api,
            title: "Insert table of contents",
            description:
              "The TOC is a generated block; update it from the References tab.",
            submitLabel: "Insert",
            onSubmit: () => api.onAction("insertTocConfirm"),
            children: h(
              "div",
              { className: "dx-field-row" },
              h(DialogField, {
                label: "Start level",
                type: "number",
                min: 1,
                max: 9,
                value: dialog.minLevel === undefined ? 1 : dialog.minLevel,
                onChange: (value) => api.updateDialog({ minLevel: value }),
              }),
              h(DialogField, {
                label: "End level",
                type: "number",
                min: 1,
                max: 9,
                value: dialog.maxLevel === undefined ? 3 : dialog.maxLevel,
                onChange: (value) => api.updateDialog({ maxLevel: value }),
              }),
            ),
          });
        case "note":
          return h(DialogShell, {
            api,
            title:
              dialog.noteType === "endnote"
                ? "Insert endnote"
                : "Insert footnote",
            description:
              "The note text is stored in the model; the reference marker is generated by the engine.",
            submitLabel: "Insert",
            submitDisabled: !String(dialog.text || "").trim(),
            onSubmit: () => api.onAction("insertNoteConfirm"),
            children: h(DialogTextArea, {
              label: "Note text",
              value: dialog.text || "",
              onChange: (value) => api.updateDialog({ text: value }),
            }),
          });
        case "equation":
          return h(DialogShell, {
            api,
            title: "Insert equation",
            description: "Linear form, e.g. E = mc^2.",
            submitLabel: "Insert",
            submitDisabled: !String(dialog.linear || "").trim(),
            onSubmit: () => api.onAction("insertEquationConfirm"),
            children: h(DialogField, {
              label: "Equation",
              value: dialog.linear || "",
              placeholder: "x = (-b ± √(b^2-4ac)) / 2a",
              onChange: (value) => api.updateDialog({ linear: value }),
            }),
          });
        case "headerFooter":
          return h(DialogShell, {
            api,
            title: dialog.part === "header" ? "Edit header" : "Edit footer",
            description:
              "Plain text only: the editor API exposes no page-number field for headers and footers.",
            submitLabel: "Apply",
            onSubmit: () => api.onAction("editHeaderFooterConfirm"),
            children: h(DialogTextArea, {
              label: `${dialog.part === "header" ? "Header" : "Footer"} text`,
              value: dialog.text || "",
              onChange: (value) => api.updateDialog({ text: value }),
            }),
          });
        case "margins":
          return h(DialogShell, {
            api,
            title: "Custom margins",
            description: "Values in points (72pt = 1 inch).",
            submitLabel: "Apply",
            onSubmit: () => api.onAction("marginsConfirm"),
            children: h(
              "div",
              { className: "dx-field-row" },
              h(DialogField, {
                label: "Top",
                type: "number",
                value: dialog.top,
                onChange: (value) => api.updateDialog({ top: value }),
              }),
              h(DialogField, {
                label: "Right",
                type: "number",
                value: dialog.right,
                onChange: (value) => api.updateDialog({ right: value }),
              }),
              h(DialogField, {
                label: "Bottom",
                type: "number",
                value: dialog.bottom,
                onChange: (value) => api.updateDialog({ bottom: value }),
              }),
              h(DialogField, {
                label: "Left",
                type: "number",
                value: dialog.left,
                onChange: (value) => api.updateDialog({ left: value }),
              }),
            ),
          });
        case "color":
          return h(DialogShell, {
            api,
            title:
              dialog.field === "highlight" ? "Highlight colour" : "Text colour",
            description: "Applied to the current selection with formatText.",
            submitLabel: "Apply",
            onSubmit: () => api.onAction("applyColorConfirm"),
            children: h(
              "div",
              null,
              colorSwatches(
                ctx,
                api,
                "value",
                dialog.value,
                dialog.field === "highlight",
              ),
            ),
          });
        case "find":
          return h(DialogShell, {
            api,
            title: "Find and replace",
            description:
              "Search runs on the Host; replace posts one replaceText op per match you apply.",
            submitLabel: "Find",
            onSubmit: () => api.onAction("runSearch"),
            children: h(
              "div",
              null,
              h(DialogField, {
                label: "Find what",
                value: dialog.query || "",
                onChange: (value) => api.updateDialog({ query: value }),
              }),
              h(DialogField, {
                label: "Replace with",
                value: dialog.replacement || "",
                onChange: (value) => api.updateDialog({ replacement: value }),
              }),
              h(
                "div",
                { className: "dx-find-list" },
                (dialog.matches || []).length === 0
                  ? h(
                      "p",
                      { className: "dx-file-meta" },
                      dialog.searched
                        ? "No matches."
                        : "Run a search to list matches.",
                    )
                  : dialog.matches.map((match, index) =>
                      h(
                        "button",
                        {
                          key: `${match.blockId}-${match.start}-${index}`,
                          type: "button",
                          className: "dx-find-item",
                          "aria-current":
                            dialog.activeMatch === index ? true : undefined,
                          title: `Page ${match.page || 1}`,
                          onClick: () => api.onAction("selectFindMatch", index),
                        },
                        `${match.text} — page ${match.page || 1}`,
                      ),
                    ),
              ),
              h(
                "button",
                {
                  type: "button",
                  className: "dx-btn",
                  disabled: !dialog.activeMatch && dialog.activeMatch !== 0,
                  title: "Replace the selected match",
                  onClick: () => api.onAction("replaceMatch"),
                },
                "Replace",
              ),
            ),
          });
        case "delete":
          return h(DialogShell, {
            api,
            title: "Delete document",
            description:
              "The document is removed from the workspace. This cannot be undone from the panel.",
            submitLabel: "Delete",
            onSubmit: () => api.onAction("confirmDelete"),
            children: h("p", null, `Delete “${ctx.title || "this document"}”?`),
          });
        default:
          return h(DialogShell, {
            api,
            title: "Unsupported dialog",
            description: "This dialog has no implementation.",
            children: h(
              "p",
              null,
              "Close this dialog and use another command.",
            ),
          });
      }
    }

    function EmptyState(props) {
      const api = props.api;
      const ctx = props;
      return h(
        "div",
        {
          className: "dx-editor",
          style: {
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          },
        },
        h(
          "div",
          { style: { maxWidth: 420, textAlign: "center" } },
          h(Icon, { name: "newDoc", size: 42 }),
          h(
            "h2",
            { style: { margin: "10px 0 6px", fontSize: 16 } },
            "No document open",
          ),
          h(
            "p",
            { className: "dx-file-meta" },
            "Create a document, or open one with the in-panel file chooser. The document list lives in the rail on the left.",
          ),
          h(
            "div",
            {
              style: {
                display: "flex",
                gap: 8,
                justifyContent: "center",
                marginTop: 10,
              },
            },
            h(
              "button",
              {
                type: "button",
                className: "dx-btn dx-primary",
                disabled: Boolean(ctx.readOnly),
                title: ctx.readOnly ? REASONS.readOnly : "Create a document",
                onClick: () => api.onAction("newDocument"),
              },
              "New document",
            ),
            h(
              "button",
              {
                type: "button",
                className: "dx-btn",
                title:
                  "Browse folders and open a document without the OS file dialog (Ctrl+O)",
                onClick: () => api.onAction("openFileChooser"),
              },
              "Open file…",
            ),
            h(
              "button",
              {
                type: "button",
                className: "dx-btn",
                disabled: Boolean(ctx.readOnly),
                title: "Open a file by path",
                onClick: () => api.onAction("openByPath"),
              },
              "Open by path",
            ),
          ),
        ),
      );
    }

    function FallbackPreview(props) {
      const api = props.api;
      const fallback = props.fallback || {};
      const outline = fallback.outline || [];
      return h(
        "div",
        { className: "dx-editor dx-preview" },
        h(
          "div",
          { className: "dx-preview-sheet" },
          h(
            "h2",
            { className: "dx-preview-title" },
            fallback.title || "Document",
          ),
          h(
            "p",
            { className: "dx-notice", style: { margin: "0 0 10px" } },
            "The Host is running an older Documents API without /bootstrap and /op, so this panel cannot edit this document in place. The text preview and the export flows still work; update the Host half to the contract in EDITOR-API.md to enable editing.",
          ),
          outline.length
            ? h(
                "div",
                { className: "dx-file-meta" },
                h("b", null, "Outline"),
                h(
                  "ul",
                  { className: "dx-preview-outline" },
                  outline.slice(0, 40).map((entry, index) =>
                    h(
                      "li",
                      {
                        key: `${entry.blockId || index}-${index}`,
                        style: {
                          paddingLeft: `${Math.max(0, (entry.level || 1) - 1) * 14}px`,
                        },
                      },
                      entry.text,
                    ),
                  ),
                ),
              )
            : null,
          h(
            "pre",
            { className: "dx-pre" },
            fallback.text || "(the Host returned no text)",
          ),
          fallback.truncated
            ? h(
                "p",
                { className: "dx-file-meta" },
                "The text preview is truncated.",
              )
            : null,
          h(
            "div",
            { style: { display: "flex", gap: 8, marginTop: 10 } },
            ["docx", "pdf", "html", "md"].map((format) =>
              h(
                "button",
                {
                  key: format,
                  type: "button",
                  className: "dx-btn",
                  title: `Export ${format.toUpperCase()}`,
                  onClick: () => api.onAction("export", format),
                },
                format.toUpperCase(),
              ),
            ),
            h(
              "button",
              {
                type: "button",
                className: "dx-btn",
                title: "Refresh the document list and the preview",
                onClick: () => api.onAction("refresh"),
              },
              "Refresh",
            ),
          ),
        ),
      );
    }

    function DocumentsPanelView(props) {
      const ctx = props;
      const api = props.api;
      const surfaceCtx = {
        document: ctx.document || { sections: [] },
        blockPages: ctx.bootstrap ? ctx.bootstrap.blockPages : null,
        readOnly: ctx.readOnly,
        showMarks: ctx.showMarks,
        showRuler: ctx.showRuler,
        showPageSeparators: ctx.showPageSeparators,
        zoom: ctx.zoom,
        activeSection: ctx.activeSection,
        sections: ctx.sections,
        media: ctx.media,
        listLabels: ctx.listLabels,
        outline: ctx.outline,
        api,
        epoch: ctx.surfaceEpoch || 0,
      };
      return h(
        "div",
        {
          className: ctx.showLabels ? "dx-root dx-show-labels" : "dx-root",
          ref: api.rootRef,
          onKeyDown: api.onRootKeyDown,
        },
        styles(),
        h(TitleBar, props),
        h(TabRow, props),
        ctx.sessionChip
          ? h(SessionChip, {
              chip: ctx.sessionChip,
              onSwitch: () => api.onAction("switchToChat"),
              onDismiss: () => api.onAction("dismissSessionChip"),
            })
          : null,
        h(Ribbon, props),
        ctx.error ? h(ErrorBar, props) : null,
        ctx.notice
          ? h(
              "div",
              { className: "dx-notice" },
              ctx.notice,
              ctx.noticeDismissible
                ? h(
                    "button",
                    {
                      type: "button",
                      className: "dx-notice-dismiss",
                      "aria-label": "Dismiss notice",
                      title: "Dismiss",
                      onClick: () => api.onAction("dismissError"),
                    },
                    "×",
                  )
                : null,
            )
          : null,
        h(
          "div",
          { className: "dx-work" },
          ctx.showRail ? h(DocumentRail, props) : null,
          ctx.bootstrap
            ? h(MemoSurface, { ctx: surfaceCtx, api })
            : ctx.fallback
              ? h(FallbackPreview, props)
              : h(EmptyState, props),
        ),
        h(StatusBar, props),
        ctx.dialog ? h(DialogContent, props) : null,
      );
    }

    /* --------------------------------------------------------- container */

    function noop() {
      return undefined;
    }

    function makeNoopApi() {
      return {
        onAction: noop,
        onTitleChange: noop,
        onTitleCommit: noop,
        onTab: noop,
        onToggleRibbon: noop,
        onZoom: noop,
        updateDialog: noop,
        onDialogClose: noop,
        onDialogSubmit: noop,
        onImagePicked: noop,
        onFilePicked: noop,
        onAdvancedPath: noop,
        onToggleAdvanced: noop,
        onOpenDocument: noop,
        onOpenUrl: noop,
        onSelectionChange: noop,
        onInput: noop,
        onPaste: noop,
        onSurfaceMouseDown: noop,
        onSurfaceDoubleClick: noop,
        onSurfaceCopy: noop,
        onSurfaceCut: noop,
        onBeforeInput: noop,
        onCompositionStart: noop,
        onCompositionEnd: noop,
        onBlur: noop,
        onRootKeyDown: noop,
        onRibbonHeight: noop,
        rootRef: null,
        surfaceRef: null,
        fileInputRef: null,
        ribbonRef: null,
      };
    }

    function renderForTest(model, options) {
      const state = seedState(model, options || {});
      return h(DocumentsPanelView, buildViewProps(state, makeNoopApi()));
    }

    function DocumentsPanel() {
      const [state, setState] = React.useState(initialState);
      const stateRef = React.useRef(state);
      stateRef.current = state;

      const rootRef = React.useRef(null);
      const surfaceRef = React.useRef(null);
      const fileInputRef = React.useRef(null);
      const ribbonRef = React.useRef(null);
      const pendingTextRef = React.useRef(new Map());
      const hostTextRef = React.useRef(new Map());
      const flushTimerRef = React.useRef(null);
      const titleTimerRef = React.useRef(null);
      const selectionRef = React.useRef(null);
      const caretRef = React.useRef(null);
      const dialogReturnFocusRef = React.useRef(null);
      const offlineRef = React.useRef(false);
      // Pointer drag that started inside one paragraph and left it: a DOM
      // Range can span two contenteditable islands, so the panel builds it.
      const dragSelectRef = React.useRef(null);
      // IME composition must not be diffed mid-flight: the browser owns the
      // text until compositionend, otherwise the model and the composition
      // buffer each apply the same characters.
      const composingRef = React.useRef(false);
      // One flush at a time; concurrent flushes used to interleave snapshots
      // and re-send text that was already applied.
      const flushPromiseRef = React.useRef(null);
      // Fixed end of a Shift+Arrow selection.
      const selectionAnchorRef = React.useRef(null);
      // A programmatic caret/selection placement owns the selection until its
      // DOM write lands; a keyup/mouseup in that window must not publish the
      // previous (now stale) DOM selection.
      const pendingPlacementRef = React.useRef(false);
      const placementResolveRef = React.useRef(null);
      const placementPromiseRef = React.useRef(null);
      const placeCaretRef = React.useRef(null);
      const placementSeqRef = React.useRef(0);
      // Fallback path for the React handler and the capture listener.
      const execShortcutRef = React.useRef(null);
      const mountedRef = React.useRef(false);

      const patch = React.useCallback((next) => {
        setState((current) =>
          Object.assign(
            {},
            current,
            typeof next === "function" ? next(current) : next,
          ),
        );
      }, []);

      const isReadOnly = React.useCallback((current) => {
        if (!current) return true;
        if (current.readOnlyOverride !== null) return current.readOnlyOverride;
        return Boolean(
          (current.bootstrap && current.bootstrap.readOnly) ||
          current.workspaceReadOnly,
        );
      }, []);

      /**
       * Resize the ribbon. `persist` also writes the height to `localStorage`;
       * a `null` value removes the key, which restores the 1/10 default.
       */
      const onRibbonHeight = React.useCallback(
        (value, persist) => {
          const next =
            value === null || value === undefined
              ? null
              : clampRibbonHeight(value, viewportHeight());
          patch({ ribbonHeight: next });
          if (persist === true) storeRibbonHeight(next);
        },
        [patch],
      );

      const reportError = React.useCallback(
        (cause) => {
          const code = (cause && cause.code) || "REQUEST_FAILED";
          const message =
            (cause && cause.message) || String(cause || "Operation failed");
          const next = { error: { code, message } };
          if (isMissingRoute(cause)) {
            offlineRef.current = true;
            next.offline = true;
            next.notice =
              `The Host editor API is unavailable (${code}). ` +
              "The document stays editable locally, but commands that need the Host are disabled. " +
              "Reload this Harness window, and restart the app once if the editor routes are still missing.";
          }
          patch(next);
          return { code, message };
        },
        [patch],
      );

      const applyEnvelope = React.useCallback(
        (result) => {
          if (!result) return;
          patch((current) => {
            const bootstrap = current.bootstrap;
            if (!bootstrap) return {};
            const dirty =
              result.dirty === undefined ? true : result.dirty === true;
            return {
              bootstrap: Object.assign({}, bootstrap, {
                revision:
                  typeof result.revision === "number"
                    ? result.revision
                    : bootstrap.revision,
                dirty,
                title: result.title || bootstrap.title,
                stats: Object.assign({}, bootstrap.stats, result.stats || {}),
                pageCount: result.pageCount || bootstrap.pageCount,
              }),
              saveState: dirty ? "unsaved" : "saved",
              canUndo: true,
              canRedo: true,
            };
          });
        },
        [patch],
      );

      /**
       * Rebuild the "last text the Host confirmed" cache after a refetch.
       *
       * Typed-but-unflushed text is local state the Host has never seen, so a
       * refetch must not silently drop it (loss) or let it replay on top of the
       * server copy (duplication). Instead the pending text is re-applied to the
       * fresh model — a no-op when the Host already has it — and stays in the
       * pending map until one flush reconciles it against the new cache.
       */
      const rebuildTextCache = React.useCallback((document, options) => {
        const keepPending = Boolean(options && options.keepPending);
        if (!keepPending) pendingTextRef.current.clear();
        const cache = new Map();
        const { list, byId } = collectBlocks(document);
        for (const entry of list) {
          if (entry.block && entry.block.type === "paragraph")
            cache.set(entry.block.id, paragraphText(entry.block));
        }
        hostTextRef.current = cache;
        if (!keepPending) return;
        for (const [blockId, text] of Array.from(pendingTextRef.current)) {
          const entry = byId.get(blockId);
          if (!entry || !entry.block || entry.block.type !== "paragraph") {
            pendingTextRef.current.delete(blockId);
            continue;
          }
          const serverText = paragraphText(entry.block);
          if (serverText !== text) {
            const diff = diffTextRange(serverText, text);
            if (diff)
              applyTextReplace(entry.block, diff.start, diff.end, diff.insert);
          }
        }
      }, []);

      const refreshList = React.useCallback(async () => {
        try {
          const payload = await apiClient.list();
          patch({
            documents: payload.documents || [],
            workspace: payload.workspace || "",
            workspaceReadOnly: payload.readOnly === true,
          });
          return payload;
        } catch (cause) {
          reportError(cause);
          return null;
        }
      }, [patch, reportError]);

      const reload = React.useCallback(
        async (caret, options) => {
          const current = stateRef.current;
          const documentId =
            (options && options.documentId) ||
            (current.bootstrap && current.bootstrap.documentId) ||
            current.selectedId;
          if (!documentId) {
            reportError(HostError("NO_DOCUMENT", REASONS.noDocument));
            return null;
          }
          try {
            const payload = await apiClient.bootstrap(documentId);
            const bootstrap = normalizeBootstrap(payload);
            bootstrap.documentId = documentId;
            const reserved =
              stateRef.current.boldReserved ||
              detectBoldReservation(clientService("shortcuts"));
            const showBoldNote = reserved && !boldReservedNoticeShown;
            if (showBoldNote) boldReservedNoticeShown = true;
            // Keep typed text that has not reached the Host yet: re-apply it to
            // the fresh model and let the next flush send exactly the delta.
            rebuildTextCache(bootstrap.document, { keepPending: true });
            // `null` means "the caller places the caret itself" (split,
            // paste, lists). Resurrecting the Host's stale selection here would
            // run after that placement and pull the caret back into the old
            // paragraph.
            const explicit = caret !== undefined && caret !== null;
            // `null` normally means "use the Host's selection" (initial load);
            // structural callers pass skipPlacement because they place the
            // caret themselves and must not have the stale one re-applied.
            const skipPlacement = Boolean(
              caret === null && options && options.skipPlacement,
            );
            const selection = explicit ? caret : bootstrap.selection;
            caretRef.current = explicit
              ? caret
              : skipPlacement
                ? null
                : bootstrap.selection;
            offlineRef.current = false;
            patch((previous) => ({
              bootstrap,
              selectedId: documentId,
              selection: explicit ? selection : previous.selection || selection,
              workspace: bootstrap.workspace || previous.workspace,
              documents: (previous.documents || []).map((item) =>
                item.id === documentId
                  ? Object.assign({}, item, {
                      title: bootstrap.title,
                      dirty: bootstrap.dirty,
                      words: bootstrap.stats.words,
                    })
                  : item,
              ),
              saveState: bootstrap.dirty ? "unsaved" : "saved",
              offline: false,
              boldReserved: previous.boldReserved || reserved,
              ...(showBoldNote
                ? {
                    notice: BOLD_RESERVED_NOTE,
                    noticeDismissible: true,
                  }
                : {}),
              // The surface is memoised on document identity; a refetch that
              // returns structurally equal content still needs a render when
              // the DOM was edited optimistically (undo/redo, refetch after a
              // failed op). An explicit epoch makes that render unconditional.
              surfaceEpoch: (previous.surfaceEpoch || 0) + 1,
            }));
            return bootstrap;
          } catch (cause) {
            reportError(cause);
            return null;
          }
        },
        [patch, rebuildTextCache, reportError],
      );

      /**
       * Push typed text to the Host.
       *
       * Single-flight: concurrent callers share one promise, so two flushes can
       * never interleave snapshots of `pendingTextRef` and both send the same
       * characters. Each pass takes the current entries, removes exactly those
       * (a keystroke arriving mid-flight stays queued), and sends at most one
       * `replaceText` op per block — only when the text differs from what the
       * Host last confirmed.
       *
       * A failed batch is not blindly retried: the Host may have applied it and
       * lost only the response. The panel refetches the authoritative text,
       * rebuilds the cache, and re-diffs the local text, so a retry sends
       * nothing when the first attempt actually landed.
       */
      const flushPending = React.useCallback(async () => {
        if (offlineRef.current) return;
        const current = stateRef.current;
        const bootstrap = current.bootstrap;
        if (!bootstrap) return;
        if (flushPromiseRef.current) return flushPromiseRef.current;
        const documentId = bootstrap.documentId;
        const run = (async () => {
          let attempt = 0;
          while (attempt < 3) {
            const pending = pendingTextRef.current;
            if (pending.size === 0) return;
            const entries = Array.from(pending.entries());
            for (const [blockId, text] of entries) {
              if (pending.get(blockId) === text) pending.delete(blockId);
            }
            const plan = planTextFlush(hostTextRef.current, entries);
            for (const item of plan.skipped) {
              if (item.reason === "unknown-block")
                pendingTextRef.current.delete(item.blockId);
            }
            if (plan.ops.length === 0) continue;
            // Invariant: the DOM paragraph must still hold exactly the text the
            // plan is about to send. If it does not (a keystroke landed while
            // the model was being rebuilt), the plan is stale: keep the DOM
            // text pending and re-read the Host document instead of posting a
            // range against a document that moved.
            let domMismatch = null;
            const surface = surfaceRef.current;
            for (const op of plan.ops) {
              const element = surface
                ? findBlockElement(surface, op.blockId)
                : null;
              const domText = element ? domTextOfBlock(element) : null;
              if (domText !== null && domText !== op.text) {
                domMismatch = { blockId: op.blockId, text: domText };
                break;
              }
            }
            if (domMismatch) {
              for (const [blockId, text] of entries)
                pendingTextRef.current.set(blockId, text);
              pendingTextRef.current.set(domMismatch.blockId, domMismatch.text);
              attempt += 1;
              if (attempt > 2) return;
              const reference = stateRef.current.bootstrap;
              if (!reference) return;
              try {
                const payload = await apiClient.bootstrap(reference.documentId);
                const next = normalizeBootstrap(payload);
                next.documentId = reference.documentId;
                rebuildTextCache(next.document, { keepPending: true });
                patch((previous) => ({
                  bootstrap: previous.bootstrap
                    ? Object.assign({}, previous.bootstrap, {
                        document: next.document,
                        stats: next.stats,
                        blockPages: next.blockPages,
                        pageCount: next.pageCount,
                        outline: next.outline,
                        revision: next.revision,
                      })
                    : previous.bootstrap,
                }));
              } catch (_error) {
                return;
              }
              continue;
            }
            const ops = plan.ops.map((op) => ({
              op: "replaceText",
              args: {
                range: {
                  start: { blockId: op.blockId, offset: op.start },
                  end: { blockId: op.blockId, offset: op.end },
                },
                text: op.insert,
              },
            }));
            try {
              const result = await apiClient.ops(documentId, ops);
              applyEnvelope(result);
              for (const op of plan.ops)
                hostTextRef.current.set(op.blockId, op.text);
              attempt = 0;
            } catch (cause) {
              for (const op of plan.ops) {
                if (!pendingTextRef.current.has(op.blockId))
                  pendingTextRef.current.set(op.blockId, op.text);
              }
              reportError(cause);
              attempt += 1;
              // Re-read the authoritative text; keep pending so the next pass
              // sends only what is still missing.
              const reference = stateRef.current.bootstrap;
              if (!reference) return;
              try {
                const payload = await apiClient.bootstrap(documentId);
                const next = normalizeBootstrap(payload);
                next.documentId = documentId;
                rebuildTextCache(next.document, { keepPending: true });
                patch((previous) => ({
                  bootstrap: previous.bootstrap
                    ? Object.assign({}, previous.bootstrap, {
                        document: next.document,
                        stats: next.stats,
                        blockPages: next.blockPages,
                        pageCount: next.pageCount,
                        outline: next.outline,
                        revision: next.revision,
                      })
                    : previous.bootstrap,
                }));
              } catch (_error) {
                return;
              }
            }
          }
        })();
        flushPromiseRef.current = run
          .catch(() => undefined)
          .finally(() => {
            flushPromiseRef.current = null;
          });
        return flushPromiseRef.current;
      }, [applyEnvelope, patch, rebuildTextCache, reportError]);

      /**
       * Structural operations own the document until their refetch lands.
       * Typing that arrives meanwhile is queued (printable keys) or reconciled
       * from the DOM, never diffed against a pre-split snapshot.
       */
      const structuralPendingRef = React.useRef(0);
      const deferredTextRef = React.useRef("");
      const inputDuringStructuralRef = React.useRef(false);
      const insertAtCaretRef = React.useRef(null);

      const beginStructural = React.useCallback(() => {
        structuralPendingRef.current += 1;
        selectionRuntime.structuralDepth = structuralPendingRef.current;
      }, []);

      const endStructural = React.useCallback(async () => {
        structuralPendingRef.current = Math.max(
          0,
          structuralPendingRef.current - 1,
        );
        if (structuralPendingRef.current > 0) return;
        try {
          if (inputDuringStructuralRef.current) {
            inputDuringStructuralRef.current = false;
            if (stateRef.current.bootstrap) {
              const restore = selectionRef.current;
              await reload(restore);
              if (restore && placeCaretRef.current) {
                placeCaretRef.current(restore);
                if (placementPromiseRef.current) {
                  await Promise.race([
                    placementPromiseRef.current,
                    new Promise((resolve) => setTimeout(resolve, 300)),
                  ]);
                  placementPromiseRef.current = null;
                }
              }
              await flushPending();
            }
          }
          const deferred = deferredTextRef.current;
          deferredTextRef.current = "";
          if (deferred) {
            // The structural operation placed the caret from the fresh model;
            // wait for that layout-effect placement before inserting.
            if (placementPromiseRef.current) {
              await Promise.race([
                placementPromiseRef.current,
                new Promise((resolve) => setTimeout(resolve, 300)),
              ]);
              placementPromiseRef.current = null;
            }
            await flushPending();
            if (insertAtCaretRef.current)
              await insertAtCaretRef.current(deferred);
            else await flushPending();
          }
        } catch (_error) {
          /* the operation that raised already reported */
        }
      }, [flushPending, reload]);

      const loadFallback = React.useCallback(
        async (documentId) => {
          // Older Hosts have no /bootstrap. Fall back to the legacy /document
          // route so the panel still shows the document and its exports.
          try {
            const payload = await apiClient.document(documentId);
            const metadata = payload.metadata || {};
            const inner = metadata.metadata || metadata;
            const statsSource = metadata.stats || metadata.statistics || {};
            patch({
              fallback: {
                title:
                  (inner && inner.title) ||
                  payload.title ||
                  (payload.sourceFormat ? "Untitled document" : "Document"),
                text: payload.text || "",
                truncated: payload.truncated === true,
                outline: payload.outline || [],
                stats: statsSource,
              },
            });
          } catch (_cause) {
            patch({ fallback: null });
          }
        },
        [patch],
      );

      const loadDocument = React.useCallback(
        async (documentId) => {
          if (!documentId) return;
          // Push local edits before leaving the current document so switching
          // or closing never silently drops typed text.
          await flushPending();
          // A different document has different blocks: never carry pending
          // text across, even when the flush above failed.
          const previousId =
            stateRef.current.bootstrap && stateRef.current.bootstrap.documentId;
          if (previousId && previousId !== documentId)
            pendingTextRef.current.clear();
          patch({
            selectedId: documentId,
            bootstrap: null,
            selection: null,
            fallback: null,
          });
          const bootstrap = await reload(null, { documentId });
          if (!bootstrap) await loadFallback(documentId);
          return bootstrap;
        },
        [flushPending, loadFallback, patch, reload],
      );

      const performOp = React.useCallback(
        async (op, args, options) => {
          const current = stateRef.current;
          const bootstrap = current.bootstrap;
          if (!bootstrap) {
            reportError(HostError("NO_DOCUMENT", REASONS.noDocument));
            return null;
          }
          if (isReadOnly(current)) {
            reportError(HostError("READ_ONLY", REASONS.readOnly));
            return null;
          }
          if (offlineRef.current) {
            reportError(HostError("OFFLINE", REASONS.offline));
            return null;
          }
          beginStructural();
          try {
            await flushPending();
            patch({ saving: true });
            const result = await apiClient.op(
              bootstrap.documentId,
              op,
              args || {},
            );
            applyEnvelope(result);
            if (options && options.reload) {
              const caret =
                options.caret === undefined
                  ? selectionRef.current
                  : options.caret;
              await reload(caret);
            }
            return result;
          } catch (cause) {
            reportError(cause);
            try {
              await reload(selectionRef.current);
            } catch (_error) {
              /* the error bar already reports the failure */
            }
            return null;
          } finally {
            patch({ saving: false });
            await endStructural();
          }
        },
        [
          applyEnvelope,
          beginStructural,
          endStructural,
          flushPending,
          isReadOnly,
          patch,
          reload,
          reportError,
        ],
      );

      const readSelection = React.useCallback(() => {
        const selection = domSelectionToModel(surfaceRef.current);
        if (!selection) return selectionRef.current;
        const current = stateRef.current;
        const order = selectionOrderOf(
          current.bootstrap ? current.bootstrap.document : {},
        );
        return normalizeSelection(selection, order) || selection;
      }, []);

      const placeCaret = React.useCallback(
        (selection) => {
          selectionRef.current = selection;
          placementSeqRef.current += 1;
          pendingPlacementRef.current = true;
          placementPromiseRef.current = new Promise((resolve) => {
            placementResolveRef.current = resolve;
          });
          // The caret is placed by a layout effect after the commit that renders
          // the refetched paragraphs. A timer/rAF can run either side of that
          // commit; a layout effect cannot, so typing that follows a structural
          // operation always lands in the paragraph the model selected.
          patch({
            selection,
            caretEpoch: (stateRef.current.caretEpoch || 0) + 1,
          });
        },
        [patch],
      );

      placeCaretRef.current = placeCaret;

      const updateSelection = React.useCallback(() => {
        // A programmatic selection write owns the next read: the focus event it
        // causes must not publish the previous caret.
        if (selectionWriteDepth > 0 || pendingPlacementRef.current) return;
        const selection = readSelection();
        if (!selection) return;
        selectionRef.current = selection;
        const current = stateRef.current;
        const previous = current.selection;
        if (
          previous &&
          previous.start &&
          previous.end &&
          previous.start.blockId === selection.start.blockId &&
          previous.end.blockId === selection.end.blockId &&
          previous.start.offset === selection.start.offset &&
          previous.end.offset === selection.end.offset
        )
          return;
        selectionRuntime.lastUpdate = selection;
        patch({ selection });
      }, [patch, readSelection]);

      const scheduleFlush = React.useCallback(() => {
        if (flushTimerRef.current) clearTimeout(flushTimerRef.current);
        flushTimerRef.current = setTimeout(() => {
          flushTimerRef.current = null;
          void flushPending();
        }, 350);
      }, [flushPending]);

      /**
       * Reconcile one edited paragraph from the DOM into the model.
       *
       * The model is the only source the flush reads, and the flush diffs it
       * against the text the Host confirmed, so applying the DOM text to the
       * model must happen exactly once per input event. `force` lets
       * compositionend read the final text when the browser commits without a
       * trailing `input` event.
       */
      const syncBlockFromDom = React.useCallback(
        (blockElement, options) => {
          const current = stateRef.current;
          if (!current.bootstrap || isReadOnly(current)) return false;
          if (!blockElement || !surfaceRef.current) return false;
          if (!surfaceRef.current.contains(blockElement)) return false;
          if (blockElement.getAttribute("contenteditable") !== "true")
            return false;
          const blockId = blockElement.getAttribute("data-block-id");
          const { byId } = collectBlocks(current.bootstrap.document);
          const entry = byId.get(blockId);
          if (!entry || !entry.block || entry.block.type !== "paragraph")
            return false;
          const text = domTextOfBlock(blockElement);
          const previous = paragraphText(entry.block);
          if (text === previous) return false;
          const diff = diffTextRange(previous, text);
          if (!diff) return false;
          applyTextReplace(entry.block, diff.start, diff.end, diff.insert);
          const pending = pendingTextRef.current;
          if (pending.get(blockId) !== text) pending.set(blockId, text);
          const selection = domSelectionToModel(surfaceRef.current);
          if (selection) selectionRef.current = selection;
          if (options && options.schedule === false) return true;
          scheduleFlush();
          patch({
            saveState: "unsaved",
            selection: selectionRef.current,
            canUndo: true,
          });
          return true;
        },
        [isReadOnly, patch, scheduleFlush],
      );

      const handleInput = React.useCallback(
        (event) => {
          const current = stateRef.current;
          if (!current.bootstrap || isReadOnly(current)) return;
          // The browser owns the text while the IME is composing. Diffing the
          // intermediate buffer is what duplicated text: the model, the
          // composition buffer and the final commit each added characters.
          if (composingRef.current || eventIsComposing(event)) {
            const selection = domSelectionToModel(surfaceRef.current);
            if (selection) selectionRef.current = selection;
            return;
          }
          const target = event && event.target;
          if (!target || typeof target.closest !== "function") return;
          const blockElement = target.closest("[data-block-id]");
          if (!blockElement) return;
          if (blockElement.getAttribute("contenteditable") !== "true") return;
          if (structuralPendingRef.current > 0) {
            // The document is mid-refetch: the DOM may be a remount with the
            // caret at offset 0, so its text is not authoritative. Queue only
            // the inserted characters; they are typed into the post-refetch
            // caret when the structural operation finishes.
            inputDuringStructuralRef.current = true;
            selectionRuntime.structuralInputs =
              (selectionRuntime.structuralInputs || 0) + 1;
            const blockId = blockElement.getAttribute("data-block-id");
            const { byId } = collectBlocks(current.bootstrap.document);
            const entry = byId.get(blockId);
            if (entry && entry.block && entry.block.type === "paragraph") {
              const modelText = paragraphText(entry.block);
              const dom = domTextOfBlock(blockElement);
              if (dom !== modelText) {
                const diff = diffTextRange(modelText, dom);
                if (diff && diff.end === diff.start && diff.insert) {
                  deferredTextRef.current += diff.insert;
                  selectionRuntime.deferredChars =
                    (selectionRuntime.deferredChars || 0) + diff.insert.length;
                }
              }
            }
            return;
          }
          syncBlockFromDom(blockElement);
        },
        [isReadOnly, syncBlockFromDom],
      );

      const handleBeforeInput = React.useCallback((event) => {
        if (structuralPendingRef.current === 0) return;
        const type = (event && event.inputType) || "";
        if (type === "insertText" || type === "insertCompositionText") {
          event.preventDefault();
          deferredTextRef.current += (event && event.data) || "";
          return;
        }
        if (type.startsWith("delete")) event.preventDefault();
      }, []);

      const handleCompositionStart = React.useCallback(() => {
        composingRef.current = true;
      }, []);

      const handleCompositionEnd = React.useCallback(
        (event) => {
          composingRef.current = false;
          const target = event && event.target;
          const blockElement =
            target && typeof target.closest === "function"
              ? target.closest("[data-block-id]")
              : null;
          if (blockElement) syncBlockFromDom(blockElement);
        },
        [syncBlockFromDom],
      );

      const applyRunPropsToSelection = React.useCallback(
        async (props, options) => {
          const current = stateRef.current;
          const document = current.bootstrap
            ? current.bootstrap.document
            : null;
          if (!document) {
            reportError(HostError("NO_DOCUMENT", REASONS.noDocument));
            return;
          }
          if (isReadOnly(current)) {
            reportError(HostError("READ_ONLY", REASONS.readOnly));
            return;
          }
          const selection = normalizeSelection(
            selectionRef.current,
            selectionOrderOf(document),
          );
          // Word behaviour: a collapsed caret formats the caret paragraph so
          // typing inherits the change; a real selection formats exactly that
          // range. `formatTextOpsForSelection` owns both cases.
          const ops = formatTextOpsForSelection(
            document,
            selection,
            props,
            options,
          );
          if (ops.length === 0) {
            reportError(HostError("CARET_REQUIRED", REASONS.caretRequired));
            return;
          }
          await flushPending();
          patch({ saving: true });
          try {
            const result = await apiClient.ops(
              current.bootstrap.documentId,
              ops,
            );
            applyEnvelope(result);
            await reload(selection);
          } catch (cause) {
            reportError(cause);
            try {
              await reload(selection);
            } catch (_error) {
              /* reported above */
            }
          } finally {
            patch({ saving: false });
          }
        },
        [applyEnvelope, flushPending, isReadOnly, patch, reload, reportError],
      );

      const applyParagraphProps = React.useCallback(
        async (props) => {
          const current = stateRef.current;
          const document = current.bootstrap
            ? current.bootstrap.document
            : null;
          if (!document) {
            reportError(HostError("NO_DOCUMENT", REASONS.noDocument));
            return;
          }
          if (isReadOnly(current)) {
            reportError(HostError("READ_ONLY", REASONS.readOnly));
            return;
          }
          const selection = normalizeSelection(
            selectionRef.current,
            selectionOrderOf(document),
          );
          const { list } = collectBlocks(document);
          let blockIds = [];
          if (selection) {
            const startIndex = list.findIndex(
              (entry) => entry.block.id === selection.start.blockId,
            );
            const endIndex = list.findIndex(
              (entry) => entry.block.id === selection.end.blockId,
            );
            for (
              let index = Math.max(0, startIndex);
              index <= Math.max(startIndex, endIndex) && index < list.length;
              index += 1
            ) {
              const entry = list[index];
              if (entry.block && entry.block.type === "paragraph")
                blockIds.push(entry.block.id);
            }
          }
          if (blockIds.length === 0) {
            reportError(HostError("NO_BLOCK", REASONS.noDocument));
            return;
          }
          await flushPending();
          patch({ saving: true });
          try {
            const result = await apiClient.op(
              current.bootstrap.documentId,
              "formatParagraph",
              { blockIds, props },
            );
            applyEnvelope(result);
            await reload(selection || undefined);
          } catch (cause) {
            reportError(cause);
            try {
              await reload(selection || undefined);
            } catch (_error) {
              /* reported above */
            }
          } finally {
            patch({ saving: false });
          }
        },
        [applyEnvelope, flushPending, isReadOnly, patch, reload, reportError],
      );

      const caretPosition = React.useCallback(() => {
        const current = stateRef.current;
        const document = current.bootstrap ? current.bootstrap.document : null;
        if (!document) return null;
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(document),
        );
        if (selection) return selection.start;
        const first = documentBodyBlocks(document)[0];
        return first ? { blockId: first.id, offset: 0 } : null;
      }, []);

      const rangeForSelection = React.useCallback(() => {
        const current = stateRef.current;
        const document = current.bootstrap ? current.bootstrap.document : null;
        if (!document) return null;
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(document),
        );
        if (!selection) return null;
        return { start: selection.start, end: selection.end };
      }, []);

      const selectionText = React.useCallback(() => {
        const current = stateRef.current;
        const document = current.bootstrap ? current.bootstrap.document : null;
        if (!document) return "";
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(document),
        );
        if (!selection || isCollapsed(selection)) return "";
        const { list } = collectBlocks(document);
        const startIndex = list.findIndex(
          (entry) => entry.block.id === selection.start.blockId,
        );
        const endIndex = list.findIndex(
          (entry) => entry.block.id === selection.end.blockId,
        );
        if (startIndex < 0 || endIndex < 0) return "";
        const pieces = [];
        for (let index = startIndex; index <= endIndex; index += 1) {
          const entry = list[index];
          if (!entry.block || entry.block.type !== "paragraph") continue;
          const text = paragraphText(entry.block);
          const from = index === startIndex ? selection.start.offset : 0;
          const to = index === endIndex ? selection.end.offset : text.length;
          pieces.push(
            text.slice(
              Math.min(from, to),
              Math.max(Math.min(from, to), Math.max(from, to)),
            ),
          );
        }
        return pieces.join("\n");
      }, []);

      const selectAllDocument = React.useCallback(() => {
        const current = stateRef.current;
        const document = current.bootstrap ? current.bootstrap.document : null;
        if (!document) return;
        const list = collectBlocks(document).list.filter(
          (entry) => entry.block && entry.block.type === "paragraph",
        );
        if (list.length === 0) return;
        const first = list[0].block;
        const last = list[list.length - 1].block;
        const selection = {
          start: { blockId: first.id, offset: 0 },
          end: { blockId: last.id, offset: paragraphText(last).length },
        };
        placeCaret(selection);
      }, [placeCaret]);

      /**
       * Publish a selection in the model *and* in the DOM.
       *
       * A DOM Range may span two paragraph contenteditables, which is exactly
       * what makes a cross-paragraph selection visible and copyable; the model
       * range is what every command reads. Writing both here keeps them in
       * lockstep (the old code only wrote the DOM, so a drag that crossed a
       * paragraph boundary produced a DOM selection the panel never learned).
       */
      const applyModelSelection = React.useCallback(
        (startPoint, endPoint, options) => {
          const current = stateRef.current;
          const document = current.bootstrap
            ? current.bootstrap.document
            : null;
          if (!document || !startPoint) {
            selectionRuntime.failed += 1;
            selectionRuntime.lastError = "apply:no-document";
            return null;
          }
          const order = selectionOrderOf(document);
          const normalized = normalizeSelection(
            { start: startPoint, end: endPoint || startPoint },
            order,
          );
          if (!normalized) {
            selectionRuntime.failed += 1;
            selectionRuntime.lastError = "apply:not-normalised";
            return null;
          }
          selectionRuntime.applied += 1;
          selectionRuntime.lastApplied = normalized;
          selectionRuntime.patchCount += 1;
          selectionRuntime.lastPoint = normalized.end;
          selectionRef.current = normalized;
          const root = surfaceRef.current;
          const opts = options || {};
          if (root && opts.dom !== false) {
            const startElement = findBlockElement(
              root,
              normalized.start.blockId,
            );
            const range = domRangeForPoints(
              root,
              normalized.start,
              normalized.end,
            );
            if (startElement && typeof startElement.focus === "function") {
              beginSelectionWrite();
              try {
                startElement.focus({ preventScroll: true });
              } catch (_error) {
                /* focus is best-effort; the range still applies */
              } finally {
                endSelectionWrite();
              }
            }
            const selected = range ? selectDomRange(root, range) : false;
            selectionRuntime.lastApply = {
              start: normalized.start,
              end: normalized.end,
              hasRange: Boolean(range),
              selected,
              domText: (() => {
                if (typeof window === "undefined" || !window.getSelection)
                  return "";
                const domSelection = window.getSelection();
                return domSelection ? String(domSelection) : "";
              })(),
            };
          }
          patch({ selection: normalized });
          return normalized;
        },
        [patch],
      );

      /** Select the whole paragraph under a pointer point (triple click). */
      const selectBlockAtPoint = React.useCallback(
        (point) => {
          const current = stateRef.current;
          const document = current.bootstrap
            ? current.bootstrap.document
            : null;
          if (!document || !point) return;
          const entry = collectBlocks(document).byId.get(point.blockId);
          if (!entry || !entry.block) return;
          const end = paragraphText(entry.block).length;
          selectionAnchorRef.current = {
            anchor: { blockId: point.blockId, offset: 0 },
            focus: { blockId: point.blockId, offset: end },
          };
          applyModelSelection(
            { blockId: point.blockId, offset: 0 },
            { blockId: point.blockId, offset: end },
            { dom: true },
          );
        },
        [applyModelSelection],
      );

      /** Select the word under a pointer point (double click). */
      const selectWordAtPoint = React.useCallback(
        (point) => {
          const current = stateRef.current;
          const document = current.bootstrap
            ? current.bootstrap.document
            : null;
          if (!document || !point) return;
          const entry = collectBlocks(document).byId.get(point.blockId);
          if (!entry || !entry.block) return;
          const range = wordRangeAt(paragraphText(entry.block), point.offset);
          selectionAnchorRef.current = {
            anchor: { blockId: point.blockId, offset: range.start },
            focus: { blockId: point.blockId, offset: range.end },
          };
          applyModelSelection(
            { blockId: point.blockId, offset: range.start },
            { blockId: point.blockId, offset: range.end },
            { dom: true },
          );
        },
        [applyModelSelection],
      );

      const handleSurfaceMouseDown = React.useCallback(
        (event) => {
          if (!event || event.button !== 0) return;
          const current = stateRef.current;
          if (!current.bootstrap || isReadOnly(current)) return;
          const root = surfaceRef.current;
          if (!root) return;
          const target = event.target;
          if (
            target &&
            typeof target.closest === "function" &&
            target.closest("input, textarea, select")
          )
            return;
          selectionRuntime.pointerDown += 1;
          const point = selectionPointAt(root, event.clientX, event.clientY);
          selectionRuntime.lastDown = {
            x: event.clientX,
            y: event.clientY,
            detail: event.detail,
            point,
          };
          if (!point || !point.blockId) {
            selectionRuntime.failed += 1;
            selectionRuntime.lastError = "down:no-point";
            return;
          }
          // The panel owns the mouse selection. Each paragraph is its own
          // contenteditable, so Chrome's native drag clamps at that boundary and
          // overrides a programmatic cross-block Range with its own anchor.
          // Cancelling the default action removes that competing selection; the
          // caret/range is then placed from the model, which is also what makes
          // a drag into the next paragraph a real multi-block selection.
          event.preventDefault();
          if (event.detail >= 3) {
            selectBlockAtPoint(point);
            return;
          }
          if (event.detail === 2) {
            selectWordAtPoint(point);
            return;
          }
          if (event.shiftKey && selectionRef.current) {
            const document = current.bootstrap.document;
            const normalized = normalizeSelection(
              selectionRef.current,
              selectionOrderOf(document),
            );
            const anchor = normalized ? normalized.start : point;
            selectionAnchorRef.current = { anchor, focus: point };
            applyModelSelection(anchor, point, { dom: true });
            return;
          }
          selectionAnchorRef.current = null;
          dragSelectRef.current = { anchor: point, active: true };
          applyModelSelection(point, point, { dom: true });
        },
        [
          applyModelSelection,
          isReadOnly,
          selectBlockAtPoint,
          selectWordAtPoint,
        ],
      );

      const handleSurfaceDoubleClick = React.useCallback(
        (event) => {
          if (!event || isReadOnly(stateRef.current)) return;
          const root = surfaceRef.current;
          if (!root) return;
          const point = selectionPointAt(root, event.clientX, event.clientY);
          if (!point || !point.blockId) return;
          event.preventDefault();
          selectWordAtPoint(point);
        },
        [isReadOnly, selectWordAtPoint],
      );

      /** Continue a drag: every move writes the model range and the DOM range. */
      const extendDragSelection = React.useCallback(
        (event) => {
          const drag = dragSelectRef.current;
          selectionRuntime.pointerMove += 1;
          selectionRuntime.lastMove = {
            x: event.clientX,
            y: event.clientY,
            buttons: event.buttons,
            drag: Boolean(drag),
          };
          if (!drag) return;
          if (!(event.buttons & 1)) return;
          const root = surfaceRef.current;
          if (!root) return;
          const point = selectionPointAt(root, event.clientX, event.clientY);
          if (!point || !point.blockId) return;
          if (composingRef.current) return;
          event.preventDefault();
          selectionAnchorRef.current = { anchor: drag.anchor, focus: point };
          applyModelSelection(drag.anchor, point, { dom: true });
        },
        [applyModelSelection],
      );

      const finishDragSelection = React.useCallback(() => {
        const drag = dragSelectRef.current;
        if (!drag) return;
        dragSelectRef.current = null;
        if (!drag.active) return;
        updateSelection();
      }, [updateSelection]);

      const extendSelectionByKey = React.useCallback(
        (key) => {
          const current = stateRef.current;
          const document = current.bootstrap
            ? current.bootstrap.document
            : null;
          if (!document) return;
          const normalized = normalizeSelection(
            selectionRef.current,
            selectionOrderOf(document),
          );
          const previous = selectionAnchorRef.current;
          const anchor = previous
            ? previous.anchor
            : normalized
              ? normalized.start
              : null;
          const focusStart = previous
            ? previous.focus
            : normalized
              ? normalized.end
              : null;
          if (!anchor || !focusStart) return;
          const focus = moveSelectionFocus(document, focusStart, key);
          if (!focus) return;
          selectionAnchorRef.current = { anchor, focus };
          applyModelSelection(anchor, focus, { dom: true });
        },
        [applyModelSelection],
      );

      // A native beforeinput listener is required: React's synthetic
      // onBeforeInput is not dispatched for CDP/IME text insertion, and a text
      // insertion that lands before the post-refetch caret is placed rewrites
      // the wrong paragraph.
      React.useEffect(() => {
        const surface = surfaceRef.current;
        if (!surface || typeof surface.addEventListener !== "function")
          return undefined;
        const onBefore = (event) => {
          if (structuralPendingRef.current === 0) return;
          const type = event.inputType || "";
          if (type === "insertText" || type === "insertCompositionText") {
            event.preventDefault();
            deferredTextRef.current += event.data || "";
            return;
          }
          if (type.startsWith("delete")) event.preventDefault();
        };
        surface.addEventListener("beforeinput", onBefore, true);
        return () => surface.removeEventListener("beforeinput", onBefore, true);
      }, []);

      React.useLayoutEffect(() => {
        if (!state.caretEpoch) return;
        try {
          const ok = Boolean(
            surfaceRef.current &&
            selectionRef.current &&
            placeSelection(surfaceRef.current, selectionRef.current),
          );
          selectionRuntime.lastPlacement = {
            selection: selectionRef.current,
            ok,
          };
        } finally {
          pendingPlacementRef.current = false;
          if (placementResolveRef.current) {
            placementResolveRef.current();
            placementResolveRef.current = null;
          }
        }
      }, [state.caretEpoch]);

      React.useEffect(() => {
        if (typeof window === "undefined" || !window.addEventListener)
          return undefined;
        const onMove = (event) => extendDragSelection(event);
        const onUp = () => finishDragSelection();
        window.addEventListener("mousemove", onMove, true);
        window.addEventListener("mouseup", onUp, true);
        return () => {
          window.removeEventListener("mousemove", onMove, true);
          window.removeEventListener("mouseup", onUp, true);
        };
      }, [extendDragSelection, finishDragSelection]);

      const deleteSelectedRange = React.useCallback(
        async (selection) => {
          const current = stateRef.current;
          if (!current.bootstrap) return;
          beginStructural();
          try {
            await flushPending();
            patch({ saving: true });
            const result = await apiClient.op(
              current.bootstrap.documentId,
              "deleteRange",
              { range: { start: selection.start, end: selection.end } },
            );
            applyEnvelope(result);
            await reload({
              start: selection.start,
              end: selection.start,
            });
          } catch (cause) {
            reportError(cause);
          } finally {
            patch({ saving: false });
            await endStructural();
          }
        },
        [
          applyEnvelope,
          beginStructural,
          endStructural,
          flushPending,
          patch,
          reload,
          reportError,
        ],
      );

      const splitParagraph = React.useCallback(async () => {
        const current = stateRef.current;
        if (!current.bootstrap) return;
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(current.bootstrap.document),
        );
        if (!selection) return;
        const blockId = selection.start.blockId;
        const offset = selection.start.offset;
        beginStructural();
        try {
          await flushPending();
          patch({ saving: true });
          const result = await apiClient.op(
            current.bootstrap.documentId,
            "insertText",
            { position: { blockId, offset }, text: "\n" },
          );
          applyEnvelope(result);
          const created = newBlockIdFromOpResult(result, blockId);
          selectionRuntime.lastSplit = { blockId, created };
          const bootstrap = await reload(null, { skipPlacement: true });
          let targetId = created;
          if (!targetId && bootstrap) {
            const list = collectBlocks(bootstrap.document).list;
            const index = list.findIndex((entry) => entry.block.id === blockId);
            for (let next = index + 1; next < list.length; next += 1) {
              if (list[next].block && list[next].block.type === "paragraph") {
                targetId = list[next].block.id;
                break;
              }
            }
          }
          if (targetId)
            placeCaret({
              start: { blockId: targetId, offset: 0 },
              end: { blockId: targetId, offset: 0 },
            });
        } catch (cause) {
          reportError(cause);
        } finally {
          patch({ saving: false });
          await endStructural();
        }
      }, [
        applyEnvelope,
        beginStructural,
        endStructural,
        flushPending,
        patch,
        placeCaret,
        reload,
        reportError,
      ]);

      /**
       * Word's Shift+Enter: a line break inside the current paragraph. The Host
       * accepts `insertPageBreak` with `kind:"line"`, which splices one `break`
       * inline at the caret; afterwards the caret sits immediately after that
       * break, still in the same paragraph.
       */
      const insertLineBreakAtCaret = React.useCallback(async () => {
        const current = stateRef.current;
        if (!current.bootstrap) return;
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(current.bootstrap.document),
        );
        if (!selection) return;
        const blockId = selection.start.blockId;
        const offset = selection.start.offset || 0;
        // The break is one character in the offset space, so "after the break"
        // is offset + 1 in the same paragraph. Passing it as the reload caret
        // lets the structural-refresh helper restore it once the new model has
        // been committed (placing it from this callback alone loses the race
        // with that commit and lands before the break).
        const target = {
          start: { blockId, offset: offset + 1 },
          end: { blockId, offset: offset + 1 },
        };
        const result = await performOp(
          "insertPageBreak",
          { position: { blockId, offset }, kind: "line" },
          { reload: true, caret: target },
        );
        if (!result) return;
        placeCaret(target);
      }, [performOp, placeCaret]);

      const mergeBackward = React.useCallback(async () => {
        const current = stateRef.current;
        const document = current.bootstrap ? current.bootstrap.document : null;
        if (!document) return;
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(document),
        );
        if (!selection || !isCollapsed(selection)) return;
        const blockId = selection.start.blockId;
        const offset = selection.start.offset;
        if (offset > 0) return;
        const { list } = collectBlocks(document);
        const index = list.findIndex((entry) => entry.block.id === blockId);
        let previousIndex = -1;
        for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
          if (list[cursor].block && list[cursor].block.type === "paragraph") {
            previousIndex = cursor;
            break;
          }
        }
        if (previousIndex < 0) return;
        const previous = list[previousIndex].block;
        const end = paragraphText(previous).length;
        await deleteSelectedRange({
          start: { blockId: previous.id, offset: end },
          end: { blockId, offset },
        });
        placeCaret({
          start: { blockId: previous.id, offset: end },
          end: { blockId: previous.id, offset: end },
        });
      }, [deleteSelectedRange, placeCaret]);

      const mergeForward = React.useCallback(async () => {
        const current = stateRef.current;
        const document = current.bootstrap ? current.bootstrap.document : null;
        if (!document) return;
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(document),
        );
        if (!selection || !isCollapsed(selection)) return;
        const blockId = selection.start.blockId;
        const offset = selection.start.offset;
        const { list, byId } = collectBlocks(document);
        const entry = byId.get(blockId);
        if (!entry || entry.block.type !== "paragraph") return;
        if (offset < paragraphText(entry.block).length) return;
        const index = list.findIndex((item) => item.block.id === blockId);
        let nextIndex = -1;
        for (let cursor = index + 1; cursor < list.length; cursor += 1) {
          if (list[cursor].block && list[cursor].block.type === "paragraph") {
            nextIndex = cursor;
            break;
          }
        }
        if (nextIndex < 0) return;
        const next = list[nextIndex].block;
        await deleteSelectedRange({
          start: { blockId, offset },
          end: { blockId: next.id, offset: 0 },
        });
        placeCaret({
          start: { blockId, offset },
          end: { blockId, offset },
        });
      }, [deleteSelectedRange, placeCaret]);

      const insertAtCaret = React.useCallback(
        async (text, options) => {
          const current = stateRef.current;
          if (!current.bootstrap) return;
          const selection = normalizeSelection(
            selectionRef.current,
            selectionOrderOf(current.bootstrap.document),
          );
          if (!selection) return;
          beginStructural();
          try {
            await flushPending();
            patch({ saving: true });
            const result = isCollapsed(selection)
              ? await apiClient.op(current.bootstrap.documentId, "insertText", {
                  position: selection.start,
                  text,
                  props: options && options.props,
                })
              : await apiClient.op(
                  current.bootstrap.documentId,
                  "replaceText",
                  {
                    range: { start: selection.start, end: selection.end },
                    text,
                  },
                );
            applyEnvelope(result);
            const bootstrap = await reload(null, { skipPlacement: true });
            if (bootstrap) {
              const entry = collectBlocks(bootstrap.document).byId.get(
                selection.start.blockId,
              );
              const length = entry ? paragraphText(entry.block).length : 0;
              const offset = Math.min(
                selection.start.offset + String(text).length,
                length,
              );
              placeCaret({
                start: { blockId: selection.start.blockId, offset },
                end: { blockId: selection.start.blockId, offset },
              });
            }
          } catch (cause) {
            reportError(cause);
          } finally {
            patch({ saving: false });
            await endStructural();
          }
        },
        [
          applyEnvelope,
          beginStructural,
          endStructural,
          flushPending,
          patch,
          placeCaret,
          reload,
          reportError,
        ],
      );
      insertAtCaretRef.current = insertAtCaret;

      const handlePaste = React.useCallback(
        async (event) => {
          const current = stateRef.current;
          if (!current.bootstrap || isReadOnly(current)) return;
          // Always swallow the native paste: only `text/plain` ever reaches the
          // model, so pasted HTML can never be injected into the document.
          if (event && typeof event.preventDefault === "function")
            event.preventDefault();
          const clipboard =
            event && event.clipboardData
              ? event.clipboardData.getData("text/plain")
              : "";
          if (!clipboard) return;
          await flushPending();
          await insertAtCaret(clipboard);
        },
        [flushPending, insertAtCaret, isReadOnly],
      );

      const handleBackspace = React.useCallback(async () => {
        const current = stateRef.current;
        if (!current.bootstrap || isReadOnly(current)) return;
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(current.bootstrap.document),
        );
        if (!selection) return;
        if (!isCollapsed(selection)) {
          await deleteSelectedRange(selection);
          return;
        }
        await mergeBackward();
      }, [deleteSelectedRange, isReadOnly, mergeBackward]);

      const handleDelete = React.useCallback(async () => {
        const current = stateRef.current;
        if (!current.bootstrap || isReadOnly(current)) return;
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(current.bootstrap.document),
        );
        if (!selection) return;
        if (!isCollapsed(selection)) {
          await deleteSelectedRange(selection);
          return;
        }
        await mergeForward();
      }, [deleteSelectedRange, isReadOnly, mergeForward]);

      const handleSurfaceCopy = React.useCallback(
        (event) => {
          const text = selectionText();
          if (!text) return;
          if (
            event &&
            event.clipboardData &&
            typeof event.clipboardData.setData === "function"
          ) {
            // Plain text only: pasted HTML must never enter the model, and a
            // multi-block selection joins paragraphs with "\n".
            event.clipboardData.setData("text/plain", text);
            event.preventDefault();
          }
        },
        [selectionText],
      );

      const handleSurfaceCut = React.useCallback(
        (event) => {
          if (isReadOnly(stateRef.current)) return;
          const text = selectionText();
          if (!text) return;
          if (
            event &&
            event.clipboardData &&
            typeof event.clipboardData.setData === "function"
          ) {
            event.clipboardData.setData("text/plain", text);
            event.preventDefault();
          }
          void handleBackspace();
        },
        [handleBackspace, isReadOnly, selectionText],
      );

      const handleKeyDown = React.useCallback(
        (event) => {
          // The same vetoes the capture listener applies: a composition or a
          // modal outside the panel owns the keystroke, and the React fallback
          // must not apply it a second time (or at all).
          if (eventIsComposing(event) || composingRef.current) return;
          if (modalOutsidePanel(rootRef.current)) return;
          const current = stateRef.current;
          // The capture listener intentionally lets a "must not repeat" command
          // through on auto-repeat; the fallback must not apply it a second time.
          const repeatEntry = shortcutForCombo(shortcutCombo(event));
          if (repeatEntry && repeatEntry.repeat === false && event.repeat)
            return;
          const target = event.target;
          const inDialog = Boolean(current.dialog);
          const inEditable =
            target &&
            typeof target.closest === "function" &&
            Boolean(target.closest("[data-block-id]"));
          const mod = event.ctrlKey || event.metaKey;
          if (structuralPendingRef.current > 0 && event.key !== "Escape") {
            const printable =
              !mod &&
              !event.altKey &&
              typeof event.key === "string" &&
              event.key.length === 1;
            if (event.key === "Enter") {
              event.preventDefault();
              deferredTextRef.current += "\n";
              return;
            }
            if (printable) {
              event.preventDefault();
              deferredTextRef.current += event.key;
              return;
            }
            if (
              event.key === "Backspace" ||
              event.key === "Delete" ||
              event.key === "Tab"
            ) {
              event.preventDefault();
              return;
            }
          }
          // Ctrl+O / Cmd+O opens the in-panel chooser (the panel's own "Open
          // file…"), never the OS dialog. A text field or the contenteditable
          // document owns the keystroke, and an open dialog owns everything.
          const typing =
            target &&
            typeof target.closest === "function" &&
            Boolean(
              target.closest(
                "input:not([type=file]), textarea, [contenteditable=true], [contenteditable='']",
              ),
            );
          if (inDialog) {
            if (event.key === "Escape") {
              event.preventDefault();
              event.stopPropagation();
              closeDialogRef.current();
            }
            return;
          }
          if (
            mod &&
            (event.key === "o" || event.key === "O") &&
            !typing &&
            typeof openFileChooserRef.current === "function"
          ) {
            event.preventDefault();
            void openFileChooserRef.current();
            return;
          }
          if (mod && (event.key === "s" || event.key === "S")) {
            event.preventDefault();
            void saveDocumentRef.current();
            return;
          }
          if (mod && (event.key === "f" || event.key === "F")) {
            event.preventDefault();
            openDialogRef.current("find", {
              query: "",
              replacement: "",
              matches: [],
              searched: false,
              activeMatch: null,
            });
            return;
          }
          if (mod && (event.key === "a" || event.key === "A") && inEditable) {
            event.preventDefault();
            selectAllDocument();
            return;
          }
          if (mod && inEditable && !isReadOnly(current)) {
            const key = event.key.toLowerCase();
            if (key === "b" || key === "i" || key === "u") {
              event.preventDefault();
              const formatting = readFormatting(
                current.bootstrap ? current.bootstrap.document : null,
                selectionRef.current,
                selectionOrderOf(
                  current.bootstrap ? current.bootstrap.document : {},
                ),
              );
              if (key === "b")
                void applyRunPropsToSelection({ bold: !formatting.bold });
              else if (key === "i")
                void applyRunPropsToSelection({ italic: !formatting.italic });
              else
                void applyRunPropsToSelection({
                  underline: formatting.underline ? "none" : "single",
                });
              return;
            }
            if (key === "z" && event.shiftKey) {
              event.preventDefault();
              void runAction("redo");
              return;
            }
            if (key === "z") {
              event.preventDefault();
              void runAction("undo");
              return;
            }
            if (key === "y") {
              event.preventDefault();
              void runAction("redo");
              return;
            }
            if (key === "k") {
              event.preventDefault();
              void runAction("hyperlink");
              return;
            }
            // Shift changes the reported `key` on most layouts (Shift+7 is "&"
            // on a US keyboard), so the list commands match the physical code.
            if (event.shiftKey && event.code === "Digit7") {
              event.preventDefault();
              void runAction("numbering");
              return;
            }
            if (event.shiftKey && event.code === "Digit8") {
              event.preventDefault();
              void runAction("bullets");
              return;
            }
          }
          if (
            inEditable &&
            !event.altKey &&
            (event.key === "ArrowLeft" ||
              event.key === "ArrowRight" ||
              event.key === "ArrowUp" ||
              event.key === "ArrowDown" ||
              event.key === "Home" ||
              event.key === "End")
          ) {
            if (!event.shiftKey) {
              selectionAnchorRef.current = null;
              return;
            }
            // Shift+Arrow extends the model selection, which is the only way to
            // select *across* paragraphs: each paragraph is its own
            // contenteditable, so the browser would clamp at its boundary.
            event.preventDefault();
            extendSelectionByKey(event.key);
            return;
          }
          // A non-collapsed selection may span paragraphs; the browser would
          // only replace the part inside the focused contenteditable. Route the
          // first printable character through the model so the whole range is
          // replaced in one op.
          if (
            inEditable &&
            !mod &&
            !event.altKey &&
            typeof event.key === "string" &&
            event.key.length === 1
          ) {
            const selection = normalizeSelection(
              selectionRef.current,
              selectionOrderOf(
                current.bootstrap ? current.bootstrap.document : {},
              ),
            );
            if (selection && !isCollapsed(selection)) {
              event.preventDefault();
              void insertAtCaret(event.key);
              return;
            }
          }
          if (!inEditable) return;
          if (event.key === "Enter") {
            event.preventDefault();
            // Word: Enter starts a paragraph, Shift+Enter breaks the line.
            if (event.shiftKey) void insertLineBreakAtCaret();
            else void splitParagraph();
            return;
          }
          if (event.key === "Backspace") {
            const selection = normalizeSelection(
              selectionRef.current,
              selectionOrderOf(
                current.bootstrap ? current.bootstrap.document : {},
              ),
            );
            if (!selection) return;
            if (!isCollapsed(selection)) {
              event.preventDefault();
              void handleBackspace();
              return;
            }
            if ((selection.start.offset || 0) > 0) return;
            event.preventDefault();
            void handleBackspace();
            return;
          }
          if (event.key === "Delete") {
            const selection = normalizeSelection(
              selectionRef.current,
              selectionOrderOf(
                current.bootstrap ? current.bootstrap.document : {},
              ),
            );
            if (!selection) return;
            if (!isCollapsed(selection)) {
              event.preventDefault();
              void handleDelete();
              return;
            }
            const byId = collectBlocks(
              current.bootstrap ? current.bootstrap.document : {},
            ).byId;
            const entry = byId.get(selection.start.blockId);
            const length = entry ? paragraphText(entry.block).length : 0;
            if ((selection.start.offset || 0) < length) return;
            event.preventDefault();
            void handleDelete();
            return;
          }
          if (event.key === "Tab") {
            event.preventDefault();
            void insertAtCaret("\t");
            return;
          }
          if (event.key === "Escape") {
            event.preventDefault();
            if (typeof window !== "undefined" && window.getSelection)
              window.getSelection().removeAllRanges();
            return;
          }
        },
        [
          applyRunPropsToSelection,
          deleteSelectedRange,
          extendSelectionByKey,
          handleBackspace,
          handleDelete,
          insertAtCaret,
          insertLineBreakAtCaret,
          isReadOnly,
          selectAllDocument,
          splitParagraph,
        ],
      );

      /* ------------------------------------------------------- file + io */

      const pushTitle = React.useCallback(
        async (title) => {
          const current = stateRef.current;
          if (!current.bootstrap) return;
          try {
            const result = await apiClient.op(
              current.bootstrap.documentId,
              "setMetadata",
              { patch: { title } },
            );
            applyEnvelope(result);
          } catch (cause) {
            reportError(cause);
          }
        },
        [applyEnvelope, reportError],
      );

      const onTitleChange = React.useCallback(
        (value) => {
          patch((current) => ({
            bootstrap: current.bootstrap
              ? Object.assign({}, current.bootstrap, { title: value })
              : null,
            saveState: "unsaved",
          }));
          if (titleTimerRef.current) clearTimeout(titleTimerRef.current);
          titleTimerRef.current = setTimeout(() => {
            titleTimerRef.current = null;
            const current = stateRef.current;
            if (current.bootstrap) void pushTitle(current.bootstrap.title);
          }, 600);
        },
        [patch, pushTitle],
      );

      const onTitleCommit = React.useCallback(() => {
        if (titleTimerRef.current) {
          clearTimeout(titleTimerRef.current);
          titleTimerRef.current = null;
        }
        const current = stateRef.current;
        if (current.bootstrap) void pushTitle(current.bootstrap.title);
      }, [pushTitle]);

      const saveDocument = React.useCallback(
        async (path) => {
          const current = stateRef.current;
          if (!current.bootstrap) {
            reportError(HostError("NO_DOCUMENT", REASONS.noDocument));
            return;
          }
          if (isReadOnly(current)) {
            reportError(HostError("READ_ONLY", REASONS.readOnly));
            return;
          }
          await flushPending();
          patch({ saving: true, saveState: "saving" });
          try {
            const result = path
              ? await apiClient.saveAs(current.bootstrap.documentId, path)
              : await apiClient.save(current.bootstrap.documentId);
            patch((previous) => ({
              saving: false,
              saveState: "saved",
              dialog: null,
              notice: result.path
                ? `Saved to ${result.path}`
                : "Document saved.",
              bootstrap: previous.bootstrap
                ? Object.assign({}, previous.bootstrap, {
                    dirty: false,
                    revision: result.revision || previous.bootstrap.revision,
                  })
                : null,
            }));
            await refreshList();
          } catch (cause) {
            patch({ saving: false, saveState: "error" });
            reportError(cause);
          }
        },
        [flushPending, isReadOnly, patch, refreshList, reportError],
      );

      const createDocument = React.useCallback(
        async (title) => {
          try {
            const result = await apiClient.create(
              String(title || "").trim() || "Untitled document",
            );
            patch({ dialog: null });
            const list = await refreshList();
            const documentId =
              result.documentId ||
              (list && list.documents && list.documents[0]
                ? list.documents[0].id
                : null);
            if (documentId) await loadDocument(documentId);
            else
              reportError(
                HostError(
                  "CREATE_FAILED",
                  "The Host did not return a document id.",
                ),
              );
          } catch (cause) {
            reportError(cause);
          }
        },
        [loadDocument, patch, refreshList, reportError],
      );

      const openDocumentByPath = React.useCallback(
        async (path) => {
          const trimmed = String(path || "").trim();
          if (!trimmed) return;
          patch({ saving: true, error: null });
          try {
            // Same flow as the folder browser: open through /open-path, switch
            // the workspace to the file's folder, prepare a chat session and
            // then reveal the document (the File backstage must not stay on
            // top of the newly opened document).
            const result = await runOpenDocumentPathFlow(trimmed, {
              uiWorkspace: workspaceService(),
              useBoundServices: false,
            });
            const next = {
              dialog: null,
              saving: false,
              error: null,
              advancedOpen: false,
              tab: "home",
              ribbonCollapsed: false,
            };
            if (result.chip) {
              next.sessionChip = result.chip;
              next.workspace = result.chip.path;
              next.notice = result.chip.warning || "";
            } else if (result.workspaceError) {
              next.notice =
                "The document opened, but the workspace could not be switched: " +
                (result.workspaceError.message || result.workspaceError);
            }
            patch(next);
            await refreshList();
            if (result.opened && result.opened.documentId)
              await loadDocument(result.opened.documentId);
            else
              reportError(
                HostError(
                  "OPEN_FAILED",
                  "The Host did not return a document id.",
                ),
              );
          } catch (cause) {
            patch({ saving: false, error: null });
            reportError(cause);
          }
        },
        [loadDocument, patch, refreshList, reportError],
      );

      const exportDocument = React.useCallback(
        async (format) => {
          const current = stateRef.current;
          if (!current.bootstrap) {
            reportError(HostError("NO_DOCUMENT", REASONS.noDocument));
            return;
          }
          try {
            await downloadExport(
              current.bootstrap.documentId,
              format,
              current.bootstrap.title,
            );
            patch({ notice: `Exported ${String(format).toUpperCase()}.` });
          } catch (cause) {
            reportError(cause);
          }
        },
        [patch, reportError],
      );

      const deleteCurrentDocument = React.useCallback(async () => {
        const current = stateRef.current;
        const documentId =
          (current.bootstrap && current.bootstrap.documentId) ||
          current.selectedId;
        if (!documentId) return;
        try {
          await apiClient.remove(documentId);
          patch({
            dialog: null,
            bootstrap: null,
            selectedId: null,
            selection: null,
            fallback: null,
            notice: "Document deleted.",
            error: null,
          });
          const list = await refreshList();
          if (list && list.documents && list.documents.length)
            await loadDocument(list.documents[0].id);
        } catch (cause) {
          reportError(cause);
        }
      }, [loadDocument, patch, refreshList, reportError]);

      const runSearch = React.useCallback(async () => {
        const current = stateRef.current;
        if (!current.bootstrap || !current.dialog) return;
        try {
          const result = await apiClient.search(
            current.bootstrap.documentId,
            String(current.dialog.query || ""),
          );
          const matches = result.matches || [];
          patch((previous) => ({
            dialog: previous.dialog
              ? Object.assign({}, previous.dialog, {
                  matches,
                  searched: true,
                  activeMatch: matches.length ? 0 : null,
                })
              : null,
          }));
        } catch (cause) {
          reportError(cause);
        }
      }, [patch, reportError]);

      const selectFindMatch = React.useCallback(
        (index) => {
          const current = stateRef.current;
          if (!current.dialog) return;
          const matches = current.dialog.matches || [];
          const match = matches[index];
          if (!match) return;
          patch((previous) => ({
            dialog: previous.dialog
              ? Object.assign({}, previous.dialog, { activeMatch: index })
              : null,
          }));
          placeCaret({
            start: { blockId: match.blockId, offset: match.start },
            end: { blockId: match.blockId, offset: match.end },
          });
        },
        [patch, placeCaret],
      );

      const replaceMatch = React.useCallback(async () => {
        const current = stateRef.current;
        if (!current.bootstrap || !current.dialog) return;
        const matches = current.dialog.matches || [];
        const match = matches[current.dialog.activeMatch];
        if (!match) return;
        const replacement = String(current.dialog.replacement || "");
        await flushPending();
        try {
          const result = await apiClient.op(
            current.bootstrap.documentId,
            "replaceText",
            {
              range: {
                start: { blockId: match.blockId, offset: match.start },
                end: { blockId: match.blockId, offset: match.end },
              },
              text: replacement,
            },
          );
          applyEnvelope(result);
          await reload(null, { skipPlacement: true });
          await runSearch();
        } catch (cause) {
          reportError(cause);
        }
      }, [applyEnvelope, flushPending, reload, reportError, runSearch]);

      /* -------------------------------------------------------- dialogs */

      const openDialog = React.useCallback(
        (kind, extra) => {
          dialogReturnFocusRef.current =
            typeof window !== "undefined"
              ? window.document.activeElement
              : null;
          patch({ dialog: Object.assign({ kind }, extra || {}) });
        },
        [patch],
      );

      const closeDialog = React.useCallback(() => {
        const returnFocus = dialogReturnFocusRef.current;
        patch({ dialog: null });
        dialogReturnFocusRef.current = null;
        if (returnFocus && typeof returnFocus.focus === "function")
          setTimeout(() => returnFocus.focus(), 0);
      }, [patch]);

      const updateDialog = React.useCallback(
        (values) => {
          patch((current) =>
            current.dialog
              ? { dialog: Object.assign({}, current.dialog, values) }
              : {},
          );
        },
        [patch],
      );

      const openDialogRef = React.useRef(openDialog);
      openDialogRef.current = openDialog;
      const closeDialogRef = React.useRef(closeDialog);
      closeDialogRef.current = closeDialog;
      const saveDocumentRef = React.useRef(saveDocument);
      saveDocumentRef.current = saveDocument;
      // The root keydown handler is declared before the chooser exists, so
      // Ctrl+O reaches it through a ref assigned next to its definition.
      const openFileChooserRef = React.useRef(null);

      const onImagePicked = React.useCallback(
        (event) => {
          const file =
            event && event.target && event.target.files
              ? event.target.files[0]
              : null;
          if (!file) return;
          if (SAFE_IMAGE_TYPES.indexOf(String(file.type)) === -1) {
            reportError(
              HostError(
                "UNSUPPORTED_IMAGE",
                `Only PNG, JPEG, GIF, WebP and BMP images can be inserted (got ${file.type || "unknown"}).`,
              ),
            );
            return;
          }
          if (typeof FileReader === "undefined") {
            reportError(
              HostError(
                "IMAGE_READ_FAILED",
                "This browser cannot read local files.",
              ),
            );
            return;
          }
          const reader = new FileReader();
          reader.onload = () => {
            const result = String(reader.result || "");
            const comma = result.indexOf(",");
            const base64 = comma >= 0 ? result.slice(comma + 1) : "";
            patch((current) =>
              current.dialog
                ? {
                    dialog: Object.assign({}, current.dialog, {
                      dataBase64: base64,
                      contentType: file.type,
                      fileName: file.name,
                    }),
                  }
                : {},
            );
          };
          reader.onerror = () =>
            reportError(
              HostError(
                "IMAGE_READ_FAILED",
                "The selected image could not be read.",
              ),
            );
          reader.readAsDataURL(file);
        },
        [patch, reportError],
      );

      /* ------------------------------------- workspaces, sessions, pickers */

      const prepareWorkspaceAt = React.useCallback(
        async (path, create) => {
          const chip = await prepareWorkspaceSession({
            path,
            create,
            uiWorkspace: workspaceService(),
            useBoundServices: false,
          });
          patch({
            sessionChip: chip,
            workspace: chip.path,
            notice: chip.warning || "",
            error: null,
          });
          await refreshList();
          return chip;
        },
        [patch, refreshList],
      );

      const loadBrowse = React.useCallback(
        async (targetPath) => {
          try {
            const payload = await apiClient.browse(targetPath || "");
            patch((current) => ({
              dialog:
                current.dialog && current.dialog.kind === "browse"
                  ? browseDialogState(payload, {
                      newFolder: current.dialog.newFolder,
                    })
                  : current.dialog,
              browseRoots: payload.roots || current.browseRoots,
            }));
            return payload;
          } catch (cause) {
            const missing = isMissingRoute(cause);
            patch((current) => ({
              dialog:
                current.dialog && current.dialog.kind === "browse"
                  ? Object.assign({}, current.dialog, {
                      loading: false,
                      error: describeBrowseFailure(cause),
                    })
                  : current.dialog,
              advancedOpen: missing ? true : current.advancedOpen,
            }));
            return null;
          }
        },
        [patch],
      );

      const openBrowseDialog = React.useCallback(
        async (targetPath) => {
          // Remember where focus came from so Escape can restore it. The
          // chooser opens without going through `openDialog`, so it records
          // its own return target (Ctrl+O from the tab row, a rail button, …).
          if (
            !dialogReturnFocusRef.current &&
            typeof window !== "undefined" &&
            window.document &&
            window.document.activeElement
          )
            dialogReturnFocusRef.current = window.document.activeElement;
          patch({
            dialog: {
              kind: "browse",
              path: targetPath || "",
              parent: null,
              roots: (stateRef.current && stateRef.current.browseRoots) || [],
              entries: [],
              truncated: false,
              loading: true,
              error: null,
              newFolder: null,
            },
          });
          await loadBrowse(targetPath);
        },
        [loadBrowse, patch],
      );

      const loadBrowseRoots = React.useCallback(async () => {
        try {
          const payload = await apiClient.browse();
          patch({ browseRoots: payload.roots || [], advancedOpen: true });
          return payload;
        } catch (cause) {
          patch({
            advancedOpen: true,
            notice: describeBrowseFailure(cause),
          });
          return null;
        }
      }, [patch]);

      const finishOpenedDocument = React.useCallback(
        async (result) => {
          // A real open ends the chooser; its saved focus target must not leak
          // into the next chooser (the button may have been unmounted).
          dialogReturnFocusRef.current = null;
          // Opening from the File backstage must reveal the document again.
          patch({
            dialog: null,
            saving: false,
            error: null,
            tab: "home",
            ribbonCollapsed: false,
          });
          if (result.chip) {
            patch({
              sessionChip: result.chip,
              workspace: result.chip.path,
              notice: result.chip.warning || "",
            });
          } else if (result.workspaceError) {
            // The document is open on the Host; only the workspace switch
            // failed, so say that instead of discarding the document.
            patch({
              notice: `The document opened, but the workspace could not be switched: ${
                result.workspaceError.message || result.workspaceError
              }`,
            });
          }
          await refreshList();
          if (result.documentId) await loadDocument(result.documentId);
          else
            patch({
              notice:
                "The Host opened the file but did not return a document id.",
            });
        },
        [loadDocument, patch, refreshList],
      );

      /**
       * Open the in-panel file chooser (`GET /browse`) at the folder the user
       * last opened a document from, else at the panel workspace (the Host
       * falls back to its own home root when the path is empty). This is the
       * "choose file" path that never depends on the OS file dialog.
       */
      const openFileChooser = React.useCallback(async () => {
        const current = stateRef.current;
        const remembered = readStoredFolder();
        const workspace = (current && current.workspace) || "";
        const target = remembered || workspace;
        if (!remembered && !workspace) {
          patch({
            notice:
              "No remembered folder and no workspace yet; the chooser opens at the Host's own root.",
          });
        }
        await openBrowseDialog(target);
      }, [openBrowseDialog, patch]);
      openFileChooserRef.current = openFileChooser;

      const chooseFolder = React.useCallback(async () => {
        let result;
        try {
          result = await runChooseFolderFlow({
            uiWorkspace: workspaceService(),
            useBoundServices: false,
          });
        } catch (cause) {
          patch({
            advancedOpen: true,
            notice:
              (cause && cause.message) ||
              "The native folder picker is unavailable on this Host.",
          });
          if (!(cause && cause.fallback === "path")) reportError(cause);
          return;
        }
        if (!result || result.cancelled) return;
        patch({
          sessionChip: result.chip,
          workspace: result.path,
          notice: result.chip.warning || "",
          error: null,
        });
        storeFolder(result.path);
        await refreshList();
        await openBrowseDialog(result.path);
      }, [openBrowseDialog, patch, refreshList, reportError]);

      const openUploadedFile = React.useCallback(
        async (file) => {
          patch({
            saving: true,
            notice: `Opening ${file.name || "the document"}…`,
          });
          try {
            const result = await runOpenUploadFlow(file, {
              uiWorkspace: workspaceService(),
              useBoundServices: false,
            });
            await finishOpenedDocument(result);
          } catch (cause) {
            patch({ saving: false });
            reportError(cause);
          }
        },
        [finishOpenedDocument, patch, reportError],
      );

      const openFromBrowse = React.useCallback(
        async (path) => {
          // The folder that is actually opened is remembered for the next
          // "Open file…"/Ctrl+O, so the chooser reopens where the user was.
          const dialog = stateRef.current.dialog || {};
          const folder =
            (dialog.kind === "browse" && dialog.path) || folderPathOf(path);
          patch((previous) => ({
            saving: true,
            dialog:
              previous.dialog && previous.dialog.kind === "browse"
                ? Object.assign({}, previous.dialog, {
                    loading: true,
                    error: null,
                  })
                : previous.dialog,
          }));
          try {
            const result = await runOpenDocumentPathFlow(path, {
              uiWorkspace: workspaceService(),
              useBoundServices: false,
            });
            if (folder) storeFolder(folder);
            await finishOpenedDocument(result);
          } catch (cause) {
            patch((previous) => ({
              saving: false,
              advancedOpen: isMissingRoute(cause)
                ? true
                : previous.advancedOpen,
              dialog:
                previous.dialog && previous.dialog.kind === "browse"
                  ? Object.assign({}, previous.dialog, {
                      loading: false,
                      error:
                        (cause && cause.message) ||
                        "The document could not be opened by the Host.",
                    })
                  : previous.dialog,
            }));
          }
        },
        [finishOpenedDocument, patch],
      );

      const openCurrentFolder = React.useCallback(async () => {
        const dialog = stateRef.current.dialog || {};
        const target = dialog.path;
        if (!target) return;
        try {
          await prepareWorkspaceAt(target, false);
          storeFolder(target);
          patch({ dialog: null });
        } catch (cause) {
          patch((previous) => ({
            advancedOpen: isMissingRoute(cause) ? true : previous.advancedOpen,
            dialog:
              previous.dialog && previous.dialog.kind === "browse"
                ? Object.assign({}, previous.dialog, {
                    loading: false,
                    error:
                      (cause && cause.message) ||
                      "The workspace could not be switched to this folder.",
                  })
                : previous.dialog,
          }));
        }
      }, [patch, prepareWorkspaceAt]);

      const createFolder = React.useCallback(async () => {
        const dialog = stateRef.current.dialog || {};
        const name = String(dialog.newFolder || "")
          .trim()
          .replace(/[\\/]+/g, "");
        if (!name) {
          patch((previous) => ({
            dialog: Object.assign({}, previous.dialog, {
              error: "Enter a folder name first.",
            }),
          }));
          return;
        }
        const target = joinPath(dialog.path || "", name);
        try {
          await prepareWorkspaceAt(target, true);
          patch((previous) => ({
            dialog: previous.dialog
              ? Object.assign({}, previous.dialog, {
                  newFolder: null,
                  error: null,
                })
              : previous.dialog,
          }));
          await loadBrowse(target);
        } catch (cause) {
          patch((previous) => ({
            advancedOpen: isMissingRoute(cause) ? true : previous.advancedOpen,
            dialog:
              previous.dialog && previous.dialog.kind === "browse"
                ? Object.assign({}, previous.dialog, {
                    loading: false,
                    error:
                      (cause && cause.message) ||
                      "The folder could not be created in the workspace.",
                  })
                : previous.dialog,
          }));
        }
      }, [loadBrowse, patch, prepareWorkspaceAt]);

      const switchToChat = React.useCallback(() => {
        const chip = stateRef.current.sessionChip;
        const service = workspaceService();
        try {
          // Opening the prepared Session is a real navigation: do it only when
          // the user presses "Switch to chat".
          if (
            chip &&
            chip.sessionId &&
            service &&
            typeof service.openSession === "function"
          ) {
            service.openSession(chip.sessionId);
            return;
          }
          if (
            chip &&
            chip.dshWorkspaceId &&
            service &&
            typeof service.openWorkspace === "function"
          ) {
            void Promise.resolve(
              service.openWorkspace(chip.dshWorkspaceId),
            ).catch(() => undefined);
            return;
          }
        } catch {
          // Fall through to the layout switch below.
        }
        const result = switchToConversation({
          layout: layoutService(),
          useBoundServices: false,
        });
        if (!result.ok)
          patch({
            error: { code: "PANEL_UNAVAILABLE", message: result.message },
          });
      }, [patch]);

      const onFilePicked = React.useCallback(
        (event) => {
          const target = event && event.target;
          const file = target && target.files ? target.files[0] : null;
          // Reset so choosing the same file twice still fires onChange.
          if (target) target.value = "";
          if (!file) return;
          void openUploadedFile(file);
        },
        [openUploadedFile],
      );

      const onAdvancedPath = React.useCallback(
        (value) => patch({ advancedPath: value }),
        [patch],
      );

      const onToggleAdvanced = React.useCallback(
        (open) => {
          patch({ advancedOpen: open === true });
          if (open) void loadBrowseRoots();
        },
        [loadBrowseRoots, patch],
      );

      /* -------------------------------------------------------- actions */

      const copySelection = React.useCallback(async () => {
        const text = selectionText();
        if (!text) {
          reportError(HostError("RANGE_REQUIRED", REASONS.selectionRequired));
          return;
        }
        try {
          if (
            typeof navigator === "undefined" ||
            !navigator.clipboard ||
            typeof navigator.clipboard.writeText !== "function"
          )
            throw new Error("clipboard unavailable");
          await navigator.clipboard.writeText(text);
          patch({ notice: "Selection copied." });
        } catch (_cause) {
          reportError(
            HostError("CLIPBOARD_DENIED", REASONS.clipboardUnavailable),
          );
        }
      }, [patch, reportError, selectionText]);

      const cutSelection = React.useCallback(async () => {
        const text = selectionText();
        if (!text) {
          reportError(HostError("RANGE_REQUIRED", REASONS.selectionRequired));
          return;
        }
        try {
          if (
            typeof navigator === "undefined" ||
            !navigator.clipboard ||
            typeof navigator.clipboard.writeText !== "function"
          )
            throw new Error("clipboard unavailable");
          await navigator.clipboard.writeText(text);
        } catch (_cause) {
          reportError(
            HostError("CLIPBOARD_DENIED", REASONS.clipboardUnavailable),
          );
          return;
        }
        const current = stateRef.current;
        const selection = normalizeSelection(
          selectionRef.current,
          selectionOrderOf(current.bootstrap ? current.bootstrap.document : {}),
        );
        if (selection) await deleteSelectedRange(selection);
      }, [deleteSelectedRange, reportError, selectionText]);

      const pasteSelection = React.useCallback(async () => {
        try {
          if (
            typeof navigator === "undefined" ||
            !navigator.clipboard ||
            typeof navigator.clipboard.readText !== "function"
          )
            throw new Error("clipboard unavailable");
          const text = await navigator.clipboard.readText();
          if (!text) return;
          await insertAtCaret(text);
        } catch (_cause) {
          reportError(
            HostError("CLIPBOARD_DENIED", REASONS.clipboardUnavailable),
          );
        }
      }, [insertAtCaret, reportError]);

      const applyColor = React.useCallback(
        async (field, value) => {
          const props = {};
          if (field === "highlight") props.highlight = value || null;
          else props.color = value || null;
          await applyRunPropsToSelection(props);
        },
        [applyRunPropsToSelection],
      );

      const runHistory = React.useCallback(
        async (op) => {
          const current = stateRef.current;
          if (!current.bootstrap) return;
          beginStructural();
          try {
            await flushPending();
            const result = await apiClient.op(
              current.bootstrap.documentId,
              op,
              {},
            );
            if (result && result.document) {
              const document = result.document;
              rebuildTextCache(document);
              const stats = Object.assign(
                {},
                computeStats(document),
                result.stats || {},
              );
              caretRef.current = null;
              patch((previous) => ({
                bootstrap: previous.bootstrap
                  ? Object.assign({}, previous.bootstrap, {
                      document,
                      stats,
                      blockPages:
                        result.blockPages || previous.bootstrap.blockPages,
                      pageCount:
                        result.pageCount || previous.bootstrap.pageCount,
                      outline: computeOutline(document),
                      revision: result.revision || previous.bootstrap.revision,
                      dirty:
                        result.dirty === undefined
                          ? previous.bootstrap.dirty
                          : result.dirty === true,
                    })
                  : null,
                canUndo: op === "undo" ? true : previous.canUndo,
                canRedo: op === "redo" ? true : previous.canRedo,
                saveState: "unsaved",
              }));
              // The undo/redo document is authoritative: any local text still
              // pending would replay the very edit the history just reverted,
              // so it is dropped before the refetch.
              pendingTextRef.current.clear();
              await reload(selectionRef.current);
            } else {
              await reload(selectionRef.current);
            }
            applyEnvelope(result);
          } catch (cause) {
            if (op === "undo") patch({ canUndo: false });
            else patch({ canRedo: false });
            reportError(cause);
          } finally {
            await endStructural();
          }
        },
        [
          applyEnvelope,
          beginStructural,
          endStructural,
          flushPending,
          patch,
          rebuildTextCache,
          reload,
          reportError,
        ],
      );

      const applyStyleById = React.useCallback(
        async (styleId) => {
          const current = stateRef.current;
          if (!current.bootstrap) return;
          const selection = normalizeSelection(
            selectionRef.current,
            selectionOrderOf(current.bootstrap.document),
          );
          const { list } = collectBlocks(current.bootstrap.document);
          const blockIds = [];
          if (selection) {
            const startIndex = list.findIndex(
              (entry) => entry.block.id === selection.start.blockId,
            );
            const endIndex = list.findIndex(
              (entry) => entry.block.id === selection.end.blockId,
            );
            for (
              let index = startIndex;
              index <= Math.max(startIndex, endIndex) && index < list.length;
              index += 1
            ) {
              const entry = list[index];
              if (entry.block && entry.block.type === "paragraph")
                blockIds.push(entry.block.id);
            }
          }
          if (blockIds.length === 0) {
            reportError(HostError("NO_BLOCK", REASONS.noDocument));
            return;
          }
          await performOp(
            "applyStyle",
            { blockIds, styleId, clearDirectFormatting: true },
            { reload: true },
          );
        },
        [performOp, reportError],
      );

      const insertList = React.useCallback(
        async (kind) => {
          const caret = caretPosition();
          if (!caret) {
            reportError(HostError("NO_BLOCK", REASONS.noDocument));
            return;
          }
          const selected = selectionText();
          if (selected) patch({ notice: REASONS.listFromText });
          await performOp(
            "insertList",
            {
              items: [""],
              kind: kind === "bullet" ? "bullet" : "number",
              afterBlockId: caret.blockId,
              level: kind === "multilevel" ? 1 : 0,
            },
            { reload: true },
          );
        },
        [caretPosition, performOp, patch, reportError, selectionText],
      );

      const setIndent = React.useCallback(
        async (delta) => {
          const current = stateRef.current;
          const formatting = readFormatting(
            current.bootstrap ? current.bootstrap.document : null,
            selectionRef.current,
            selectionOrderOf(
              current.bootstrap ? current.bootstrap.document : {},
            ),
          );
          const base = formatting.indentLeft || 0;
          const next = Math.max(0, base + delta);
          await applyParagraphProps({ indent: { left: next } });
        },
        [applyParagraphProps],
      );

      const insertBreak = React.useCallback(
        async (kind) => {
          const caret = caretPosition();
          if (!caret) return;
          await performOp(
            "insertPageBreak",
            { position: caret, kind },
            { reload: true },
          );
        },
        [caretPosition, performOp],
      );

      const setPageLayout = React.useCallback(
        async (patchArgs, options) => {
          const current = stateRef.current;
          const document = current.bootstrap
            ? current.bootstrap.document
            : null;
          const sectionIndex =
            options && typeof options.sectionIndex === "number"
              ? options.sectionIndex
              : 0;
          await performOp(
            "setPageLayout",
            Object.assign({ sectionIndex }, patchArgs),
            { reload: true },
          );
          if (!document) return;
        },
        [performOp],
      );

      const runAction = React.useCallback(
        async (name, arg) => {
          const current = stateRef.current;
          // A new command clears the previous error bar; if this one fails it
          // reports again, and a failed op never clears its own message while
          // the panel refetches the document.
          if (name !== "dismissError" && current.error) patch({ error: null });
          const readOnly = isReadOnly(current);
          const guardEdit = () => {
            if (readOnly) {
              reportError(HostError("READ_ONLY", REASONS.readOnly));
              return false;
            }
            if (offlineRef.current) {
              reportError(HostError("OFFLINE", REASONS.offline));
              return false;
            }
            return true;
          };
          switch (name) {
            case "dismissError":
              patch({ error: null, notice: "", noticeDismissible: false });
              return;
            case "undo":
              await runHistory("undo");
              return;
            case "redo":
              await runHistory("redo");
              return;
            case "copy":
              await copySelection();
              return;
            case "cut":
              if (!guardEdit()) return;
              await cutSelection();
              return;
            case "paste":
              if (!guardEdit()) return;
              await pasteSelection();
              return;
            case "bold":
            case "italic":
            case "underline":
            case "strike":
            case "subscript":
            case "superscript":
            case "clearFormatting": {
              if (!guardEdit()) return;
              const formatting = readFormatting(
                current.bootstrap ? current.bootstrap.document : null,
                selectionRef.current,
                selectionOrderOf(
                  current.bootstrap ? current.bootstrap.document : {},
                ),
              );
              if (name === "clearFormatting") {
                await applyRunPropsToSelection({}, { clear: true });
                return;
              }
              if (name === "subscript" || name === "superscript") {
                const target =
                  formatting.vertAlign === name ? "baseline" : name;
                await applyRunPropsToSelection({ vertAlign: target });
                return;
              }
              const keyMap = {
                bold: "bold",
                italic: "italic",
                underline: "underline",
                strike: "strike",
              };
              const key = keyMap[name];
              await applyRunPropsToSelection({ [key]: !formatting[key] });
              return;
            }
            case "fontFamily":
              if (!guardEdit()) return;
              await applyRunPropsToSelection({ fontFamily: arg });
              return;
            case "fontSize":
              if (!guardEdit()) return;
              await applyRunPropsToSelection({ fontSize: Number(arg) });
              return;
            case "textColor":
              if (!guardEdit()) return;
              openDialog("color", {
                field: "color",
                value: current.contextColor || "#000000",
              });
              return;
            case "highlight":
              if (!guardEdit()) return;
              openDialog("color", { field: "highlight", value: "" });
              return;
            case "alignLeft":
            case "alignCenter":
            case "alignRight":
            case "alignJustify": {
              if (!guardEdit()) return;
              const alignment =
                name === "alignLeft"
                  ? "left"
                  : name === "alignCenter"
                    ? "center"
                    : name === "alignRight"
                      ? "right"
                      : "justify";
              await applyParagraphProps({ alignment });
              return;
            }
            case "increaseIndent":
              if (!guardEdit()) return;
              await setIndent(36);
              return;
            case "decreaseIndent":
              if (!guardEdit()) return;
              await setIndent(-36);
              return;
            case "lineSpacing":
              if (!guardEdit()) return;
              await applyParagraphProps({
                lineSpacing: { rule: "auto", value: Number(arg) * 240 },
              });
              return;
            case "spaceBefore":
              if (!guardEdit()) return;
              await applyParagraphProps({ spaceBefore: Number(arg) });
              return;
            case "spaceAfter":
              if (!guardEdit()) return;
              await applyParagraphProps({ spaceAfter: Number(arg) });
              return;
            case "bullets":
              if (!guardEdit()) return;
              await insertList("bullet");
              return;
            case "numbering":
              if (!guardEdit()) return;
              await insertList("number");
              return;
            case "multilevel":
              if (!guardEdit()) return;
              await insertList("multilevel");
              return;
            case "applyStyle":
              if (!guardEdit()) return;
              await applyStyleById(arg);
              return;
            case "find":
            case "replace":
              openDialog("find", {
                query: "",
                replacement: "",
                matches: [],
                searched: false,
                activeMatch: null,
              });
              return;
            case "selectAll":
              selectAllDocument();
              return;
            case "pageBreak":
              if (!guardEdit()) return;
              await insertBreak("page");
              return;
            case "columnBreak":
              if (!guardEdit()) return;
              await insertBreak("column");
              return;
            case "sectionBreak":
              if (!guardEdit()) return;
              await performOp(
                "insertSectionBreak",
                { sectionIndex: current.activeSection || 0 },
                { reload: true },
              );
              return;
            case "insertTable":
              if (!guardEdit()) return;
              openDialog("table", { rows: 3, columns: 3 });
              return;
            case "insertImage":
              if (!guardEdit()) return;
              openDialog("image", { dataBase64: "", contentType: "" });
              return;
            case "hyperlink":
              if (!guardEdit()) return;
              openDialog("link", { href: "", tooltip: "" });
              return;
            case "editHeader":
            case "editFooter": {
              if (!guardEdit()) return;
              const sections =
                (current.bootstrap &&
                  current.bootstrap.document &&
                  current.bootstrap.document.sections) ||
                [];
              const section =
                sections[current.activeSection || 0] || sections[0];
              openDialog("headerFooter", {
                part: name === "editHeader" ? "header" : "footer",
                text: headerFooterText(
                  section,
                  name === "editHeader" ? "header" : "footer",
                ),
              });
              return;
            }
            case "insertToc":
              if (!guardEdit()) return;
              openDialog("toc", { minLevel: 1, maxLevel: 3 });
              return;
            case "updateToc":
              await flushPending();
              await reload(selectionRef.current);
              patch({ notice: "Table of contents refetched from the Host." });
              return;
            case "insertFootnote":
            case "insertEndnote":
              if (!guardEdit()) return;
              openDialog("note", {
                noteType: name === "insertEndnote" ? "endnote" : "footnote",
                text: "",
              });
              return;
            case "insertEquation":
              if (!guardEdit()) return;
              openDialog("equation", { linear: "" });
              return;
            case "paragraphSpacingPreset": {
              if (!guardEdit()) return;
              const preset = SPACING_PRESETS.find((entry) => entry.id === arg);
              if (!preset) return;
              await applyParagraphProps({
                spaceBefore: preset.before,
                spaceAfter: preset.after,
              });
              return;
            }
            case "designFont": {
              if (!guardEdit()) return;
              const document = current.bootstrap
                ? current.bootstrap.document
                : null;
              const selection = document
                ? normalizeSelection(
                    selectionRef.current,
                    selectionOrderOf(document),
                  )
                : null;
              if (selection && !isCollapsed(selection)) {
                await applyRunPropsToSelection({ fontFamily: arg });
                return;
              }
              const caret = caretPosition();
              const byId = document ? collectBlocks(document).byId : new Map();
              const entry = caret ? byId.get(caret.blockId) : null;
              const length = entry ? paragraphText(entry.block).length : 0;
              if (!entry || length === 0) {
                reportError(
                  HostError("RANGE_REQUIRED", REASONS.selectionRequired),
                );
                return;
              }
              const previous = selectionRef.current;
              selectionRef.current = {
                start: { blockId: caret.blockId, offset: 0 },
                end: { blockId: caret.blockId, offset: length },
              };
              await applyRunPropsToSelection({ fontFamily: arg });
              selectionRef.current = previous;
              return;
            }
            case "margins": {
              if (!guardEdit()) return;
              if (arg === "custom") {
                const sections =
                  (current.bootstrap &&
                    current.bootstrap.document &&
                    current.bootstrap.document.sections) ||
                  [];
                const section = sections[0];
                const margins = (section &&
                  section.page &&
                  section.page.margins) || {
                  top: 72,
                  right: 72,
                  bottom: 72,
                  left: 72,
                };
                openDialog("margins", {
                  top: Math.round(margins.top),
                  right: Math.round(margins.right),
                  bottom: Math.round(margins.bottom),
                  left: Math.round(margins.left),
                });
                return;
              }
              const preset = MARGIN_PRESETS.find((entry) => entry.id === arg);
              if (!preset) return;
              const sections =
                (current.bootstrap &&
                  current.bootstrap.document &&
                  current.bootstrap.document.sections) ||
                [];
              const section =
                sections[current.activeSection || 0] || sections[0] || null;
              const existing =
                (section && section.page && section.page.margins) || {};
              // PageMargins is a complete record: keep header/footer/gutter so
              // the Host never receives a partial margin object.
              await setPageLayout(
                {
                  margins: {
                    top: preset.top,
                    right: preset.right,
                    bottom: preset.bottom,
                    left: preset.left,
                    header:
                      typeof existing.header === "number"
                        ? existing.header
                        : 36,
                    footer:
                      typeof existing.footer === "number"
                        ? existing.footer
                        : 36,
                    gutter:
                      typeof existing.gutter === "number" ? existing.gutter : 0,
                  },
                },
                { sectionIndex: current.activeSection || 0 },
              );
              return;
            }
            case "orientation":
              if (!guardEdit()) return;
              await setPageLayout(
                { orientation: arg === "landscape" ? "landscape" : "portrait" },
                { sectionIndex: current.activeSection || 0 },
              );
              return;
            case "pageSize":
              if (!guardEdit()) return;
              await setPageLayout(
                { pageSize: arg },
                { sectionIndex: current.activeSection || 0 },
              );
              return;
            case "columns":
              if (!guardEdit()) return;
              await performOp(
                "setColumns",
                {
                  sectionIndex: current.activeSection || 0,
                  count: Number(arg),
                  space: 36,
                },
                { reload: true },
              );
              return;
            case "toggleReadOnly": {
              const next = !readOnly;
              patch({
                readOnlyOverride: next,
                notice: next ? "Editing is disabled locally." : "",
              });
              return;
            }
            case "toggleFocus":
              patch((previous) => ({ focusMode: !previous.focusMode }));
              return;
            case "toggleRuler":
              patch((previous) => ({ showRuler: !previous.showRuler }));
              return;
            case "toggleLabels": {
              const next = !current.showLabels;
              storeLabels(next);
              patch({ showLabels: next });
              return;
            }
            case "toggleMarks":
              patch((previous) => ({ showMarks: !previous.showMarks }));
              return;
            case "togglePageSeparators":
              patch((previous) => ({
                showPageSeparators: !previous.showPageSeparators,
              }));
              return;
            case "toggleRail":
              patch((previous) => ({ showRail: !previous.showRail }));
              return;
            case "zoomIn":
              patch((previous) => ({
                zoom: Math.min(
                  2,
                  Math.round((previous.zoom + 0.1) * 100) / 100,
                ),
              }));
              return;
            case "zoomOut":
              patch((previous) => ({
                zoom: Math.max(
                  0.5,
                  Math.round((previous.zoom - 0.1) * 100) / 100,
                ),
              }));
              return;
            case "setZoom":
              patch({ zoom: Math.max(0.5, Math.min(2, Number(arg) || 1)) });
              return;
            case "fitWidth": {
              const surface = surfaceRef.current;
              const page =
                surface && surface.querySelector
                  ? surface.querySelector(".dx-page")
                  : null;
              if (surface && page) {
                const available = surface.clientWidth - 48;
                const currentZoom = current.zoom || 1;
                const pageWidth =
                  page.getBoundingClientRect().width / currentZoom;
                if (available > 0 && pageWidth > 0) {
                  patch({
                    zoom: Math.max(
                      0.5,
                      Math.min(
                        2,
                        Math.round((available / pageWidth) * 100) / 100,
                      ),
                    ),
                    fitWidth: true,
                  });
                }
              }
              return;
            }
            case "save":
              await saveDocument();
              return;
            case "saveAs":
              openDialog("saveAs", { path: "" });
              return;
            case "refresh":
              await flushPending();
              await refreshList();
              if (current.selectedId) await loadDocument(current.selectedId);
              return;
            case "deleteDocument":
              if (readOnly) {
                reportError(HostError("READ_ONLY", REASONS.readOnly));
                return;
              }
              openDialog("delete", {});
              return;
            case "newDocument":
              if (readOnly) {
                reportError(HostError("READ_ONLY", REASONS.readOnly));
                return;
              }
              openDialog("new", { title: "" });
              return;
            case "openByPath":
              if (readOnly) {
                reportError(HostError("READ_ONLY", REASONS.readOnly));
                return;
              }
              openDialog("open", { path: "" });
              return;
            case "export":
              await exportDocument(arg);
              return;
            case "chooseFolder":
              await chooseFolder();
              return;
            case "openFileChooser":
              await openFileChooser();
              return;
            case "openDocument":
              // Secondary path: the browser's own file input. When it cannot be
              // clicked (sandboxed/headless/insecure contexts) fall back to the
              // in-panel chooser instead of dead-ending on a notice.
              if (
                fileInputRef.current &&
                typeof fileInputRef.current.click === "function"
              ) {
                try {
                  fileInputRef.current.click();
                  return;
                } catch (_cause) {
                  /* fall through to the chooser */
                }
              }
              patch({
                notice:
                  "This browser cannot open its own file dialog, so the in-panel chooser opened instead.",
              });
              await openFileChooser();
              return;
            case "toggleAdvanced":
              onToggleAdvanced(!current.advancedOpen);
              return;
            case "loadBrowseRoots":
              await loadBrowseRoots();
              return;
            case "browseRoot":
            case "browsePath":
              await openBrowseDialog(arg);
              return;
            case "browseUp":
              await openBrowseDialog(
                (current.dialog && current.dialog.parent) || "",
              );
              return;
            case "openBrowseDocument":
              await openFromBrowse(arg);
              return;
            case "openCurrentFolder":
              await openCurrentFolder();
              return;
            case "newFolderStart":
              patch((previous) => ({
                dialog: previous.dialog
                  ? Object.assign({}, previous.dialog, {
                      newFolder: "",
                      error: null,
                    })
                  : previous.dialog,
              }));
              return;
            case "newFolderConfirm":
              await createFolder();
              return;
            case "openAdvancedPath": {
              const advanced = String(current.advancedPath || "").trim();
              if (!advanced) {
                patch({
                  notice: "Type a document path in “Advanced: path” first.",
                });
                return;
              }
              await openDocumentByPath(advanced);
              return;
            }
            case "switchToChat":
              switchToChat();
              return;
            case "dismissSessionChip":
              patch({ sessionChip: null });
              return;
            case "createDocument": {
              const dialog = current.dialog || {};
              await createDocument(dialog.title);
              return;
            }
            case "openDocumentByPath": {
              const dialog = current.dialog || {};
              await openDocumentByPath(dialog.path);
              return;
            }
            case "saveDocumentAs": {
              const dialog = current.dialog || {};
              await saveDocument(String(dialog.path || "").trim());
              return;
            }
            case "confirmDelete":
              await deleteCurrentDocument();
              return;
            case "runSearch":
              await runSearch();
              return;
            case "selectFindMatch":
              selectFindMatch(arg);
              return;
            case "replaceMatch":
              await replaceMatch();
              return;
            case "insertTableConfirm": {
              const dialog = current.dialog || {};
              const rows = Math.max(1, Math.min(200, Number(dialog.rows) || 1));
              const columns = Math.max(
                1,
                Math.min(40, Number(dialog.columns) || 1),
              );
              const caret = caretPosition();
              patch({ dialog: null });
              await performOp(
                "insertTable",
                {
                  rows,
                  columns,
                  afterBlockId: caret ? caret.blockId : undefined,
                  headerRow: true,
                },
                { reload: true },
              );
              return;
            }
            case "insertLinkConfirm": {
              const dialog = current.dialog || {};
              const href = sanitizeUrl(dialog.href);
              if (!href) {
                patch({
                  error: {
                    code: "INVALID_URL",
                    message:
                      "Only http, https and mailto destinations are accepted.",
                  },
                });
                return;
              }
              const range = rangeForSelection();
              if (!range || isCollapsed(range)) {
                patch({
                  error: {
                    code: "RANGE_REQUIRED",
                    message: "Select the text that should become the link.",
                  },
                });
                return;
              }
              patch({ dialog: null });
              await performOp(
                "insertHyperlink",
                { range, href, tooltip: dialog.tooltip || undefined },
                { reload: true },
              );
              return;
            }
            case "insertImageConfirm": {
              const dialog = current.dialog || {};
              if (!dialog.dataBase64) return;
              const caret = caretPosition();
              patch({ dialog: null });
              await performOp(
                "insertImage",
                {
                  position: caret,
                  dataBase64: dialog.dataBase64,
                  contentType: dialog.contentType,
                  alt: dialog.fileName || "Image",
                },
                { reload: true },
              );
              return;
            }
            case "insertTocConfirm": {
              const dialog = current.dialog || {};
              const caret = caretPosition();
              patch({ dialog: null });
              await performOp(
                "insertToc",
                {
                  afterBlockId: caret ? caret.blockId : undefined,
                  minLevel: Number(dialog.minLevel) || 1,
                  maxLevel: Number(dialog.maxLevel) || 3,
                  includeTitle: true,
                },
                { reload: true },
              );
              return;
            }
            case "insertNoteConfirm": {
              const dialog = current.dialog || {};
              const position = caretPosition();
              patch({ dialog: null });
              await performOp(
                "insertFootnote",
                {
                  position,
                  text: String(dialog.text || "").trim(),
                  type: dialog.noteType === "endnote" ? "endnote" : "footnote",
                },
                { reload: true },
              );
              return;
            }
            case "insertEquationConfirm": {
              const dialog = current.dialog || {};
              const position = caretPosition();
              patch({ dialog: null });
              await performOp(
                "insertEquation",
                {
                  position,
                  linear: String(dialog.linear || ""),
                  display: true,
                },
                { reload: true },
              );
              return;
            }
            case "editHeaderFooterConfirm": {
              const dialog = current.dialog || {};
              patch({ dialog: null });
              await performOp(
                "editHeaderFooter",
                {
                  sectionIndex: current.activeSection || 0,
                  variant: "default",
                  part: dialog.part === "footer" ? "footer" : "header",
                  text: String(dialog.text || ""),
                },
                { reload: true },
              );
              return;
            }
            case "marginsConfirm": {
              const dialog = current.dialog || {};
              const sections =
                (current.bootstrap &&
                  current.bootstrap.document &&
                  current.bootstrap.document.sections) ||
                [];
              const existing = (sections[0] &&
                sections[0].page &&
                sections[0].page.margins) || {
                header: 36,
                footer: 36,
                gutter: 0,
              };
              patch({ dialog: null });
              await setPageLayout(
                {
                  margins: {
                    top: Number(dialog.top) || 0,
                    right: Number(dialog.right) || 0,
                    bottom: Number(dialog.bottom) || 0,
                    left: Number(dialog.left) || 0,
                    header: existing.header || 36,
                    footer: existing.footer || 36,
                    gutter: existing.gutter || 0,
                  },
                },
                { sectionIndex: 0 },
              );
              return;
            }
            case "applyColorConfirm": {
              const dialog = current.dialog || {};
              patch({ dialog: null });
              await applyColor(dialog.field, dialog.value);
              return;
            }
            case "onOpenDocument":
              return;
            default:
              patch({
                error: {
                  code: "UNKNOWN_ACTION",
                  message: `The panel has no implementation for “${name}”.`,
                },
              });
          }
        },
        [
          applyColor,
          applyParagraphProps,
          applyRunPropsToSelection,
          applyStyleById,
          caretPosition,
          chooseFolder,
          closeDialog,
          copySelection,
          createDocument,
          createFolder,
          cutSelection,
          deleteCurrentDocument,
          deleteSelectedRange,
          exportDocument,
          insertAtCaret,
          insertBreak,
          insertList,
          isReadOnly,
          loadBrowseRoots,
          loadDocument,
          onToggleAdvanced,
          openBrowseDialog,
          openCurrentFolder,
          openDialog,
          openDocumentByPath,
          openFileChooser,
          openFromBrowse,
          pasteSelection,
          patch,
          performOp,
          placeCaret,
          rangeForSelection,
          reload,
          replaceMatch,
          reportError,
          runHistory,
          runSearch,
          saveDocument,
          selectAllDocument,
          selectFindMatch,
          selectionText,
          setIndent,
          setPageLayout,
          switchToChat,
          refreshList,
        ],
      );

      const onZoom = React.useCallback(
        (value) => {
          const zoom = Number(value);
          if (!Number.isFinite(zoom)) return;
          patch({
            zoom: Math.max(0.5, Math.min(2, Math.round(zoom * 100) / 100)),
            fitWidth: false,
          });
        },
        [patch],
      );

      const onTab = React.useCallback(
        (tabId) => {
          patch({ tab: tabId, ribbonCollapsed: false });
        },
        [patch],
      );

      const onToggleRibbon = React.useCallback(() => {
        patch((current) => ({ ribbonCollapsed: !current.ribbonCollapsed }));
      }, [patch]);

      const onOpenDocument = React.useCallback(
        (documentId) => {
          void loadDocument(documentId);
        },
        [loadDocument],
      );

      const onOpenUrl = React.useCallback((url) => {
        const safe = sanitizeUrl(url);
        if (!safe || typeof window === "undefined") return;
        window.open(safe, "_blank", "noopener,noreferrer");
      }, []);

      const onRootKeyDown = handleKeyDown;

      /**
       * Execute one table command. The capture listener, the `shortcuts`
       * service contribution and the React fallback all route through here, so
       * there is exactly one implementation per command.
       */
      const execShortcut = React.useCallback(
        async (id) => {
          // Table command ids are user-facing ("redoShift"); runAction names the
          // operation ("redo").
          const command =
            id === "redoShift" ? "redo" : id === "boldShift" ? "bold" : id;
          // A structural operation owns input order; do not run a formatting
          // command against a document that is about to be replaced.
          if (structuralPendingRef.current > 0) return;
          shortcutRuntime.executed = (shortcutRuntime.executed || 0) + 1;
          shortcutRuntime.lastExec = command;
          switch (command) {
            case "undo":
            case "redo":
            case "bold":
            case "italic":
            case "underline":
              await runAction(command);
              return;
            case "selectAll":
              selectAllDocument();
              return;
            case "save":
              await saveDocumentRef.current();
              return;
            case "find":
              openDialogRef.current("find", {
                query: "",
                replacement: "",
                matches: [],
                searched: false,
                activeMatch: null,
              });
              return;
            case "open":
              if (typeof openFileChooserRef.current === "function")
                await openFileChooserRef.current();
              return;
            case "link":
              await runAction("hyperlink");
              return;
            case "numberedList":
              await runAction("numbering");
              return;
            case "bulletList":
              await runAction("bullets");
              return;
            default:
              return;
          }
        },
        [runAction, selectAllDocument],
      );
      execShortcutRef.current = execShortcut;

      // The capture-phase shortcut listener is marked on the root by its
      // effect. The live selection diagnostic is written imperatively: the
      // editing surface is memoised, so a render-time attribute would only
      // update when the document changed, not when the selection did.
      React.useEffect(() => {
        const surface = surfaceRef.current;
        if (!surface) return;
        const state = selectionStateOf(
          stateRef.current.bootstrap ? stateRef.current.bootstrap.document : {},
          stateRef.current.selection,
        );
        if (state) {
          surface.setAttribute(
            "data-dx-selection",
            `${state.from.blockId}:${state.from.offset}>` +
              `${state.to.blockId}:${state.to.offset}`,
          );
          surface.setAttribute(
            "data-dx-selection-blocks",
            state.blocks.join(","),
          );
        } else {
          surface.removeAttribute("data-dx-selection");
          surface.removeAttribute("data-dx-selection-blocks");
        }
      }, [state.selection, state.bootstrap]);

      // Path 1: window capture listener. This is the only position that runs
      // before DSH's window-bubble dispatcher and before React's delegation.
      React.useEffect(() => {
        mountedRef.current = true;
        const unbind = bindShortcutCapture({
          getRoot: () => rootRef.current,
          getState: () => ({
            dialog: stateRef.current.dialog,
            composing: composingRef.current,
          }),
          onCommand: (entry) => {
            if (execShortcutRef.current) void execShortcutRef.current(entry.id);
          },
        });
        const root = rootRef.current;
        if (root) root.setAttribute("data-dx-shortcuts", "window-capture");
        return () => {
          mountedRef.current = false;
          unbind();
          if (root) root.removeAttribute("data-dx-shortcuts");
        };
      }, []);

      // The host may already own Ctrl+B. Detect it once at mount so the ribbon,
      // tooltip and note tell the truth before the user tries the chord.
      React.useEffect(() => {
        if (detectBoldReservation(clientService("shortcuts")))
          patch({ boldReserved: true });
      }, [patch]);

      // Path 2: the DSH `shortcuts` client service, when the page has one. On
      // Desktop Windows/macOS that service wins at the native `before-input-event`
      // layer, so it is the only route for a key the app has reserved. A
      // missing/older/conflicting registry must never break the panel.
      React.useEffect(() => {
        const disposers = registerShortcutCommands({
          shortcuts: clientService("shortcuts"),
          onCommand: (id) => {
            if (execShortcutRef.current) void execShortcutRef.current(id);
          },
        });
        return () => {
          for (const dispose of disposers) {
            try {
              dispose();
            } catch (_error) {
              /* a disposed or replaced registry must not break unmount */
            }
          }
        };
      }, []);

      // Selection survives a structural refetch: the model range is restored in
      // the DOM by block id + offset, not by node identity. A DOM selection that
      // already matches the model is left alone (no focus stealing, no IME
      // interruptions).
      React.useEffect(() => {
        const surface = surfaceRef.current;
        if (!surface || !selectionRef.current) return undefined;
        if (
          typeof document !== "undefined" &&
          document.activeElement &&
          !surface.contains(document.activeElement)
        )
          return undefined;
        const seq = placementSeqRef.current;
        const timer = setTimeout(() => {
          if (
            !surfaceRef.current ||
            dragSelectRef.current ||
            composingRef.current
          )
            return;
          const model = selectionRef.current;
          if (!model) return;
          const expected = normalizeSelection(
            model,
            selectionOrderOf(
              stateRef.current.bootstrap
                ? stateRef.current.bootstrap.document
                : {},
            ),
          );
          if (!expected) return;
          const dom = domSelectionToModel(surfaceRef.current);
          if (dom && sameSelection(dom, expected)) return;
          if (seq !== placementSeqRef.current) return;
          placeSelection(surfaceRef.current, expected);
        }, 0);
        return () => clearTimeout(timer);
      }, [state.bootstrap]);

      const api = React.useMemo(
        () => ({
          onAction: (name, arg) => {
            void Promise.resolve(runAction(name, arg)).then(() => {
              const current = stateRef.current;
              if (
                current.dialog ||
                !surfaceRef.current ||
                !selectionRef.current
              )
                return;
              if (typeof setTimeout === "function")
                setTimeout(() => {
                  if (surfaceRef.current && selectionRef.current)
                    placeSelection(surfaceRef.current, selectionRef.current);
                }, 0);
            });
          },
          onTitleChange,
          onTitleCommit,
          onTab,
          onToggleRibbon,
          onZoom,
          updateDialog,
          onDialogClose: closeDialog,
          onDialogSubmit: () => {
            const current = stateRef.current;
            const kind = current.dialog ? current.dialog.kind : "";
            if (kind === "find") {
              void runAction("runSearch");
              return;
            }
            void runAction(
              kind === "new"
                ? "createDocument"
                : kind === "open"
                  ? "openDocumentByPath"
                  : kind === "saveAs"
                    ? "saveDocumentAs"
                    : kind === "table"
                      ? "insertTableConfirm"
                      : kind === "link"
                        ? "insertLinkConfirm"
                        : kind === "image"
                          ? "insertImageConfirm"
                          : kind === "toc"
                            ? "insertTocConfirm"
                            : kind === "note"
                              ? "insertNoteConfirm"
                              : kind === "equation"
                                ? "insertEquationConfirm"
                                : kind === "headerFooter"
                                  ? "editHeaderFooterConfirm"
                                  : kind === "margins"
                                    ? "marginsConfirm"
                                    : kind === "color"
                                      ? "applyColorConfirm"
                                      : kind === "delete"
                                        ? "confirmDelete"
                                        : kind === "browse"
                                          ? "openCurrentFolder"
                                          : "",
            );
          },
          onImagePicked,
          onFilePicked,
          onAdvancedPath,
          onToggleAdvanced,
          onOpenDocument,
          onOpenUrl,
          onSelectionChange: updateSelection,
          onInput: handleInput,
          onPaste: handlePaste,
          onSurfaceMouseDown: handleSurfaceMouseDown,
          onSurfaceDoubleClick: handleSurfaceDoubleClick,
          onSurfaceCopy: handleSurfaceCopy,
          onSurfaceCut: handleSurfaceCut,
          onBeforeInput: handleBeforeInput,
          onCompositionStart: handleCompositionStart,
          onCompositionEnd: handleCompositionEnd,
          onBlur: () => {
            updateSelection();
            void flushPending();
          },
          onRootKeyDown,
          onRibbonHeight,
          rootRef,
          surfaceRef,
          fileInputRef,
          ribbonRef,
        }),
        [
          closeDialog,
          flushPending,
          handleCompositionEnd,
          handleCompositionStart,
          handleInput,
          handlePaste,
          handleBeforeInput,
          handleSurfaceCopy,
          handleSurfaceCut,
          handleSurfaceDoubleClick,
          handleSurfaceMouseDown,
          loadDocument,
          onAdvancedPath,
          onFilePicked,
          onImagePicked,
          onOpenDocument,
          onOpenUrl,
          onRibbonHeight,
          onTab,
          onTitleChange,
          onTitleCommit,
          onToggleAdvanced,
          onToggleRibbon,
          onZoom,
          onRootKeyDown,
          runAction,
          updateDialog,
          updateSelection,
        ],
      );

      React.useEffect(() => {
        let cancelled = false;
        (async () => {
          const payload = await refreshList();
          if (cancelled || !payload) return;
          const current = stateRef.current;
          if (current.bootstrap) return;
          const documents = payload.documents || [];
          if (documents.length > 0) await loadDocument(documents[0].id);
        })();
        return () => {
          cancelled = true;
        };
      }, [loadDocument, refreshList]);

      React.useEffect(() => {
        const seq = placementSeqRef.current;
        const caret = caretRef.current;
        caretRef.current = null;
        if (!caret || !surfaceRef.current) return undefined;
        const timer = setTimeout(() => {
          if (seq !== placementSeqRef.current) return;
          if (surfaceRef.current) placeSelection(surfaceRef.current, caret);
        }, 0);
        return () => clearTimeout(timer);
      }, [state.bootstrap]);

      React.useEffect(() => {
        return () => {
          if (flushTimerRef.current) clearTimeout(flushTimerRef.current);
          if (titleTimerRef.current) clearTimeout(titleTimerRef.current);
        };
      }, []);

      /**
       * Keep the File backstage glued to the ribbon's real bottom edge. The
       * offset is published as a CSS variable on the panel root, so the
       * absolutely-positioned overlay follows every drag without re-rendering
       * the document, and the hard-coded top:72px is gone.
       */
      React.useEffect(() => {
        const root = rootRef.current;
        const ribbon = ribbonRef.current;
        if (!root || !ribbon) return undefined;
        const measure = () => {
          const ribbonRect = ribbon.getBoundingClientRect();
          const rootRect = root.getBoundingClientRect();
          if (!ribbonRect.height && ribbon.hidden) {
            // A collapsed ribbon has no bottom edge; the overlay is not shown.
            root.style.removeProperty("--dx-ribbon-bottom");
            return;
          }
          const offset = Math.max(0, ribbonRect.bottom - rootRect.top);
          root.style.setProperty(
            "--dx-ribbon-bottom",
            `${Math.round(offset * 100) / 100}px`,
          );
        };
        measure();
        let observer = null;
        if (typeof ResizeObserver !== "undefined") {
          observer = new ResizeObserver(measure);
          observer.observe(ribbon);
        }
        if (typeof window !== "undefined" && window.addEventListener)
          window.addEventListener("resize", measure);
        return () => {
          if (observer) observer.disconnect();
          if (typeof window !== "undefined" && window.removeEventListener)
            window.removeEventListener("resize", measure);
        };
      }, [state.tab, state.ribbonCollapsed, state.ribbonHeight]);

      /** Clamp a height restored from `localStorage` to this viewport. */
      React.useEffect(() => {
        const stored = stateRef.current.ribbonHeight;
        if (typeof stored !== "number") return;
        const clamped = clampRibbonHeight(stored, viewportHeight());
        if (clamped === null || clamped === stored) return;
        patch({ ribbonHeight: clamped });
        storeRibbonHeight(clamped);
      }, [patch]);

      return h(DocumentsPanelView, buildViewProps(state, api));
    }

    const __test = {
      TABS,
      TAB_LABELS: TABS.map((tab) => tab.label),
      REASONS,
      FONT_SIZES,
      ZOOM_STEPS,
      STYLE_GALLERY,
      PAGE_SIZES,
      MARGIN_PRESETS,
      LINE_SPACING,
      SPACING_PRESETS,
      inlineText,
      finalInlineText,
      paragraphText,
      finalParagraphText,
      flatInlines,
      spanAtOffset,
      runPropsAtOffset,
      clampOffset,
      diffTextRange,
      applyTextReplace,
      mergeTextRuns,
      findStyle,
      styleChain,
      resolveRunProps,
      resolveParagraphProps,
      runStyleToCss,
      paragraphStyleToCss,
      sanitizeUrl,
      mediaUrl,
      computeStats,
      computeOutline,
      groupBlocksIntoPages,
      computeListLabels,
      documentBodyBlocks,
      collectBlocks,
      fontFamiliesFor,
      styleGalleryFor,
      pageSizeLabel,
      marginLabel,
      defaultPageMap,
      normalizeBootstrap,
      seedState,
      buildViewProps,
      readFormatting,
      normalizeSelection,
      isCollapsed,
      sameSelection,
      blockRangesForSelection,
      formatTextOpsForSelection,
      wordRangeAt,
      paragraphFlow,
      moveSelectionFocus,
      selectionForKeyboardExtension,
      selectionStateOf,
      selectionTextOf,
      selectionPointAt,
      domRangeForPoints,
      selectDomRange,
      selectionDiagnostics,
      planTextFlush,
      applyTextDiff,
      SHORTCUT_TABLE,
      SHORTCUT_COMMAND_PREFIX,
      BOLD_RESERVED_NOTE,
      detectBoldReservation,
      shortcutCombo,
      shortcutForCombo,
      shortcutEnvironment,
      shortcutCommandDefaults,
      matchShortcut,
      modalOutsidePanel,
      bindShortcutCapture,
      shortcutDiagnostics,
      registerShortcutCommands,
      newBlockIdFromOpResult,
      domTextOfBlock,
      domSelectionToModel,
      eventIsComposing,
      textPointAtOffset,
      offsetInBlockElement,
      blockElementFor,
      initialState,
      makeNoopApi,
      renderForTest,
      DocumentsPanel,
      DocumentsPanelView,
      ribbonGroups,
      REQUEST_HEADER: HEADER,
      API_PREFIX: API,
      LABELS_KEY,
      RIBBON_HEIGHT_KEY,
      LAST_FOLDER_KEY,
      RIBBON_MIN_HEIGHT,
      RIBBON_DEFAULT_MIN,
      RIBBON_DEFAULT_MAX,
      RIBBON_KEYBOARD_STEP,
      viewportHeight,
      ribbonMaxHeight,
      defaultRibbonHeight,
      clampRibbonHeight,
      readStoredRibbonHeight,
      storeRibbonHeight,
      readStoredFolder,
      storeFolder,
      folderPathOf,
      Icon,
      RibbonResizer,
      ribbonStyle,
      SHORTCUTS,
      DOCUMENT_ACCEPT,
      readStoredLabels,
      storeLabels,
      baseNameOf,
      joinPath,
      trimTrailingSeparator,
      breadcrumbSegments,
      formatBytes,
      browseDialogState,
      describeBrowseFailure,
      normalizeBrowseEntry,
      prepareWorkspaceSession,
      sessionChipText,
      switchToConversation,
      runChooseFolderFlow,
      runOpenDocumentPathFlow,
      runOpenUploadFlow,
      PREPARED_SESSIONS,
      clearPreparedSessions,
      bindClientContext,
      clientService,
      workspaceService,
      layoutService,
      canPickDirectory,
      canStartSession,
      SessionChip,
      BrowseDialog,
      FileRibbon,
      FileActions,
      RibbonTool,
      ribbonToolTitle,
    };

    function DocIcon(props) {
      const size = typeof props.size === "number" ? props.size : 18;
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
          focusable: "false",
          style: {
            display: "block",
            opacity: props.active === true ? 1 : 0.86,
          },
        },
        h("path", { d: "M6.5 3.5h7.2l4.3 4.3v12.7H6.5z" }),
        h("path", { d: "M13.7 3.5v4.3h4.3" }),
        h("path", { d: "M9.2 12h6.1M9.2 15h6.1M9.2 18h3.6" }),
      );
    }

    return {
      inject: ["slots"],
      apply(ctx) {
        // Keep the Context so the panel can reach the optional Client
        // services (`uiWorkspace`, `layout`) without requiring them.
        bindClientContext(ctx);
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
      __test,
    };
  },
});
