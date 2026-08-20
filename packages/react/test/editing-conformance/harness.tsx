/**
 * Interactive-editing conformance harness.
 *
 * Mounts the REAL editor the way the products mount it and drives it through
 * the GESTURE path — real KeyboardEvent / ClipboardEvent dispatch against the
 * editor's own listeners — never through DocxViewApi mutation calls. The api
 * object is used only for caret PLACEMENT (`find`, the pattern existing tests
 * use) and read-only inspection (`save`, `document`).
 *
 * Two mounts:
 *  - "session": DocxView + LocalDocumentSession via useAgentDocumentSession —
 *    byte-for-byte the likeoffice desktop mount (App.tsx), and the collab-gated
 *    editor path (doc.stableIds set, intents emitted).
 *  - "local": DocxView editable with no collab — the solo engine path, where
 *    the local EditHistory serves undo/redo.
 */
import { expect } from "vitest";
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { DocxView, useAgentDocumentSession, type DocxViewApi } from "../../src/index.js";
import { LocalDocumentSession, localDocumentViewBinding } from "../../../agent/src/index.js";
import { DocxDocument, serializeXml, type XmlElement } from "@wordinweb/core";
import { localName } from "@wordinweb/core";

export type Mode = "session" | "local";

async function tick(ms = 2): Promise<void> {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, ms));
  });
}

export interface Mods {
  meta?: boolean;
  ctrl?: boolean;
  shift?: boolean;
  alt?: boolean;
}

export interface Editor {
  mode: Mode;
  api: DocxViewApi;
  container: HTMLElement;
  doc: () => DocxDocument;
  /** Canonical serialized body — the comparison currency of the suite. */
  xml: () => string;
  /** Concatenated visible text (w:t / w:delText excluded), "\n" between paragraphs. */
  text: () => string;
  press: (key: string, mods?: Mods) => Promise<void>;
  type: (s: string) => Promise<void>;
  /** Select the first match of `q` (placement helper; the pattern of the
   * existing keyboard tests). Throws when the text is absent. */
  select: (q: string) => Promise<void>;
  /** Collapse a `select` to the match's start/end via the arrow-key gesture. */
  caretAt: (q: string, edge?: "start" | "end") => Promise<void>;
  copy: () => Promise<void>;
  cut: () => Promise<void>;
  /** Paste the virtual clipboard, or an explicit payload. */
  paste: (payload?: { "text/plain"?: string; "text/html"?: string }) => Promise<void>;
  clipboard: Map<string, string>;
  undo: () => Promise<void>;
  redo: () => Promise<void>;
  /** Errors thrown inside event listeners (jsdom swallows them off the
   * dispatch path); drained by the caller. */
  errors: unknown[];
  save: () => Uint8Array;
  unmount: () => Promise<void>;
}

/** Serialize without pretty-printing so comparisons are byte-honest. */
export function bodyXml(doc: DocxDocument): string {
  return serializeXml(doc.docRoot);
}

export function docText(doc: DocxDocument): string {
  const paras: string[] = [];
  const walkP = (el: XmlElement): string => {
    let out = "";
    for (const c of el.children) {
      const n = localName(c.name);
      if (n === "t") out += c.text;
      else if (n === "tab") out += "\t";
      else if (n === "br" || n === "cr") out += "\v";
      else if (n === "delText") out += ""; // deleted-under-track text is not visible content
      else out += walkP(c);
    }
    return out;
  };
  const walk = (el: XmlElement): void => {
    if (localName(el.name) === "p") {
      paras.push(walkP(el));
      return;
    }
    for (const c of el.children) walk(c);
  };
  walk(doc.docRoot);
  return paras.join("\n");
}

export async function mountEditor(source: Uint8Array, mode: Mode): Promise<Editor> {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root: Root = createRoot(container);
  const seen: { api: DocxViewApi | null } = { api: null };
  const errors: unknown[] = [];
  const onWindowError = (e: ErrorEvent): void => {
    errors.push(e.error ?? e.message);
    e.preventDefault();
  };
  window.addEventListener("error", onWindowError);

  let session: LocalDocumentSession | null = null;
  function SessionHost(): ReturnType<typeof createElement> {
    const view = useAgentDocumentSession(localDocumentViewBinding(session!));
    return createElement(DocxView, {
      ...view,
      editable: true,
      onReady: (a: DocxViewApi) => {
        seen.api = a;
      },
    });
  }

  await act(async () => {
    if (mode === "session") {
      session = new LocalDocumentSession(source);
      root.render(createElement(SessionHost));
    } else {
      root.render(
        createElement(DocxView, {
          source,
          editable: true,
          onReady: (a: DocxViewApi) => {
            seen.api = a;
          },
        }),
      );
    }
  });
  for (let i = 0; i < 60 && (!seen.api || !container.querySelector(".dxw-page")); i++) await tick(5);
  expect(container.querySelector(".dxw-page"), "editor page painted").toBeTruthy();
  expect(seen.api, "onReady fired").toBeTruthy();
  const api = seen.api!;

  const doc = (): DocxDocument => api.document;

  // Events must reach the editor container's own listeners; dispatch on the
  // focused element when it is inside, else the IME textarea (both bubble).
  const target = (): HTMLElement =>
    (container.contains(document.activeElement)
      ? (document.activeElement as HTMLElement)
      : container.querySelector("textarea")) ?? container;

  const press = async (key: string, mods: Mods = {}): Promise<void> => {
    await act(async () => {
      target().dispatchEvent(
        new KeyboardEvent("keydown", {
          key,
          metaKey: !!mods.meta,
          ctrlKey: !!mods.ctrl,
          shiftKey: !!mods.shift,
          altKey: !!mods.alt,
          bubbles: true,
          cancelable: true,
        }),
      );
      await new Promise((r) => setTimeout(r, 1));
    });
  };

  const type = async (s: string): Promise<void> => {
    for (const ch of s) await press(ch);
  };

  const select = async (q: string): Promise<void> => {
    let n = 0;
    await act(async () => {
      n = api.find(q);
    });
    if (n < 1) throw new Error(`fixture landmark not found: ${JSON.stringify(q)}`);
    await tick();
  };

  const caretAt = async (q: string, edge: "start" | "end" = "end"): Promise<void> => {
    await select(q);
    await press(edge === "start" ? "ArrowLeft" : "ArrowRight");
  };

  const clipboard = new Map<string, string>();
  const clipboardEvent = (kind: "copy" | "cut" | "paste", data: Map<string, string>): Event => {
    const ev = new Event(kind, { bubbles: true, cancelable: true });
    Object.defineProperty(ev, "clipboardData", {
      value: {
        getData: (type: string) => data.get(type) ?? "",
        setData: (type: string, value: string) => {
          data.set(type, value);
        },
      },
    });
    return ev;
  };

  const copy = async (): Promise<void> => {
    clipboard.clear();
    await act(async () => {
      target().dispatchEvent(clipboardEvent("copy", clipboard));
      await new Promise((r) => setTimeout(r, 1));
    });
  };
  const cut = async (): Promise<void> => {
    clipboard.clear();
    await act(async () => {
      target().dispatchEvent(clipboardEvent("cut", clipboard));
      await new Promise((r) => setTimeout(r, 1));
    });
  };
  const paste = async (payload?: { "text/plain"?: string; "text/html"?: string }): Promise<void> => {
    const data = payload
      ? new Map(Object.entries(payload).filter(([, v]) => v !== undefined) as [string, string][])
      : clipboard;
    await act(async () => {
      target().dispatchEvent(clipboardEvent("paste", data));
      await new Promise((r) => setTimeout(r, 1));
    });
  };

  return {
    mode,
    api,
    container,
    doc,
    xml: () => bodyXml(doc()),
    text: () => docText(doc()),
    press,
    type,
    select,
    caretAt,
    copy,
    cut,
    paste,
    clipboard,
    undo: () => press("z", { meta: true }),
    redo: () => press("z", { meta: true, shift: true }),
    errors,
    save: () => api.save(),
    unmount: async () => {
      window.removeEventListener("error", onWindowError);
      await act(async () => {
        root.unmount();
      });
      container.remove();
    },
  };
}

/**
 * STRUCTURAL LINT — invariants the model itself relies on. A violation means
 * an editing gesture stranded structure that no load path would produce.
 */
export function structuralLint(doc: DocxDocument): string[] {
  const problems: string[] = [];
  const CONTENT = new Set([
    "t", "delText", "tab", "br", "cr", "drawing", "pict", "object", "sym",
    "fldChar", "instrText", "footnoteReference", "endnoteReference",
    "commentReference", "noBreakHyphen", "softHyphen", "ruby",
  ]);
  const runHasContent = (r: XmlElement): boolean =>
    r.children.some((c) => CONTENT.has(localName(c.name)));

  const walk = (el: XmlElement, inP: boolean): void => {
    const name = localName(el.name);
    if (name === "p") {
      const pPrs = el.children.filter((c) => localName(c.name) === "pPr");
      if (pPrs.length > 1) problems.push("paragraph with more than one pPr");
      if (pPrs.length === 1 && el.children.indexOf(pPrs[0]) !== 0)
        problems.push("pPr is not the paragraph's first child");
      // Zero-length w:t is the blank placeholder ONLY when it is the sole run
      // content of its paragraph.
      const runs = el.children.filter((c) => localName(c.name) === "r");
      for (const r of runs) {
        if (r.children.length === 0) problems.push("completely empty w:r");
        else if (!runHasContent(r)) problems.push("w:r holding only properties (stranded shell)");
      }
      const ts = runs.flatMap((r) => r.children.filter((c) => localName(c.name) === "t"));
      const emptyTs = ts.filter((t) => t.text.length === 0);
      if (emptyTs.length > 0 && (ts.length > 1 || runs.length > 1)) {
        // Exemption: a zero-length w:t ADJACENT to an inline separator
        // (br/tab/cr) in the paragraph's content order is the caret home a
        // soft line break / tab needs (Word's files just end the run with
        // the bare separator; our caret model keeps a placeholder to type
        // into). Everything else is a stranded run no load path produces.
        const atoms: (XmlElement | "sep")[] = [];
        const collectAtoms = (e: XmlElement): void => {
          for (const c of e.children) {
            const n = localName(c.name);
            if (n === "pPr" || n === "rPr") continue;
            if (n === "t") atoms.push(c);
            else if (n === "br" || n === "tab" || n === "cr") atoms.push("sep");
            else collectAtoms(c);
          }
        };
        collectAtoms(el);
        const exempt = (t: XmlElement): boolean => {
          const i = atoms.indexOf(t);
          return i !== -1 && (atoms[i - 1] === "sep" || atoms[i + 1] === "sep");
        };
        if (emptyTs.some((t) => !exempt(t)))
          problems.push("zero-length w:t stranded beside other content");
      }
    }
    if (name === "ins" || name === "del") {
      if (el.children.length === 0) problems.push(`empty tracked-change marker w:${name}`);
    }
    if (name === "tc") {
      if (!el.children.some((c) => localName(c.name) === "p" || localName(c.name) === "tbl"))
        problems.push("table cell without a block (invalid OOXML)");
    }
    if (name === "tr") {
      if (!el.children.some((c) => localName(c.name) === "tc"))
        problems.push("table row without cells");
    }
    for (const c of el.children) walk(c, inP || name === "p");
  };
  walk(doc.docRoot, false);
  return problems;
}

/** Save → reparse → compare: the document a gesture leaves behind must write
 * a .docx that loads and yields the same body. Returns problems. */
export function saveReloadProblems(ed: Editor): string[] {
  let bytes: Uint8Array;
  try {
    bytes = ed.save();
  } catch (e) {
    return [`save() threw: ${String(e)}`];
  }
  let reloaded: DocxDocument;
  try {
    reloaded = DocxDocument.load(bytes);
  } catch (e) {
    return [`saved bytes do not reload: ${String(e)}`];
  }
  const before = docText(ed.doc());
  const after = docText(reloaded);
  if (before !== after) return [`text changed across save/reload: ${JSON.stringify(before)} -> ${JSON.stringify(after)}`];
  return [];
}
