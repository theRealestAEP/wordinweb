import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName } from "../xml.js";
import { SelectionSegment, applyRunFormat } from "./commands.js";

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

/**
 * Create a new review comment on the selected range, like Word/Google Docs:
 * commentRangeStart/End markers around the selection, a commentReference run
 * after the range, and the comment body in comments.xml (the part is created
 * on demand for documents that have no comments yet). Callers checkpoint
 * history and rerender afterwards.
 */
export function addComment(
  doc: DocxDocument,
  segments: SelectionSegment[],
  text: string,
  author: string,
  initials?: string,
): boolean {
  if (!text.trim() || segments.length === 0) return false;
  const commentsRoot = doc.commentsTree(true);
  if (!commentsRoot) return false;
  const w = commentsRoot.name.includes(":") ? commentsRoot.name.slice(0, commentsRoot.name.indexOf(":") + 1) : "";

  // Split partially covered runs so the range lands on run boundaries; the
  // returned ranges reference the fresh post-split w:t elements in order.
  const ranges = applyRunFormat(doc, segments, {});
  if (ranges.length === 0) return false;
  const firstT = ranges[0].t;
  const lastT = ranges[ranges.length - 1].t;
  const firstR = doc.findParentOf(firstT);
  const lastR = doc.findParentOf(lastT);
  const firstP = firstR && doc.findParentOf(firstR);
  const lastP = lastR && doc.findParentOf(lastR);
  if (!firstR || !lastR || !firstP || !lastP) return false;

  // Allocate a fresh numeric id.
  let idNum = 0;
  for (const c of doc.comments) {
    const n = parseInt(c.id, 10);
    if (Number.isFinite(n)) idNum = Math.max(idNum, n);
  }
  const newId = String(idNum + 1);

  const startIdx = firstP.children.indexOf(firstR);
  firstP.children.splice(startIdx, 0, el(`${w}commentRangeStart`, { [`${w}id`]: newId }));
  const endIdx = lastP.children.indexOf(lastR);
  lastP.children.splice(
    endIdx + 1,
    0,
    el(`${w}commentRangeEnd`, { [`${w}id`]: newId }),
    el(`${w}r`, {}, [
      el(`${w}rPr`, {}, [el(`${w}rStyle`, { [`${w}val`]: "CommentReference" })]),
      el(`${w}commentReference`, { [`${w}id`]: newId }),
    ]),
  );

  // Comment body; the w14:paraId lets replies thread to it later.
  const usedParaIds = new Set(doc.comments.map((c) => c.paraId).filter(Boolean));
  let paraId = "";
  do {
    paraId = Math.floor(Math.random() * 0xfffffff0 + 1)
      .toString(16)
      .toUpperCase()
      .padStart(8, "0");
  } while (usedParaIds.has(paraId));
  commentsRoot.children.push(
    el(
      `${w}comment`,
      {
        [`${w}id`]: newId,
        [`${w}author`]: author,
        ...(initials ? { [`${w}initials`]: initials } : {}),
        [`${w}date`]: new Date().toISOString(),
      },
      [
        el(
          `${w}p`,
          {
            "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
            "w14:paraId": paraId,
          },
          [el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" }, [], text)])],
        ),
      ],
    ),
  );

  doc.markCommentsChanged();
  doc.refresh();
  return true;
}

/**
 * Delete a review comment like Word: remove its entry from comments.xml and
 * strip the commentRangeStart/End markers and commentReference run from the
 * document. Deleting a parent removes its whole reply thread. Callers
 * checkpoint history and rerender afterwards.
 */
export function deleteComment(doc: DocxDocument, id: string): boolean {
  // Cascade over the reply thread (children first, any depth).
  const ids = [id];
  for (let i = 0; i < ids.length; i++) {
    for (const c of doc.comments) {
      if (c.parentId === ids[i] && !ids.includes(c.id)) ids.push(c.id);
    }
  }

  let touched = false;
  const commentsRoot = doc.commentsTree();
  const paraIds = new Set(
    doc.comments.filter((c) => ids.includes(c.id) && c.paraId).map((c) => c.paraId!),
  );
  if (commentsRoot) {
    const before = commentsRoot.children.length;
    commentsRoot.children = commentsRoot.children.filter(
      (c) => !(localName(c.name) === "comment" && ids.includes(attr(c, "id") ?? "")),
    );
    touched = commentsRoot.children.length !== before;
  }

  // Drop the thread's commentsExtended entries too.
  const extRoot = doc.commentsExtendedTree();
  if (extRoot && paraIds.size > 0) {
    const before = extRoot.children.length;
    extRoot.children = extRoot.children.filter(
      (c) => !(localName(c.name) === "commentEx" && paraIds.has(attr(c, "paraId") ?? "")),
    );
    if (extRoot.children.length !== before) {
      doc.markCommentsExtendedChanged();
      touched = true;
    }
  }

  // Strip markers from every editable tree (comments can sit in headers too).
  for (const root of doc.editableRoots()) {
    if (root === commentsRoot) continue;
    for (const oneId of ids) {
      touched = stripMarkers(root, oneId) || touched;
    }
  }

  if (touched) {
    doc.markCommentsChanged();
    doc.refresh();
  }
  return touched;
}

/**
 * Reply to a comment like Word: a new w:comment anchored to the parent's
 * range, threaded to the parent via commentsExtended (paraIdParent).
 */
export function replyToComment(
  doc: DocxDocument,
  parentId: string,
  text: string,
  author: string,
  initials?: string,
): boolean {
  const commentsRoot = doc.commentsTree();
  const parent = doc.comments.find((c) => c.id === parentId);
  if (!commentsRoot || !parent || !text.trim()) return false;

  const w = commentsRoot.name.includes(":") ? commentsRoot.name.slice(0, commentsRoot.name.indexOf(":") + 1) : "";
  const usedIds = new Set(doc.comments.map((c) => c.id));
  let idNum = 0;
  for (const c of doc.comments) {
    const n = parseInt(c.id, 10);
    if (Number.isFinite(n)) idNum = Math.max(idNum, n);
  }
  let newId = String(idNum + 1);
  while (usedIds.has(newId)) newId = String(++idNum + 1);

  const usedParaIds = new Set(doc.comments.map((c) => c.paraId).filter(Boolean));
  const freshParaId = (): string => {
    for (;;) {
      const pid = Math.floor(Math.random() * 0xfffffff0 + 1)
        .toString(16)
        .toUpperCase()
        .padStart(8, "0");
      if (!usedParaIds.has(pid)) {
        usedParaIds.add(pid);
        return pid;
      }
    }
  };

  // Threading needs the parent's body paragraph to carry a w14:paraId.
  let parentParaId = parent.paraId;
  if (!parentParaId) {
    const parentEl = commentsRoot.children.find(
      (c) => localName(c.name) === "comment" && attr(c, "id") === parentId,
    );
    let lastP: XmlElement | undefined;
    const findP = (e: XmlElement): void => {
      if (localName(e.name) === "p") {
        lastP = e;
        return;
      }
      for (const ch of e.children) findP(ch);
    };
    if (parentEl) for (const ch of parentEl.children) findP(ch);
    if (!lastP) return false;
    parentParaId = freshParaId();
    lastP.attrs["xmlns:w14"] = "http://schemas.microsoft.com/office/word/2010/wordml";
    lastP.attrs["w14:paraId"] = parentParaId;
  }

  // The reply's comment body.
  const replyParaId = freshParaId();
  commentsRoot.children.push(
    el(
      `${w}comment`,
      {
        [`${w}id`]: newId,
        [`${w}author`]: author,
        ...(initials ? { [`${w}initials`]: initials } : {}),
        [`${w}date`]: new Date().toISOString(),
      },
      [
        el(
          `${w}p`,
          {
            "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
            "w14:paraId": replyParaId,
          },
          [el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" }, [], text)])],
        ),
      ],
    ),
  );

  // Anchor the reply to the parent's range: markers right beside the
  // parent's, reference run after the parent's reference run.
  for (const root of doc.editableRoots()) {
    if (root === commentsRoot) continue;
    insertReplyMarkers(root, parentId, newId, w);
  }

  // Thread it.
  const extRoot = doc.commentsExtendedTree(true);
  if (extRoot) {
    const hasParentEx = extRoot.children.some(
      (c) => localName(c.name) === "commentEx" && attr(c, "paraId") === parentParaId,
    );
    if (!hasParentEx) {
      extRoot.children.push(
        el("w15:commentEx", { "w15:paraId": parentParaId, "w15:done": "0" }),
      );
    }
    extRoot.children.push(
      el("w15:commentEx", {
        "w15:paraId": replyParaId,
        "w15:paraIdParent": parentParaId,
        "w15:done": "0",
      }),
    );
    doc.markCommentsExtendedChanged();
  }

  doc.markCommentsChanged();
  doc.refresh();
  return true;
}

/** Insert the reply's range markers/reference adjacent to the parent's. */
function insertReplyMarkers(el0: XmlElement, parentId: string, newId: string, w: string): boolean {
  let touched = false;
  for (let i = 0; i < el0.children.length; i++) {
    const c = el0.children[i];
    const ln = localName(c.name);
    if (ln === "commentRangeStart" && attr(c, "id") === parentId) {
      el0.children.splice(i + 1, 0, el(`${w}commentRangeStart`, { [`${w}id`]: newId }));
      touched = true;
      i++;
      continue;
    }
    if (ln === "commentRangeEnd" && attr(c, "id") === parentId) {
      el0.children.splice(i + 1, 0, el(`${w}commentRangeEnd`, { [`${w}id`]: newId }));
      touched = true;
      i++;
      continue;
    }
    if (
      ln === "r" &&
      c.children.some((rc) => localName(rc.name) === "commentReference" && attr(rc, "id") === parentId)
    ) {
      el0.children.splice(i + 1, 0, el(`${w}r`, {}, [el(`${w}commentReference`, { [`${w}id`]: newId })]));
      touched = true;
      i++;
      continue;
    }
    touched = insertReplyMarkers(c, parentId, newId, w) || touched;
  }
  return touched;
}

function stripMarkers(el: XmlElement, id: string): boolean {
  let touched = false;
  el.children = el.children.filter((c) => {
    const ln = localName(c.name);
    if ((ln === "commentRangeStart" || ln === "commentRangeEnd") && attr(c, "id") === id) {
      touched = true;
      return false;
    }
    // A run whose only content is this comment's reference mark goes with it.
    if (ln === "r") {
      const refs = c.children.filter(
        (rc) => localName(rc.name) === "commentReference" && attr(rc, "id") === id,
      );
      if (refs.length > 0) {
        c.children = c.children.filter((rc) => !refs.includes(rc));
        touched = true;
        const meaningful = c.children.some((rc) => localName(rc.name) !== "rPr");
        if (!meaningful) return false;
      }
    }
    return true;
  });
  for (const c of el.children) {
    touched = stripMarkers(c, id) || touched;
  }
  return touched;
}
