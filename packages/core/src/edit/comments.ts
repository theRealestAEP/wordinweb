import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName } from "../xml.js";

/**
 * Delete a review comment like Word: remove its entry from comments.xml and
 * strip the commentRangeStart/End markers and commentReference run from the
 * document. Callers checkpoint history and rerender afterwards.
 */
export function deleteComment(doc: DocxDocument, id: string): boolean {
  let touched = false;

  const commentsRoot = doc.commentsTree();
  if (commentsRoot) {
    const before = commentsRoot.children.length;
    commentsRoot.children = commentsRoot.children.filter(
      (c) => !(localName(c.name) === "comment" && attr(c, "id") === id),
    );
    touched = commentsRoot.children.length !== before;
  }

  // Strip markers from every editable tree (comments can sit in headers too).
  for (const root of doc.editableRoots()) {
    if (root === commentsRoot) continue;
    touched = stripMarkers(root, id) || touched;
  }

  if (touched) doc.refresh();
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
