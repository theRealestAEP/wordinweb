import type { SmartArtData, SmartArtLayout } from "../model.js";
import { type XmlElement, attr, child, children, localName } from "../xml.js";

function descendants(node: XmlElement | undefined, name: string): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (current: XmlElement | undefined): void => {
    if (!current) return;
    for (const item of current.children) {
      if (localName(item.name) === name) out.push(item);
      walk(item);
    }
  };
  walk(node);
  return out;
}

/** Read editable node text and a useful layout family from native SmartArt parts. */
export function parseSmartArtParts(dataRoot: XmlElement, layoutRoot?: XmlElement): SmartArtData | null {
  const docPoint = children(child(dataRoot, "ptLst"), "pt").find((point) => attr(point, "type") === "doc");
  const layoutId = `${attr(child(docPoint, "prSet"), "loTypeId") ?? ""} ${attr(layoutRoot, "uniqueId") ?? ""}`.toLowerCase();
  const layout: SmartArtLayout = layoutId.includes("hierarchy") || layoutId.includes("orgchart")
    ? "hierarchy"
    : layoutId.includes("cycle")
      ? "cycle"
      : layoutId.includes("list")
        ? "list"
        : "process";
  const items = children(child(dataRoot, "ptLst"), "pt")
    .filter((point) => !["doc", "parTrans", "sibTrans"].includes(attr(point, "type") ?? ""))
    .map((point) => descendants(child(point, "t"), "t").map((text) => text.text).join("").trim())
    .filter(Boolean);
  return items.length ? { layout, items } : null;
}
