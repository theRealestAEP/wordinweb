import type { Theme } from "../model.js";
import { XmlElement, attr, child, intAttr, localName } from "../xml.js";

/** Apply a:lumMod/lumOff/shade/tint children to a hex color. */
export function applyClrTransforms(hex: string, clrEl: XmlElement): string {
  let r = parseInt(hex.slice(1, 3), 16) / 255;
  let g = parseInt(hex.slice(3, 5), 16) / 255;
  let b = parseInt(hex.slice(5, 7), 16) / 255;
  for (const t of clrEl.children) {
    const v = (intAttr(t, "val") ?? 100000) / 100000;
    switch (localName(t.name)) {
      case "lumMod":
        r *= v; g *= v; b *= v;
        break;
      case "lumOff":
        r += v; g += v; b += v;
        break;
      case "shade":
        r *= v; g *= v; b *= v;
        break;
      case "tint":
        r = 1 - (1 - r) * v; g = 1 - (1 - g) * v; b = 1 - (1 - b) * v;
        break;
    }
  }
  const c = (x: number) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, "0");
  return `#${c(r)}${c(g)}${c(b)}`;
}

/** Resolve the color child (srgbClr/schemeClr/sysClr) of a fill element. */
export function solidColorOf(solid: XmlElement, theme: Theme | undefined): string | undefined {
  const clrEl = child(solid, "srgbClr") ?? child(solid, "schemeClr") ?? child(solid, "sysClr");
  if (!clrEl) return undefined;
  const local = localName(clrEl.name);
  let hex: string | undefined;
  if (local === "srgbClr") hex = "#" + (attr(clrEl, "val") ?? "000000");
  else if (local === "sysClr") hex = "#" + (attr(clrEl, "lastClr") ?? "000000");
  else hex = theme?.colors.get(attr(clrEl, "val") ?? "");
  if (!hex) return undefined;
  return applyClrTransforms(hex, clrEl);
}

/** Resolve the a:solidFill inside `container` to a CSS color: srgbClr, theme
 * schemeClr, or sysClr (whose lastClr carries the rendered color), with
 * lumMod/lumOff/shade/tint transforms applied. */
export function solidFillColor(container: XmlElement | undefined, theme: Theme | undefined): string | undefined {
  if (!container) return undefined;
  const solid = child(container, "solidFill");
  if (!solid) return undefined;
  return solidColorOf(solid, theme);
}
