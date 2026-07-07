import { XmlElement, attr, child, children, localName } from "../xml.js";
import { Theme } from "../model.js";

/**
 * Parse theme1.xml: major/minor latin fonts and the color scheme.
 * Theme color keys used by w:themeColor: dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink,
 * plus the aliases text1/background1/text2/background2 Word uses in documents.
 */
export function parseTheme(root: XmlElement | undefined): Theme {
  const theme: Theme = {
    majorFont: "Calibri Light",
    minorFont: "Calibri",
    colors: new Map(),
  };
  if (!root) return theme;

  const themeElements = child(root, "themeElements");
  const fontScheme = child(themeElements, "fontScheme");
  const majorLatin = child(child(fontScheme, "majorFont"), "latin");
  const minorLatin = child(child(fontScheme, "minorFont"), "latin");
  const major = majorLatin ? attr(majorLatin, "typeface") : undefined;
  const minor = minorLatin ? attr(minorLatin, "typeface") : undefined;
  if (major) theme.majorFont = major;
  if (minor) theme.minorFont = minor;

  const clrScheme = child(themeElements, "clrScheme");
  if (clrScheme) {
    for (const c of clrScheme.children) {
      const key = localName(c.name);
      let color: string | undefined;
      const srgb = child(c, "srgbClr");
      const sys = child(c, "sysClr");
      if (srgb) color = "#" + (attr(srgb, "val") ?? "000000");
      else if (sys) color = "#" + (attr(sys, "lastClr") ?? "000000");
      if (color) theme.colors.set(key, color);
    }
    // Aliases Word uses in w:themeColor
    const alias: [string, string][] = [
      ["text1", "dk1"],
      ["background1", "lt1"],
      ["text2", "dk2"],
      ["background2", "lt2"],
      // DrawingML a:schemeClr spellings
      ["tx1", "dk1"],
      ["bg1", "lt1"],
      ["tx2", "dk2"],
      ["bg2", "lt2"],
    ];
    for (const [a, b] of alias) {
      const v = theme.colors.get(b);
      if (v) theme.colors.set(a, v);
    }
  }
  return theme;
}
