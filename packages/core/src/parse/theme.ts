import { XmlElement, attr, child, children, localName } from "../xml.js";
import { Theme } from "../model.js";

/**
 * Parse theme1.xml: major/minor latin fonts and the color scheme.
 * Theme color keys used by w:themeColor: dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink,
 * plus the aliases text1/background1/text2/background2 Word uses in documents.
 */
export function parseTheme(root: XmlElement | undefined, bidiLanguage?: string): Theme {
  const theme: Theme = {
    majorFont: "Calibri Light",
    minorFont: "Calibri",
    colors: new Map(),
  };
  if (!root) return theme;

  const themeElements = child(root, "themeElements");
  const fontScheme = child(themeElements, "fontScheme");
  const majorFonts = child(fontScheme, "majorFont");
  const minorFonts = child(fontScheme, "minorFont");
  const majorLatin = child(majorFonts, "latin");
  const minorLatin = child(minorFonts, "latin");
  const major = majorLatin ? attr(majorLatin, "typeface") : undefined;
  const minor = minorLatin ? attr(minorLatin, "typeface") : undefined;
  if (major) theme.majorFont = major;
  if (minor) theme.minorFont = minor;

  // East Asian faces (<a:ea>) feed eastAsiaTheme="minor/majorEastAsia". The
  // default theme leaves these empty (typeface=""), meaning "use the app's
  // language default"; treat an empty value as absent so the CJK fallback
  // (defaultEastAsia in layout) applies instead of an empty family.
  const majorEa = attr(child(majorFonts, "ea"), "typeface");
  const minorEa = attr(child(minorFonts, "ea"), "typeface");
  if (majorEa) theme.majorEastAsiaFont = majorEa;
  if (minorEa) theme.minorEastAsiaFont = minorEa;

  const bidiScript = bidiLanguage ? new Intl.Locale(bidiLanguage).maximize().script : undefined;
  const bidiFont = (fonts: XmlElement | undefined): string | undefined => {
    const direct = attr(child(fonts, "cs"), "typeface");
    if (direct) return direct;
    if (!bidiScript) return undefined;
    const supplemental = children(fonts, "font").find((font) => attr(font, "script") === bidiScript);
    return supplemental ? attr(supplemental, "typeface") : undefined;
  };
  theme.majorBidiFont = bidiFont(majorFonts);
  theme.minorBidiFont = bidiFont(minorFonts);

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
