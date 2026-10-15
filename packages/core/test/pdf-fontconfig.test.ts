import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { FontSet, fontconfig } from "../src/pdf/fonts.js";

/** fontconfig's family for a pattern, or undefined without fc-match. */
function fcFamily(pattern: string): string | undefined {
  try {
    return execFileSync("fc-match", ["-f", "%{family[0]}", pattern], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }) || undefined;
  } catch {
    return undefined;
  }
}

const hasFontconfig = fcFamily("sans-serif") !== undefined;
// The system font directories, as the renderer scans them by default.
const dirs = ["/usr/share/fonts", "/System/Library/Fonts", "/Library/Fonts"].filter((dir) => existsSync(dir));

describe("fontconfig", () => {
  it.skipIf(!hasFontconfig)("treats its generic answer for a family that cannot exist as not found", () => {
    expect(fontconfig.match("wordinweb-no-such-family", false, false)).toBeNull();
    expect(fontconfig.match("wordinweb-no-such-family", true, true)).toBeNull();
  });

  it.skipIf(!hasFontconfig)("answers once per question", () => {
    const first = fontconfig.match("DejaVu Sans", false, false);
    expect(fontconfig.match("DejaVu Sans", false, false)).toBe(first);
  });

  // The alias rules fonts-urw-base35 installs: Palatino Linotype -> P052.
  it.skipIf(fcFamily("Palatino Linotype") !== "P052")("resolves a family through fontconfig's alias rules when no directory holds it", () => {
    const fonts = new FontSet(dirs);
    const stack = fonts.resolve({ family: "Palatino Linotype", size: 12, bold: false, italic: false });
    expect(stack.family).toBe("P052");
    expect(stack.faces[0].face.family).toBe("P052");
    expect(fonts.missing.has("Palatino Linotype")).toBe(false);
    const bold = fonts.resolve({ family: "Palatino Linotype", size: 12, bold: true, italic: false });
    expect(bold.faces[0].face.bold).toBe(true);
    expect(bold.faces[0].fakeBold).toBe(false);
  });

  it.skipIf(!hasFontconfig)("marks a family missing when fontconfig has only its generic answer", () => {
    const fonts = new FontSet(dirs);
    const stack = fonts.resolve({ family: "Wordinweb Imaginary Grotesk", size: 12, bold: false, italic: false });
    expect(stack.faces.length).toBeGreaterThan(0);
    expect(fonts.missing.has("Wordinweb Imaginary Grotesk")).toBe(true);
  });

  // fonts-noto-cjk: an East Asian face by language, serif for Song and Mincho.
  it.skipIf(!/CJK/.test(fcFamily("serif:lang=zh-cn") ?? ""))("gives an East Asian family a face for its language when the Mac faces are absent", () => {
    const fonts = new FontSet(dirs);
    const song = fonts.resolve({ family: "Songti SC", paintFamily: "SimSun", size: 12, bold: false, italic: false });
    const mincho = fonts.resolve({ family: "Hiragino Mincho ProN", paintFamily: "MS Mincho", size: 12, bold: false, italic: false });
    const hei = fonts.resolve({ family: "PingFang TC", paintFamily: "Microsoft JhengHei", size: 12, bold: false, italic: false });
    if (!dirs.includes("/System/Library/Fonts")) {
      expect(song.family).toBe("Noto Serif CJK SC");
      expect(mincho.family).toBe("Noto Serif CJK JP");
      expect(hei.family).toBe("Noto Sans CJK TC");
      expect(fonts.missing.size).toBe(0);
    }
    for (const stack of [song, mincho, hei]) expect(stack.faces.some((f) => f.face.hasGlyph(0x4e2d))).toBe(true);
    // A Chinese character in a Latin run falls to the language's face too.
    fonts.shape("中文 text", { family: "Carlito", size: 12, bold: false, italic: false });
    expect(fonts.missingCodePoints.size).toBe(0);
  });
});
