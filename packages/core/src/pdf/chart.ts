import { ChartItem, FontSpec, TextItem } from "../layout/types.js";
import { Theme } from "../model.js";
import { CHART_LABEL_COLOR, CHART_TEXT_PX, ChartPen, paintChart } from "../render/chart-paint.js";
import { PageDrawer, parseColor } from "./draw.js";
import { pathToPdf } from "./path.js";
import { num } from "./writer.js";

/** A chart drawn with the DOM renderer's painters, through a pen that
 * turns their SVG-shaped marks into PDF operators at the item's box. */
export function drawChart(drawer: PageDrawer, item: ChartItem, theme: Theme): void {
  const data = item.data;
  const family = (theme.colors.size ? theme.minorFont : "Aptos").split(",")[0].trim();
  const size = data.textSize ?? CHART_TEXT_PX;
  const fill = (attrs: Record<string, string | number>, fallback: string): string => {
    const color = parseColor(String(attrs.fill ?? fallback));
    return color ? `${num(color.r / 255)} ${num(color.g / 255)} ${num(color.b / 255)} rg` : "";
  };
  const stroke = (attrs: Record<string, string | number>): string => {
    const color = attrs.stroke === undefined ? null : parseColor(String(attrs.stroke));
    if (!color) return "";
    const width = Number(attrs["stroke-width"] ?? 1) * 0.75;
    const dash = attrs["stroke-dasharray"] ? `[${String(attrs["stroke-dasharray"]).split(/[\s,]+/).map((v) => num(Number(v) * 0.75)).join(" ")}] 0 d ` : "";
    const cap = attrs["stroke-linecap"] === "round" ? "1 J " : "";
    const join = attrs["stroke-linejoin"] === "round" ? "1 j " : "";
    return `${dash}${cap}${join}${num(width)} w ${num(color.r / 255)} ${num(color.g / 255)} ${num(color.b / 255)} RG`;
  };
  const paint = (f: string, s: string): string => (f && s ? "B" : f ? "f" : s ? "S" : "n");
  const pen: ChartPen = {
    data,
    colors: [],
    size,
    font: family,
    add(name, attrs) {
      const n = (key: string): number => Number(attrs[key] ?? 0);
      if (name === "rect") {
        // An SVG rect fills black unless told otherwise; the painters always say.
        const f = fill(attrs, "#000000");
        const s = stroke(attrs);
        drawer.ops.push(`q ${f} ${s} ${drawer.rect(item.x + n("x"), item.y + n("y"), n("width"), n("height"))} ${paint(f, s)} Q`);
      } else if (name === "line") {
        const s = stroke(attrs);
        if (!s) return;
        drawer.ops.push(`q ${s} ${num(drawer.x(item.x + n("x1")))} ${num(drawer.y(item.y + n("y1")))} m ${num(drawer.x(item.x + n("x2")))} ${num(drawer.y(item.y + n("y2")))} l S Q`);
      } else {
        const f = attrs.fill === "none" ? "" : fill(attrs, "#000000");
        const s = stroke(attrs);
        const d = pathToPdf(String(attrs.d ?? ""), (x, y) => [drawer.x(item.x + x), drawer.y(item.y + y)]);
        if (d) drawer.ops.push(`q ${f} ${s} ${d} ${paint(f, s)} Q`);
      }
    },
    label(x, y, text, attrs = {}) {
      if (!text) return;
      const fontSize = Number(attrs["font-size"] ?? size);
      const font: FontSpec = { family, size: fontSize, bold: false, italic: false };
      const width = drawer.textWidth(text, font);
      const anchor = attrs["text-anchor"];
      const left = item.x + x - (anchor === "middle" ? width / 2 : anchor === "end" ? width : 0);
      const lineTop = item.y + y - fontSize;
      const rotate = /rotate\((-?[\d.]+) (-?[\d.]+) (-?[\d.]+)\)/.exec(String(attrs.transform ?? ""));
      const label: TextItem = {
        kind: "text",
        x: left,
        baseline: item.y + y,
        width,
        text,
        props: { color: String(attrs.fill ?? CHART_LABEL_COLOR) },
        font,
        lineTop,
        lineHeight: fontSize * 1.2,
        ...(rotate ? { rotate: { deg: Number(rotate[1]), ox: item.x + Number(rotate[2]) - left, oy: item.y + Number(rotate[3]) - lineTop } } : {}),
      };
      drawer.drawItem(label, () => null);
    },
  };
  paintChart(pen, item, theme);
}
