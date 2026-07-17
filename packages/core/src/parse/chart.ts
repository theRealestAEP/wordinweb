import type { ChartData, ChartSeries, ChartType } from "../model.js";
import { XmlElement, attr, child, children, localName } from "../xml.js";

function descendant(node: XmlElement | undefined, name: string): XmlElement | undefined {
  if (!node) return undefined;
  for (const item of node.children) {
    if (localName(item.name) === name) return item;
    const found = descendant(item, name);
    if (found) return found;
  }
  return undefined;
}

function descendants(node: XmlElement | undefined, name: string): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (item: XmlElement | undefined): void => {
    if (!item) return;
    for (const current of item.children) {
      if (localName(current.name) === name) out.push(current);
      walk(current);
    }
  };
  walk(node);
  return out;
}

function pointValues(cache: XmlElement | undefined): string[] {
  const points = children(cache, "pt")
    .map((point) => ({ index: Number(attr(point, "idx") ?? 0), value: child(point, "v")?.text ?? "" }))
    .sort((a, b) => a.index - b.index);
  return points.map((point) => point.value);
}

function seriesName(series: XmlElement): string {
  const tx = child(series, "tx");
  return descendant(tx, "v")?.text ?? "Series";
}

/** Resolve the cached display data from a native ChartML part. */
export function parseChartPart(root: XmlElement): ChartData | null {
  const plotArea = descendant(root, "plotArea");
  if (!plotArea) return null;
  const chartEl = plotArea.children.find((item) => ["barChart", "lineChart", "pieChart"].includes(localName(item.name)));
  if (!chartEl) return null;
  const local = localName(chartEl.name);
  const type: ChartType = local === "lineChart"
    ? "line"
    : local === "pieChart"
      ? "pie"
      : attr(child(chartEl, "barDir"), "val") === "bar"
        ? "bar"
        : "column";
  const seriesEls = children(chartEl, "ser");
  const series: ChartSeries[] = seriesEls.map((item) => ({
    name: seriesName(item),
    values: pointValues(descendant(child(item, "val"), "numCache")).map((value) => Number(value) || 0),
  }));
  const first = seriesEls[0];
  const categories = pointValues(
    descendant(child(first, "cat"), "strCache") ?? descendant(child(first, "cat"), "numCache"),
  );
  const title = descendants(descendant(root, "title"), "t").map((item) => item.text).join("").trim();
  return {
    type,
    ...(title ? { title } : {}),
    categories,
    series,
  };
}
