import { describe, expect, it } from "vitest";
import {
  isKnownShapeGeometry,
  knownShapeGeometryNames,
  presetFillColor,
  presetShapeGeometry,
} from "../src/preset-geometry.js";
import { SHAPE_GALLERY, shapeGalleryLabel } from "../src/shape-gallery.js";

describe("preset geometry evaluator", () => {
  it("ships the full ECMA-376 preset table minus buttons and connectors", () => {
    const names = knownShapeGeometryNames();
    expect(names.length).toBe(165);
    for (const family of ["actionButton", "bentConnector", "curvedConnector", "straightConnector"]) {
      expect(names.some((name) => name.startsWith(family))).toBe(false);
    }
    expect(isKnownShapeGeometry("heart")).toBe(true);
    expect(isKnownShapeGeometry("line")).toBe(false);
    expect(isKnownShapeGeometry("nonsense")).toBe(false);
  });

  it("evaluates every preset to finite, non-empty path data", () => {
    for (const name of knownShapeGeometryNames()) {
      for (const [w, h] of [[200, 100], [64, 64], [1219200, 914400]]) {
        const geom = presetShapeGeometry(name, w, h);
        expect(geom, name).toBeDefined();
        expect(geom!.paths.length, name).toBeGreaterThan(0);
        for (const path of geom!.paths) {
          expect(path.d.length, name).toBeGreaterThan(0);
          expect(path.d.includes("NaN"), `${name}: ${path.d}`).toBe(false);
          expect(path.d.includes("Infinity"), `${name}: ${path.d}`).toBe(false);
          // Every numeric token stays inside a sane multiple of the box.
          for (const token of path.d.split(" ")) {
            const value = Number(token);
            if (Number.isFinite(value)) {
              expect(Math.abs(value), `${name}: ${token}`).toBeLessThan(Math.max(w, h) * 20 + 21700000);
            }
          }
        }
      }
    }
  });

  it("evaluates rect to the bounding box", () => {
    const geom = presetShapeGeometry("rect", 200, 100)!;
    expect(geom.paths).toEqual([{ d: "M 0 0 L 200 0 L 200 100 L 0 100 Z", fill: "norm", stroke: true }]);
    expect(geom.textRect).toEqual({ l: 0, t: 0, r: 200, b: 100 });
  });

  it("evaluates roundRect corners from the adjustment default and override", () => {
    const geom = presetShapeGeometry("roundRect", 200, 100)!;
    // adj = 16667 of the short side 100 → radius 16.67; starts on the left
    // edge at the corner radius.
    expect(geom.paths[0].d.startsWith("M 0 16.67")).toBe(true);
    expect(geom.paths[0].d.match(/A 16\.67 16\.67/g)?.length).toBe(4);
    const square = presetShapeGeometry("roundRect", 200, 100, { adj: 0 })!;
    expect(square.paths[0].d.startsWith("M 0 0")).toBe(true);
  });

  it("evaluates the triangle apex from adj", () => {
    const geom = presetShapeGeometry("triangle", 200, 100)!;
    expect(geom.paths[0].d).toBe("M 0 100 L 100 0 L 200 100 Z");
    const skewed = presetShapeGeometry("triangle", 200, 100, { adj: 0 })!;
    expect(skewed.paths[0].d).toBe("M 0 100 L 0 0 L 200 100 Z");
  });

  it("converts ellipse arcs to four quadrant SVG arcs through the box extremes", () => {
    const geom = presetShapeGeometry("ellipse", 200, 100)!;
    const d = geom.paths[0].d;
    // Starts at the left edge midpoint (arc start angle cd2 = 180°).
    expect(d.startsWith("M 0 50")).toBe(true);
    // Four 90° arcs with rx=100 ry=50 passing top, right, bottom midpoints.
    expect(d.match(/A 100 50 0 0 1/g)?.length).toBe(4);
    expect(d).toContain("A 100 50 0 0 1 100 0");
    expect(d).toContain("A 100 50 0 0 1 200 50");
    expect(d).toContain("A 100 50 0 0 1 100 100");
    // Inscribed-rectangle text area: (1 - cos45)/2 ≈ 0.1464.
    expect(geom.textRect!.l).toBeCloseTo(200 * 0.146447, 1);
    expect(geom.textRect!.t).toBeCloseTo(100 * 0.146447, 1);
  });

  it("splits large arc sweeps at quadrant boundaries (pie with wide slice)", () => {
    // 270° slice starting at 0: three 90° arc segments.
    const geom = presetShapeGeometry("pie", 100, 100, { adj1: 0, adj2: 16200000 })!;
    expect(geom.paths[0].d.match(/A 50 50 0 0 1/g)?.length).toBe(3);
  });

  it("scales path-local coordinate spaces (flowChartDecision is a 2x2 diamond)", () => {
    const geom = presetShapeGeometry("flowChartDecision", 200, 100)!;
    expect(geom.paths[0].d).toBe("M 0 50 L 100 0 L 200 50 L 100 100 Z");
    expect(geom.textRect).toEqual({ l: 50, t: 25, r: 150, b: 75 });
  });

  it("carries per-path fill modes and stroke flags (bracketPair)", () => {
    const geom = presetShapeGeometry("bracketPair", 100, 100)!;
    // Canonical bracketPair: a fill-only path then a stroke-only path.
    expect(geom.paths.length).toBeGreaterThan(1);
    expect(geom.paths[0].stroke).toBe(false);
    expect(geom.paths.at(-1)!.fill).toBe("none");
    expect(geom.paths.at(-1)!.stroke).toBe(true);
  });

  it("backs every gallery entry with an evaluable geometry", () => {
    const seen = new Set<string>();
    let total = 0;
    for (const category of SHAPE_GALLERY) {
      for (const entry of category.items) {
        expect(isKnownShapeGeometry(entry.preset), entry.preset).toBe(true);
        expect(seen.has(entry.preset), `duplicate ${entry.preset}`).toBe(false);
        seen.add(entry.preset);
        total++;
      }
    }
    expect(total).toBe(147);
    expect(shapeGalleryLabel("heart")).toBe("Heart");
    expect(shapeGalleryLabel("gear6")).toBe("gear6");
  });

  it("resolves fill-mode colours as shades and tints of the base fill", () => {
    expect(presetFillColor("#4472C4", "norm")).toBe("#4472C4");
    expect(presetFillColor("#4472C4", "none")).toBeUndefined();
    expect(presetFillColor(undefined, "norm")).toBeUndefined();
    expect(presetFillColor("#804020", "darken")).toBe("#402010");
    expect(presetFillColor("#804020", "lighten")).toBe("#C0A090");
  });
});
