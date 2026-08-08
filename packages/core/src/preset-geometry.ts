/**
 * DrawingML preset-geometry evaluator.
 *
 * Evaluates the parametric preset definitions (preset-definitions.ts) into
 * SVG path data for a concrete shape size: the guide-formula language of
 * ECMA-376 §20.1.9.11 (17 ops over the predefined w/h/ss/… variables, angles
 * in 60000ths of a degree) plus the pathLst command set (§20.1.9.15) with
 * arcTo converted to SVG elliptical arcs.
 */
import { PRESET_DEFINITIONS, type PresetDef } from "./preset-definitions.js";

export type PresetFillMode = "norm" | "none" | "darken" | "darkenLess" | "lighten" | "lightenLess";

export interface PresetGeomPath {
  /** SVG path data in the `w x h` coordinate space handed to the evaluator. */
  d: string;
  fill: PresetFillMode;
  stroke: boolean;
}

export interface PresetGeom {
  paths: PresetGeomPath[];
  /** The geometry's text rectangle in the same coordinate space. */
  textRect?: { l: number; t: number; r: number; b: number };
}

export function isKnownShapeGeometry(prst: string): boolean {
  return Object.prototype.hasOwnProperty.call(PRESET_DEFINITIONS, prst);
}

export function knownShapeGeometryNames(): string[] {
  return Object.keys(PRESET_DEFINITIONS);
}

const DEG_PER_UNIT = 1 / 60000;
const RAD = Math.PI / 180;

/** Angle in guide units (60000ths of a degree) to radians. */
function rad(units: number): number {
  return units * DEG_PER_UNIT * RAD;
}

class GuideEnv {
  private values = new Map<string, number>();
  constructor(private w: number, private h: number) {}

  set(name: string, value: number): void {
    this.values.set(name, value);
  }

  has(name: string): boolean {
    return this.values.has(name);
  }

  /** A formula token: a number literal or a guide/predefined-variable name. */
  token(token: string): number {
    const known = this.values.get(token);
    if (known !== undefined) return known;
    const numeric = Number(token);
    if (Number.isFinite(numeric)) return numeric;
    return this.builtin(token);
  }

  /** The predefined guides of §20.1.9.11 (w/h/ss/ls, centers/edges, angle
   * constants, and the wdN/hdN/ssdN divided families). */
  private builtin(name: string): number {
    const { w, h } = this;
    switch (name) {
      case "w": case "r": return w;
      case "h": case "b": return h;
      case "l": case "t": return 0;
      case "hc": return w / 2;
      case "vc": return h / 2;
      case "ss": return Math.min(w, h);
      case "ls": return Math.max(w, h);
      case "cd8": return 2700000;
      case "cd4": return 5400000;
      case "cd2": return 10800000;
      case "3cd8": return 8100000;
      case "3cd4": return 16200000;
      case "5cd8": return 13500000;
      case "7cd8": return 18900000;
    }
    const divided = /^(wd|hd|ssd)(\d+)$/.exec(name);
    if (divided) {
      const base = divided[1] === "wd" ? w : divided[1] === "hd" ? h : Math.min(w, h);
      return base / Number(divided[2]);
    }
    return 0;
  }

  /** Evaluate one guide formula (e.g. multiply-divide "w adj 100000", "val 16667"). */
  formula(fmla: string): number {
    const parts = fmla.split(/ +/);
    const a = parts.length > 1 ? this.token(parts[1]) : 0;
    const b = parts.length > 2 ? this.token(parts[2]) : 0;
    const c = parts.length > 3 ? this.token(parts[3]) : 0;
    switch (parts[0]) {
      case "val": return a;
      case "*/": return c === 0 ? 0 : (a * b) / c;
      case "+-": return a + b - c;
      case "+/": return c === 0 ? 0 : (a + b) / c;
      case "?:": return a > 0 ? b : c;
      case "abs": return Math.abs(a);
      case "min": return Math.min(a, b);
      case "max": return Math.max(a, b);
      case "mod": return Math.sqrt(a * a + b * b + c * c);
      case "pin": return b < a ? a : b > c ? c : b;
      case "sqrt": return Math.sqrt(Math.max(a, 0));
      case "sin": return a * Math.sin(rad(b));
      case "cos": return a * Math.cos(rad(b));
      case "tan": return a * Math.tan(rad(b));
      case "at2": return Math.atan2(b, a) / RAD / DEG_PER_UNIT;
      case "cat2": return a * Math.cos(Math.atan2(c, b));
      case "sat2": return a * Math.sin(Math.atan2(c, b));
      default: return 0;
    }
  }
}

function fmt(value: number): string {
  const rounded = Math.round(value * 100) / 100;
  return String(Object.is(rounded, -0) ? 0 : rounded);
}

/**
 * The point on the ellipse with radii (rw, rh) at the stated TRUE angle theta
 * (radians, y-down): DrawingML measures arc angles as ray angles from the
 * center, so the parametric angle t with P = c + (rw·cos t, rh·sin t) solves
 * tan t = (rw/rh)·tan theta.
 */
function ellipsePoint(rw: number, rh: number, theta: number): { x: number; y: number } {
  // Snap the trig of quarter-turn angles to exact 0: with a degenerate radius
  // (rh = 0, a bracket arm) the float noise in sin(2π) ≈ -2.4e-16 would win
  // the atan2 against an exact rh·cos = 0 and collapse the endpoint to the
  // wrong end of the segment.
  let s = Math.sin(theta);
  let c = Math.cos(theta);
  if (Math.abs(s) < 1e-12) s = 0;
  if (Math.abs(c) < 1e-12) c = 0;
  const t = Math.atan2(rw * s, rh * c);
  return { x: rw * Math.cos(t), y: rh * Math.sin(t) };
}

/**
 * Append SVG elliptical-arc segments for a:arcTo starting at (x, y).
 * The sweep is split at quadrant boundaries: inside one quadrant the
 * true-angle → parametric-angle map stays in that quadrant, so every emitted
 * segment spans < 180° and the large-arc flag is always 0.
 */
function arcTo(
  out: string[],
  x: number,
  y: number,
  rw: number,
  rh: number,
  stAngUnits: number,
  swAngUnits: number,
): { x: number; y: number } {
  if (swAngUnits === 0) return { x, y };
  const start = rad(stAngUnits);
  const sweep = rad(swAngUnits);
  if (rw <= 0 || rh <= 0) {
    // Degenerate ellipse (e.g. a bracket with adj=0): the arc collapses to a
    // straight segment to its endpoint — SVG's zero-radius rule, and what Word
    // paints (wild2-med-phase23-protocol p14's rightBracket arms).
    const startPt = ellipsePoint(rw, rh, start);
    const endPt = ellipsePoint(rw, rh, start + sweep);
    const nx = x - startPt.x + endPt.x;
    const ny = y - startPt.y + endPt.y;
    out.push(`L ${fmt(nx)} ${fmt(ny)}`);
    return { x: nx, y: ny };
  }
  const startPt = ellipsePoint(rw, rh, start);
  const cx = x - startPt.x;
  const cy = y - startPt.y;
  const dir = sweep > 0 ? 1 : -1;
  const sweepFlag = sweep > 0 ? 1 : 0;
  const quarter = Math.PI / 2;
  let consumed = 0;
  let current = { x, y };
  const total = Math.abs(sweep);
  while (consumed < total - 1e-9) {
    // Advance to the next quadrant boundary (or the sweep end).
    const theta = start + dir * consumed;
    const boundary = dir > 0
      ? (Math.floor(theta / quarter + 1e-9) + 1) * quarter
      : (Math.ceil(theta / quarter - 1e-9) - 1) * quarter;
    const step = Math.min(Math.abs(boundary - theta), total - consumed);
    consumed += step;
    const endPt = ellipsePoint(rw, rh, start + dir * consumed);
    current = { x: cx + endPt.x, y: cy + endPt.y };
    out.push(`A ${fmt(rw)} ${fmt(rh)} 0 0 ${sweepFlag} ${fmt(current.x)} ${fmt(current.y)}`);
  }
  return current;
}

function buildPathD(tokens: string[], env: GuideEnv, sx: number, sy: number): string {
  const out: string[] = [];
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  let i = 0;
  const coord = (scale: number): number => env.token(tokens[i++]) * scale;
  while (i < tokens.length) {
    const cmd = tokens[i++];
    switch (cmd) {
      case "M":
        x = startX = coord(sx);
        y = startY = coord(sy);
        out.push(`M ${fmt(x)} ${fmt(y)}`);
        break;
      case "L":
        x = coord(sx);
        y = coord(sy);
        out.push(`L ${fmt(x)} ${fmt(y)}`);
        break;
      case "Q": {
        const cx = coord(sx);
        const cy = coord(sy);
        x = coord(sx);
        y = coord(sy);
        out.push(`Q ${fmt(cx)} ${fmt(cy)} ${fmt(x)} ${fmt(y)}`);
        break;
      }
      case "C": {
        const c1x = coord(sx);
        const c1y = coord(sy);
        const c2x = coord(sx);
        const c2y = coord(sy);
        x = coord(sx);
        y = coord(sy);
        out.push(`C ${fmt(c1x)} ${fmt(c1y)} ${fmt(c2x)} ${fmt(c2y)} ${fmt(x)} ${fmt(y)}`);
        break;
      }
      case "A": {
        const rw = coord(sx);
        const rh = coord(sy);
        const stAng = env.token(tokens[i++]);
        const swAng = env.token(tokens[i++]);
        ({ x, y } = arcTo(out, x, y, rw, rh, stAng, swAng));
        break;
      }
      case "Z":
        out.push("Z");
        x = startX;
        y = startY;
        break;
      default:
        return out.join(" "); // malformed table entry; keep what parsed
    }
  }
  return out.join(" ");
}

/**
 * Evaluate a preset geometry at a concrete `w x h` size.
 *
 * `adjust` carries the shape's a:avLst overrides (gd name → the numeric value
 * of its "val N" formula), replacing the preset's adjustment defaults.
 * Any unit works for w/h (EMU, px) — the output is in the same space.
 */
export function presetShapeGeometry(
  prst: string,
  w: number,
  h: number,
  adjust?: Record<string, number>,
): PresetGeom | undefined {
  const def: PresetDef | undefined = PRESET_DEFINITIONS[prst];
  if (!def || !(w > 0) || !(h > 0)) return undefined;
  const env = new GuideEnv(w, h);
  for (const [name, fmla] of def.av ?? []) {
    const override = adjust?.[name];
    env.set(name, override !== undefined ? override : env.formula(fmla));
  }
  for (const [name, fmla] of def.gd ?? []) env.set(name, env.formula(fmla));

  const paths: PresetGeomPath[] = [];
  for (const pathDef of def.paths) {
    const sx = pathDef.w ? w / pathDef.w : 1;
    const sy = pathDef.h ? h / pathDef.h : 1;
    const d = buildPathD(pathDef.c.split(" "), env, sx, sy);
    if (d) paths.push({ d, fill: pathDef.fill ?? "norm", stroke: pathDef.stroke !== false });
  }
  if (!paths.length) return undefined;
  // Strokes paint over fills: a stroke-only path listed before a filled one
  // (chartPlus's cross, flowChartInternalStorage's dividers) must not vanish
  // under the fill, so order filled paths first (stable within each group).
  paths.sort((a, b) => Number(a.fill === "none") - Number(b.fill === "none"));

  let textRect: PresetGeom["textRect"];
  if (def.rect) {
    const [l, t, r, b] = def.rect.map((token) => env.token(token));
    if (r > l && b > t) textRect = { l, t, r, b };
  }
  return { paths, ...(textRect ? { textRect } : {}) };
}

/**
 * Resolve a preset path's fill mode against the shape's fill colour: the
 * darken/lighten modes paint 3D-ish facets as shades/tints of the base fill.
 * The exact factors are not spec'd; these follow common OOXML renderers.
 */
export function presetFillColor(base: string | undefined, mode: PresetFillMode): string | undefined {
  if (!base || mode === "none") return undefined;
  if (mode === "norm") return base;
  const hex = /^#?([0-9a-fA-F]{6})$/.exec(base)?.[1];
  if (!hex) return base;
  const channel = (offset: number): number => parseInt(hex.slice(offset, offset + 2), 16);
  const mix = (value: number): string => {
    const target = mode.startsWith("darken") ? 0 : 255;
    const amount = mode.endsWith("Less") ? 0.25 : 0.5;
    return Math.round(value + (target - value) * amount).toString(16).padStart(2, "0");
  };
  return `#${mix(channel(0))}${mix(channel(2))}${mix(channel(4))}`.toUpperCase();
}
