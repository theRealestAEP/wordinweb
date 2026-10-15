import { num } from "./writer.js";

/** A parsed SVG path: moves, lines, cubic curves and closes, absolute. */
export type PathSegment =
  | { op: "m"; x: number; y: number }
  | { op: "l"; x: number; y: number }
  | { op: "c"; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: "h" };

/** SVG path data as segments; quadratics and arcs become cubics. */
export function parsePath(d: string): PathSegment[] {
  const tokens = d.match(/[MmLlHhVvCcSsQqTtAaZz]|-?(?:\d+\.?\d*|\.\d+)(?:e-?\d+)?/g);
  if (!tokens) return [];
  const out: PathSegment[] = [];
  let i = 0;
  let cmd = "";
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  // The previous control point, for the S and T shorthands.
  let cx = 0;
  let cy = 0;
  let lastCmd = "";
  const next = (): number => Number(tokens[i++]);
  const cubic = (x1: number, y1: number, x2: number, y2: number, ex: number, ey: number): void => {
    out.push({ op: "c", x1, y1, x2, y2, x: ex, y: ey });
    cx = x2;
    cy = y2;
    x = ex;
    y = ey;
  };
  const quadratic = (qx: number, qy: number, ex: number, ey: number): void => {
    cubic(x + (2 / 3) * (qx - x), y + (2 / 3) * (qy - y), ex + (2 / 3) * (qx - ex), ey + (2 / 3) * (qy - ey), ex, ey);
    cx = qx;
    cy = qy;
  };
  while (i < tokens.length) {
    if (/[A-Za-z]/.test(tokens[i])) cmd = tokens[i++];
    else if (cmd === "M") cmd = "L";
    else if (cmd === "m") cmd = "l";
    if (i >= tokens.length && cmd !== "Z" && cmd !== "z") break;
    const rel = cmd === cmd.toLowerCase() && cmd !== "z";
    const dx = rel ? x : 0;
    const dy = rel ? y : 0;
    switch (cmd.toUpperCase()) {
      case "M":
        x = dx + next();
        y = dy + next();
        startX = x;
        startY = y;
        out.push({ op: "m", x, y });
        break;
      case "L":
        x = dx + next();
        y = dy + next();
        out.push({ op: "l", x, y });
        break;
      case "H":
        x = dx + next();
        out.push({ op: "l", x, y });
        break;
      case "V":
        y = dy + next();
        out.push({ op: "l", x, y });
        break;
      case "C": {
        const x1 = dx + next();
        const y1 = dy + next();
        const x2 = dx + next();
        const y2 = dy + next();
        cubic(x1, y1, x2, y2, dx + next(), dy + next());
        break;
      }
      case "S": {
        const reflect = /[CcSs]/.test(lastCmd);
        const x1 = reflect ? 2 * x - cx : x;
        const y1 = reflect ? 2 * y - cy : y;
        const x2 = dx + next();
        const y2 = dy + next();
        cubic(x1, y1, x2, y2, dx + next(), dy + next());
        break;
      }
      case "Q": {
        const qx = dx + next();
        const qy = dy + next();
        quadratic(qx, qy, dx + next(), dy + next());
        break;
      }
      case "T": {
        const reflect = /[QqTt]/.test(lastCmd);
        quadratic(reflect ? 2 * x - cx : x, reflect ? 2 * y - cy : y, dx + next(), dy + next());
        break;
      }
      case "A": {
        const rx = next();
        const ry = next();
        const rotation = next();
        const large = next() !== 0;
        const sweep = next() !== 0;
        const ex = dx + next();
        const ey = dy + next();
        for (const [x1, y1, x2, y2, px, py] of arcToCubics(x, y, rx, ry, rotation, large, sweep, ex, ey)) cubic(x1, y1, x2, y2, px, py);
        x = ex;
        y = ey;
        break;
      }
      case "Z":
        out.push({ op: "h" });
        x = startX;
        y = startY;
        break;
    }
    lastCmd = cmd;
  }
  return out;
}

/**
 * SVG path data as PDF path operators, each point mapped through `point`
 * (the item's scale and the page's flip).
 */
export function pathToPdf(d: string, point: (x: number, y: number) => [number, number]): string {
  const P = (px: number, py: number): string => {
    const [ox, oy] = point(px, py);
    return `${num(ox)} ${num(oy)}`;
  };
  return parsePath(d)
    .map((s) => (s.op === "h" ? "h" : s.op === "c" ? `${P(s.x1, s.y1)} ${P(s.x2, s.y2)} ${P(s.x, s.y)} c` : `${P(s.x, s.y)} ${s.op}`))
    .join(" ");
}

/** The path as a polyline (its first subpath), curves in 16 steps each. */
export function flattenPath(d: string): [number, number][] {
  const points: [number, number][] = [];
  for (const s of parsePath(d)) {
    if (s.op === "h") break;
    if (s.op === "m" && points.length > 0) break;
    if (s.op === "m" || s.op === "l") points.push([s.x, s.y]);
    else {
      const [x0, y0] = points[points.length - 1] ?? [s.x1, s.y1];
      for (let k = 1; k <= 16; k++) {
        const t = k / 16;
        const u = 1 - t;
        points.push([
          u * u * u * x0 + 3 * u * u * t * s.x1 + 3 * u * t * t * s.x2 + t * t * t * s.x,
          u * u * u * y0 + 3 * u * u * t * s.y1 + 3 * u * t * t * s.y2 + t * t * t * s.y,
        ]);
      }
    }
  }
  return points;
}

/** An SVG arc as cubic Bézier segments of at most a quarter turn each. */
function arcToCubics(x1: number, y1: number, rx: number, ry: number, rotation: number, large: boolean, sweep: boolean, x2: number, y2: number): number[][] {
  if (rx === 0 || ry === 0 || (x1 === x2 && y1 === y2)) return [[x1, y1, x2, y2, x2, y2]];
  rx = Math.abs(rx);
  ry = Math.abs(ry);
  const phi = (rotation * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const mx = (x1 - x2) / 2;
  const my = (y1 - y2) / 2;
  const x1p = cos * mx + sin * my;
  const y1p = -sin * mx + cos * my;
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) {
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  const sign = large === sweep ? -1 : 1;
  const numerator = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const coefficient = sign * Math.sqrt(Math.max(0, numerator) / (rx * rx * y1p * y1p + ry * ry * x1p * x1p));
  const cxp = (coefficient * rx * y1p) / ry;
  const cyp = (-coefficient * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const angle = (ux: number, uy: number, vx: number, vy: number): number => {
    const dot = ux * vx + uy * vy;
    const len = Math.hypot(ux, uy) * Math.hypot(vx, vy);
    let a = Math.acos(Math.max(-1, Math.min(1, dot / len)));
    if (ux * vy - uy * vx < 0) a = -a;
    return a;
  };
  const theta1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let delta = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  else if (sweep && delta < 0) delta += 2 * Math.PI;
  const segments = Math.max(1, Math.ceil(Math.abs(delta) / (Math.PI / 2)));
  const step = delta / segments;
  const out: number[][] = [];
  let theta = theta1;
  const at = (t: number): [number, number] => [cx + rx * Math.cos(t) * cos - ry * Math.sin(t) * sin, cy + rx * Math.cos(t) * sin + ry * Math.sin(t) * cos];
  const tangent = (t: number): [number, number] => [-rx * Math.sin(t) * cos - ry * Math.cos(t) * sin, -rx * Math.sin(t) * sin + ry * Math.cos(t) * cos];
  for (let s = 0; s < segments; s++) {
    const t2 = theta + step;
    const k = (4 / 3) * Math.tan(step / 4);
    const [px, py] = at(theta);
    const [qx, qy] = at(t2);
    const [tx1, ty1] = tangent(theta);
    const [tx2, ty2] = tangent(t2);
    out.push([px + k * tx1, py + k * ty1, qx - k * tx2, qy - k * ty2, qx, qy]);
    theta = t2;
  }
  return out;
}
