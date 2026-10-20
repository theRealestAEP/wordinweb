import { FontSpec } from "../layout/types.js";
import { parseColor } from "./draw.js";
import { FontSet } from "./fonts.js";
import { SampleImage, boundedRgba } from "./images.js";
import { Resources } from "./resources.js";
import { num } from "./writer.js";

/**
 * A Canvas 2D context that records PDF operators instead of painting: what
 * the metafile converters (the `wmf` package and emf-converter) draw on in
 * the browser, so their output becomes vector content here. Coordinates
 * are the canvas's own pixel space (y down); the page drawer maps that onto
 * the picture's box with one `cm`. Bitmaps a converter draws are kept as
 * samples for the renderer to embed, named by placeholders in the operators.
 */
/** An image placeholder in the recorded operators: /__IMG{index}__ */
export const IMAGE_PLACEHOLDER = /\/__IMG(\d+)__/g;

type Matrix = [number, number, number, number, number, number];

interface State {
  ctm: Matrix;
  fillStyle: string;
  strokeStyle: string;
  lineWidth: number;
  lineCap: string;
  lineJoin: string;
  dash: number[];
  globalAlpha: number;
  font: string;
  textAlign: string;
  textBaseline: string;
}

export class ImageDataShim {
  readonly data: Uint8ClampedArray;
  readonly width: number;
  readonly height: number;
  constructor(dataOrWidth: Uint8ClampedArray | number, widthOrHeight: number, height?: number) {
    if (typeof dataOrWidth === "number") {
      this.width = dataOrWidth;
      this.height = widthOrHeight;
      this.data = new Uint8ClampedArray(this.width * this.height * 4);
    } else {
      this.data = dataOrWidth;
      this.width = widthOrHeight;
      this.height = height ?? dataOrWidth.length / 4 / widthOrHeight;
    }
  }
}

export class RecordingCanvas {
  readonly ops: string[] = [];
  readonly images: SampleImage[] = [];
  /** Pixels put on this canvas with putImageData, for a drawImage of it. */
  pixels: ImageDataShim | null = null;
  private context: RecordingContext | null = null;

  constructor(
    public width: number,
    public height: number,
    readonly fonts: FontSet,
    readonly resources: Resources,
  ) {}

  getContext(kind: string): RecordingContext | null {
    if (kind !== "2d") return null;
    return (this.context ??= new RecordingContext(this));
  }

  /** The converters export the canvas at the end; nothing is needed from
   * that, the operators are already recorded. */
  convertToBlob(): Promise<Blob> {
    return Promise.resolve(new Blob([]));
  }

  toDataURL(): string {
    return "data:image/png;base64,";
  }
}

function multiply(m: Matrix, n: Matrix): Matrix {
  // The point transform n applied after m (canvas transform(): ctm = ctm x n).
  return [
    m[0] * n[0] + m[2] * n[1],
    m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3],
    m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4],
    m[1] * n[4] + m[3] * n[5] + m[5],
  ];
}

const FONT = /^\s*(?:(italic|oblique)\s+)?(?:(normal|bold|bolder|lighter|[1-9]00)\s+)?(?:(normal|small-caps)\s+)?(\d+(?:\.\d+)?)(px|pt)\s+(.+)$/;

export class RecordingContext {
  private state: State = {
    ctm: [1, 0, 0, 1, 0, 0],
    fillStyle: "#000000",
    strokeStyle: "#000000",
    lineWidth: 1,
    lineCap: "butt",
    lineJoin: "miter",
    dash: [],
    globalAlpha: 1,
    font: "10px sans-serif",
    textAlign: "start",
    textBaseline: "alphabetic",
  };
  private readonly stack: State[] = [];
  /** The current path's operators, points already in canvas pixel space. */
  private path: string[] = [];
  /** Last point in user space, for quadratic curves and arcTo. */
  private last: [number, number] | null = null;
  private start: [number, number] | null = null;
  lineDashOffset = 0;
  globalCompositeOperation = "source-over";
  miterLimit = 10;

  constructor(readonly canvas: RecordingCanvas) {}

  get fillStyle(): string {
    return this.state.fillStyle;
  }
  set fillStyle(value: unknown) {
    if (typeof value === "string") this.state.fillStyle = value;
  }
  get strokeStyle(): string {
    return this.state.strokeStyle;
  }
  set strokeStyle(value: unknown) {
    if (typeof value === "string") this.state.strokeStyle = value;
  }
  get lineWidth(): number {
    return this.state.lineWidth;
  }
  set lineWidth(value: number) {
    if (value > 0) this.state.lineWidth = value;
  }
  get lineCap(): string {
    return this.state.lineCap;
  }
  set lineCap(value: string) {
    this.state.lineCap = value;
  }
  get lineJoin(): string {
    return this.state.lineJoin;
  }
  set lineJoin(value: string) {
    this.state.lineJoin = value;
  }
  get globalAlpha(): number {
    return this.state.globalAlpha;
  }
  set globalAlpha(value: number) {
    if (value >= 0 && value <= 1) this.state.globalAlpha = value;
  }
  get font(): string {
    return this.state.font;
  }
  set font(value: string) {
    if (FONT.test(value)) this.state.font = value;
  }
  get textAlign(): string {
    return this.state.textAlign;
  }
  set textAlign(value: string) {
    this.state.textAlign = value;
  }
  get textBaseline(): string {
    return this.state.textBaseline;
  }
  set textBaseline(value: string) {
    this.state.textBaseline = value;
  }

  setLineDash(segments: number[]): void {
    this.state.dash = segments.filter((v) => Number.isFinite(v) && v >= 0);
  }
  getLineDash(): number[] {
    return [...this.state.dash];
  }

  save(): void {
    this.stack.push({ ...this.state, ctm: [...this.state.ctm], dash: [...this.state.dash] });
    this.canvas.ops.push("q");
  }
  restore(): void {
    const state = this.stack.pop();
    if (!state) return;
    this.state = state;
    this.canvas.ops.push("Q");
  }
  /** Restores every save still open. A converter can leave saves open (an
   * EMF's SaveDC or clip with no restore): harmless on a canvas it throws
   * away, but here the operators go into the page's content. */
  restoreAll(): void {
    while (this.stack.length > 0) this.restore();
  }

  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.state.ctm = [a, b, c, d, e, f];
  }
  resetTransform(): void {
    this.state.ctm = [1, 0, 0, 1, 0, 0];
  }
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.state.ctm = multiply(this.state.ctm, [a, b, c, d, e, f]);
  }
  translate(x: number, y: number): void {
    this.transform(1, 0, 0, 1, x, y);
  }
  scale(x: number, y: number): void {
    this.transform(x, 0, 0, y, 0, 0);
  }
  rotate(angle: number): void {
    this.transform(Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0);
  }

  /** A user-space point in canvas pixels. */
  private device(x: number, y: number): [number, number] {
    const m = this.state.ctm;
    return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
  }
  private pt(x: number, y: number): string {
    const [dx, dy] = this.device(x, y);
    return `${num(dx)} ${num(dy)}`;
  }
  /** How the transform scales lengths, for line widths and dashes. */
  private lengthScale(): number {
    const m = this.state.ctm;
    return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
  }

  beginPath(): void {
    this.path = [];
    this.last = null;
    this.start = null;
  }
  moveTo(x: number, y: number): void {
    this.path.push(`${this.pt(x, y)} m`);
    this.last = [x, y];
    this.start = [x, y];
  }
  lineTo(x: number, y: number): void {
    if (!this.last) return this.moveTo(x, y);
    this.path.push(`${this.pt(x, y)} l`);
    this.last = [x, y];
  }
  bezierCurveTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): void {
    if (!this.last) this.moveTo(x1, y1);
    this.path.push(`${this.pt(x1, y1)} ${this.pt(x2, y2)} ${this.pt(x, y)} c`);
    this.last = [x, y];
  }
  quadraticCurveTo(cx: number, cy: number, x: number, y: number): void {
    if (!this.last) this.moveTo(cx, cy);
    const [x0, y0] = this.last as [number, number];
    this.bezierCurveTo(x0 + (2 / 3) * (cx - x0), y0 + (2 / 3) * (cy - y0), x + (2 / 3) * (cx - x), y + (2 / 3) * (cy - y), x, y);
  }
  closePath(): void {
    if (this.path.length === 0) return;
    this.path.push("h");
    if (this.start) this.last = this.start;
  }
  rect(x: number, y: number, w: number, h: number): void {
    this.moveTo(x, y);
    this.lineTo(x + w, y);
    this.lineTo(x + w, y + h);
    this.lineTo(x, y + h);
    this.closePath();
  }
  arc(x: number, y: number, r: number, a0: number, a1: number, counterclockwise = false): void {
    this.ellipse(x, y, r, r, 0, a0, a1, counterclockwise);
  }
  ellipse(x: number, y: number, rx: number, ry: number, rotation: number, a0: number, a1: number, counterclockwise = false): void {
    let sweep = a1 - a0;
    if (counterclockwise) {
      if (sweep > 0) sweep -= 2 * Math.PI * Math.ceil(sweep / (2 * Math.PI));
      if (sweep <= -2 * Math.PI) sweep = -2 * Math.PI;
    } else {
      if (sweep < 0) sweep += 2 * Math.PI * Math.ceil(-sweep / (2 * Math.PI));
      if (sweep >= 2 * Math.PI) sweep = 2 * Math.PI;
    }
    const cos = Math.cos(rotation);
    const sin = Math.sin(rotation);
    const at = (t: number): [number, number] => [x + rx * Math.cos(t) * cos - ry * Math.sin(t) * sin, y + rx * Math.cos(t) * sin + ry * Math.sin(t) * cos];
    const tangent = (t: number): [number, number] => [-rx * Math.sin(t) * cos - ry * Math.cos(t) * sin, -rx * Math.sin(t) * sin + ry * Math.cos(t) * cos];
    const [sx, sy] = at(a0);
    if (this.last) this.lineTo(sx, sy);
    else this.moveTo(sx, sy);
    const segments = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2)));
    const step = sweep / segments;
    const k = (4 / 3) * Math.tan(step / 4);
    let t = a0;
    for (let i = 0; i < segments; i++) {
      const [px, py] = at(t);
      const [qx, qy] = at(t + step);
      const [tx1, ty1] = tangent(t);
      const [tx2, ty2] = tangent(t + step);
      this.bezierCurveTo(px + k * tx1, py + k * ty1, qx - k * tx2, qy - k * ty2, qx, qy);
      t += step;
    }
  }
  arcTo(x1: number, y1: number, x2: number, y2: number, r: number): void {
    if (!this.last) return this.moveTo(x1, y1);
    const [x0, y0] = this.last;
    const d0 = Math.hypot(x1 - x0, y1 - y0);
    const d1 = Math.hypot(x2 - x1, y2 - y1);
    if (r <= 0 || d0 === 0 || d1 === 0) return this.lineTo(x1, y1);
    const ux = (x0 - x1) / d0;
    const uy = (y0 - y1) / d0;
    const vx = (x2 - x1) / d1;
    const vy = (y2 - y1) / d1;
    const angle = Math.acos(Math.max(-1, Math.min(1, ux * vx + uy * vy)));
    if (angle < 1e-6 || Math.abs(angle - Math.PI) < 1e-6) return this.lineTo(x1, y1);
    const tangentLength = r / Math.tan(angle / 2);
    const t1: [number, number] = [x1 + ux * tangentLength, y1 + uy * tangentLength];
    const t2: [number, number] = [x1 + vx * tangentLength, y1 + vy * tangentLength];
    const cross = ux * vy - uy * vx;
    const bx = (ux + vx) / Math.hypot(ux + vx, uy + vy);
    const by = (uy + vy) / Math.hypot(ux + vx, uy + vy);
    const cx = x1 + (bx * r) / Math.sin(angle / 2);
    const cy = y1 + (by * r) / Math.sin(angle / 2);
    this.lineTo(t1[0], t1[1]);
    const a0 = Math.atan2(t1[1] - cy, t1[0] - cx);
    const a1 = Math.atan2(t2[1] - cy, t2[0] - cx);
    this.ellipse(cx, cy, r, r, 0, a0, a1, cross > 0);
  }

  private paintOps(fill: boolean, stroke: boolean): string {
    const ops: string[] = [];
    if (this.state.globalAlpha < 1) ops.push(`/${this.canvas.resources.alpha(this.state.globalAlpha)} gs`);
    if (fill) {
      const c = parseColor(this.state.fillStyle) ?? { r: 0, g: 0, b: 0, a: 1 };
      ops.push(`${num(c.r / 255)} ${num(c.g / 255)} ${num(c.b / 255)} rg`);
      if (c.a < 1 && this.state.globalAlpha === 1) ops.push(`/${this.canvas.resources.alpha(c.a)} gs`);
    }
    if (stroke) {
      const c = parseColor(this.state.strokeStyle) ?? { r: 0, g: 0, b: 0, a: 1 };
      const scale = this.lengthScale();
      ops.push(`${num(c.r / 255)} ${num(c.g / 255)} ${num(c.b / 255)} RG ${num(Math.max(this.state.lineWidth * scale, 0.2))} w`);
      ops.push(`${this.state.lineCap === "round" ? 1 : this.state.lineCap === "square" ? 2 : 0} J ${this.state.lineJoin === "round" ? 1 : this.state.lineJoin === "bevel" ? 2 : 0} j`);
      if (this.state.dash.length > 0) ops.push(`[${this.state.dash.map((v) => num(v * scale)).join(" ")}] ${num(this.lineDashOffset * scale)} d`);
    }
    return ops.join(" ");
  }

  fill(rule?: string): void {
    if (this.path.length === 0) return;
    this.canvas.ops.push(`q ${this.paintOps(true, false)} ${this.path.join(" ")} ${rule === "evenodd" ? "f*" : "f"} Q`);
  }
  stroke(): void {
    if (this.path.length === 0) return;
    this.canvas.ops.push(`q ${this.paintOps(false, true)} ${this.path.join(" ")} S Q`);
  }
  clip(rule?: string): void {
    if (this.path.length === 0) return;
    this.canvas.ops.push(`${this.path.join(" ")} ${rule === "evenodd" ? "W*" : "W"} n`);
  }
  fillRect(x: number, y: number, w: number, h: number): void {
    const saved = this.path;
    this.path = [];
    this.rect(x, y, w, h);
    this.fill();
    this.path = saved;
  }
  strokeRect(x: number, y: number, w: number, h: number): void {
    const saved = this.path;
    this.path = [];
    this.rect(x, y, w, h);
    this.stroke();
    this.path = saved;
  }
  clearRect(x: number, y: number, w: number, h: number): void {
    const saved = this.path;
    this.path = [];
    this.rect(x, y, w, h);
    this.canvas.ops.push(`q 1 g ${this.path.join(" ")} f Q`);
    this.path = saved;
  }

  /** The FontSpec a CSS font string names, and its size in px. */
  private fontSpec(): FontSpec {
    const m = FONT.exec(this.state.font);
    const size = m ? Number(m[4]) * (m[5] === "pt" ? 96 / 72 : 1) : 10;
    const family = (m ? m[6] : "sans-serif").split(",")[0].trim().replace(/^["']|["']$/g, "");
    const weight = m?.[2];
    return { family, size, bold: weight === "bold" || weight === "bolder" || Number(weight) >= 600, italic: !!m?.[1] };
  }

  measureText(text: string): { width: number } {
    const font = this.fontSpec();
    let width = 0;
    for (const run of this.canvas.fonts.shape(text, font)) for (const advance of run.advances) width += advance;
    return { width };
  }

  fillText(text: string, x: number, y: number): void {
    if (!text) return;
    const font = this.fontSpec();
    const runs = this.canvas.fonts.shape(text, font);
    const width = runs.reduce((sum, run) => sum + run.advances.reduce((s, a) => s + a, 0), 0);
    const align = this.state.textAlign;
    let dx = align === "center" ? -width / 2 : align === "right" || align === "end" ? -width : 0;
    const face = runs[0]?.face.face;
    const ascent = face ? (face.ascent / face.unitsPerEm) * font.size : font.size * 0.8;
    const descent = face ? (-face.descent / face.unitsPerEm) * font.size : font.size * 0.2;
    const baseline = this.state.textBaseline;
    const dy = baseline === "top" || baseline === "hanging" ? ascent : baseline === "middle" ? (ascent - descent) / 2 : baseline === "bottom" || baseline === "ideographic" ? -descent : 0;
    const m = this.state.ctm;
    // The text's direction and size follow the transform; glyphs stay
    // upright in the y-down space through the flipped text matrix.
    const scale = Math.hypot(m[0], m[1]) || 1;
    const cos = m[0] / scale;
    const sin = m[1] / scale;
    const c = parseColor(this.state.fillStyle) ?? { r: 0, g: 0, b: 0, a: 1 };
    const ops = [`q ${num(c.r / 255)} ${num(c.g / 255)} ${num(c.b / 255)} rg BT`];
    for (const run of runs) {
      const embedded = this.canvas.resources.font(run.face.face);
      const [ox, oy] = this.device(x + dx, y + dy);
      ops.push(`/${embedded.name} ${num(font.size * scale)} Tf ${num(cos)} ${num(sin)} ${num(sin)} ${num(-cos)} ${num(ox)} ${num(oy)} Tm`);
      let hex = "";
      run.gids.forEach((gid, i) => {
        hex += embedded.cid(gid, run.codePoints[i]).toString(16).padStart(4, "0");
        dx += run.advances[i];
      });
      ops.push(`<${hex}> Tj`);
    }
    ops.push("ET Q");
    this.canvas.ops.push(ops.join(" "));
  }
  strokeText(text: string, x: number, y: number): void {
    this.fillText(text, x, y);
  }

  createImageData(width: number, height: number): ImageDataShim {
    return new ImageDataShim(width, height);
  }
  getImageData(_x: number, _y: number, width: number, height: number): ImageDataShim {
    return new ImageDataShim(Math.max(1, width), Math.max(1, height));
  }
  /** Pixels put on a temporary canvas are what a later drawImage of it
   * shows; put on a canvas being drawn, they are drawn where they land. */
  putImageData(data: ImageDataShim, x: number, y: number): void {
    if (this.canvas.ops.length === 0 && x === 0 && y === 0) {
      this.canvas.pixels = data;
      return;
    }
    const saved = this.state.ctm;
    this.state.ctm = [1, 0, 0, 1, 0, 0];
    this.drawImage({ pixels: data, width: data.width, height: data.height }, x, y, data.width, data.height);
    this.state.ctm = saved;
  }

  drawImage(source: { pixels?: ImageDataShim | null; width: number; height: number }, ...args: number[]): void {
    const pixels = source.pixels;
    if (!pixels) return;
    let [dx, dy, dw, dh] = args;
    if (args.length >= 8) [, , , , dx, dy, dw, dh] = args;
    if (dw === undefined) dw = source.width;
    if (dh === undefined) dh = source.height;
    const index = this.canvas.images.length;
    this.canvas.images.push(boundedRgba(pixels.data, pixels.width, pixels.height));
    // The image's unit square onto its destination, in canvas pixels.
    const m = multiply(this.state.ctm, [dw, 0, 0, dh, dx, dy]);
    const ops = ["q"];
    if (this.state.globalAlpha < 1) ops.push(`/${this.canvas.resources.alpha(this.state.globalAlpha)} gs`);
    // The image space has y up: flip within the destination box.
    ops.push(`${num(m[0])} ${num(m[1])} ${num(-m[2])} ${num(-m[3])} ${num(m[4] + m[2])} ${num(m[5] + m[3])} cm /__IMG${index}__ Do Q`);
    this.canvas.ops.push(ops.join(" "));
  }
}

/**
 * Run `work` with OffscreenCanvas and ImageData globals that record into
 * canvases bound to these fonts and resources, as emf-converter looks for
 * them, and give back the first canvas it made: the picture itself (any
 * later ones are its bitmaps' temporaries).
 */
export async function withRecordingCanvas<T>(fonts: FontSet, resources: Resources, work: () => Promise<T>): Promise<{ canvas: RecordingCanvas | null; result: T }> {
  const g = globalThis as Record<string, unknown>;
  const hadCanvas = "OffscreenCanvas" in g;
  const hadImageData = "ImageData" in g;
  const savedCanvas = g.OffscreenCanvas;
  const savedImageData = g.ImageData;
  let first: RecordingCanvas | null = null;
  g.OffscreenCanvas = class extends RecordingCanvas {
    constructor(width: number, height: number) {
      super(width, height, fonts, resources);
      first ??= this;
    }
  };
  if (!hadImageData) g.ImageData = ImageDataShim;
  try {
    const result = await work();
    return { canvas: first, result };
  } finally {
    if (hadCanvas) g.OffscreenCanvas = savedCanvas;
    else delete g.OffscreenCanvas;
    if (hadImageData) g.ImageData = savedImageData;
    else delete g.ImageData;
  }
}
