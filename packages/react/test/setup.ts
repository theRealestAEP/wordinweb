// React 18 requires this flag for act() to flush effects in a test env.
(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

// jsdom has no canvas; DocxView's text measurer needs one. Provide a minimal
// fake 2D context that measures text width ~ length so layout can run headless
// (proportional metrics are fine — these tests assert content is painted, not
// pixel-exact geometry). Mirrors the stub in core's measure.test.ts.
function fakeCtx() {
  return {
    font: "",
    measureText(text: string) {
      const size = Number(/([\d.]+)px/.exec(this.font)?.[1]) || 12;
      return { width: size * 0.5 * text.length } as TextMetrics;
    },
    fillText() {},
    clearRect() {},
    setTransform() {},
    scale() {},
    translate() {},
    save() {},
    restore() {},
    beginPath() {},
    fill() {},
    drawImage() {},
  } as unknown as CanvasRenderingContext2D;
}

if (typeof HTMLCanvasElement !== "undefined") {
  HTMLCanvasElement.prototype.getContext = function () { return fakeCtx(); } as never;
}
class FakeOffscreenCanvas {
  constructor(public width = 0, public height = 0) {}
  getContext() { return fakeCtx(); }
}
(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = FakeOffscreenCanvas;

// jsdom doesn't implement elementFromPoint (no layout). Returning null sends
// the editor's caret hit-testing down its nearest-page fallback, which scans
// .dxw-page rects — the same code path a real browser uses for clicks outside
// any text span.
if (typeof document !== "undefined" && !document.elementFromPoint) {
  (document as { elementFromPoint?: (x: number, y: number) => Element | null }).elementFromPoint = () => null;
}

// jsdom Ranges have no layout; the editor measures a collapsed Range to place
// the caret bar. A zero rect keeps positionCaret on its degenerate-rect branch
// (which uses layout-item geometry instead), matching how empty/whitespace
// spans behave in a real browser.
if (typeof Range !== "undefined" && !Range.prototype.getBoundingClientRect) {
  Range.prototype.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0, x: 0, y: 0, toJSON() {} }) as DOMRect;
}
// scrollIntoView is a no-op in jsdom.
if (typeof Element !== "undefined" && !Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => {};
}
