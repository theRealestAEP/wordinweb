(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

function fakeContext() {
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
  HTMLCanvasElement.prototype.getContext = function () {
    return fakeContext();
  } as never;
}

class FakeOffscreenCanvas {
  constructor(
    public width = 0,
    public height = 0,
  ) {}

  getContext() {
    return fakeContext();
  }
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = FakeOffscreenCanvas;

if (typeof document !== "undefined" && !document.elementFromPoint) {
  (document as { elementFromPoint?: (x: number, y: number) => Element | null }).elementFromPoint =
    () => null;
}

if (typeof Range !== "undefined" && !Range.prototype.getBoundingClientRect) {
  Range.prototype.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
      width: 0,
      height: 0,
      x: 0,
      y: 0,
      toJSON() {},
    }) as DOMRect;
}

if (typeof Element !== "undefined" && !Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => {};
}
