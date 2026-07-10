declare module "wmf" {
  interface WmfAction {
    t: string;
    v?: string;
    s?: {
      Extent?: [number, number];
      Font?: { Name: string; Height: number; Weight: number; Italic: boolean; Angle: number };
      Pen?: { Style: number; Width: number; Color: number };
      Brush?: { Style: number; Color: number };
      [key: string]: unknown;
    };
  }

  const WMF: {
    image_size(data: ArrayBuffer | Uint8Array): [number, number];
    get_actions(data: ArrayBuffer | Uint8Array): WmfAction[];
    render_canvas(actions: WmfAction[], canvas: HTMLCanvasElement): void;
  };

  export default WMF;
}
