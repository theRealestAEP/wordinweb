/**
 * The baseline geometry of DrawingML preset text warps (a:prstTxWarp), free
 * of any surface: the DOM renderer rides SVG text on it and the PDF renderer
 * places glyphs along it.
 */

export interface WarpGeo {
  d: string;
  /** Fixed font px, or 0 to derive it (see `natural`/`pour`). */
  fontPx: number;
  pathLen: number;
  /** Where along the path the text starts, 0..1. */
  startFrac: number;
  /** Anchor the text at the path start ("start") or its midpoint ("middle"). */
  anchor: "start" | "middle";
  /** Stretch the text to span the whole baseline (fills the box width). */
  fill: boolean;
  /** Ride the text at its natural run font size instead of a box-derived one
   * (textArchUp: Word keeps the glyphs small and only bows the baseline). */
  natural?: boolean;
  /** Size the glyphs so the natural-spaced string wraps most of the circle. */
  pour?: boolean;
}

/** DrawingML a:prstTxWarp baseline geometry in the shape's `W x H` box.
 * Calibrated against Word's probe3-wordart-warps PDF: arch keeps small
 * top-centred text on a shallow bow; wave/chevron scale the text to fill the
 * box; circle-pour wraps the string clockwise from the top of a circle. */
export function warpBaseline(warp: string, W: number, H: number): WarpGeo {
  const mx = W * 0.02;
  const bw = W - 2 * mx;
  switch (warp) {
    case "textArchUp":
    case "textArchUpPour": {
      // Word rides "ARCH UP" at its run size, centred, near the top, on a very
      // shallow ∩ (text ink measured at x∈[0.34,0.64]·W, y∈[0.01,0.13]·H).
      const yEnd = H * 0.16;
      const yApex = H * 0.09;
      const cy = 2 * yApex - yEnd;
      const d = `M ${W * 0.12} ${yEnd} Q ${W / 2} ${cy} ${W * 0.88} ${yEnd}`;
      return { d, fontPx: 0, pathLen: (bw + 2 * Math.hypot(bw * 0.38, yEnd - cy)) / 2, startFrac: 0.5, anchor: "middle", fill: false, natural: true };
    }
    case "textArchDown":
    case "textArchDownPour": {
      const yTop = H * 0.84;
      const yApex = H * 0.91;
      const cy = 2 * yApex - yTop;
      const d = `M ${W * 0.12} ${yTop} Q ${W / 2} ${cy} ${W * 0.88} ${yTop}`;
      return { d, fontPx: 0, pathLen: (bw + 2 * Math.hypot(bw * 0.38, cy - yTop)) / 2, startFrac: 0.5, anchor: "middle", fill: false, natural: true };
    }
    case "textWave1":
    case "textWave2": {
      // Word's WAVE ONE fills the box (ink ~93%×90%, centred). One symmetric
      // sine period centred on the mid-line: crest left, trough right, big
      // amplitude so the large glyphs sweep from near the top to near the
      // bottom (measured centroid 0.28·H crest → 0.76·H trough).
      const fontPx = H * 0.78;
      const d =
        `M ${mx} ${H * 0.5} ` +
        `C ${W * 0.16} ${H * 0.33} ${W * 0.34} ${H * 0.33} ${W * 0.5} ${H * 0.5} ` +
        `C ${W * 0.66} ${H * 0.67} ${W * 0.84} ${H * 0.67} ${W - mx} ${H * 0.5}`;
      return { d, fontPx, pathLen: bw * 1.06, startFrac: 0, anchor: "start", fill: true };
    }
    case "textChevron": {
      // Word's CHEVRON is a symmetric downward valley (∨) filling the box, its
      // apex ROUNDED — a quadratic, not two straight legs, so no glyph straddles
      // a sharp vertex (the stray-letter artifact). Ends high, middle low.
      const fontPx = H * 0.78;
      const yEnds = H * 0.26;
      const yCtrl = H * 1.18; // quadratic control → midpoint baseline ≈ 0.72·H
      const d = `M ${mx} ${yEnds} Q ${W / 2} ${yCtrl} ${W - mx} ${yEnds}`;
      const legs = 2 * Math.hypot(bw / 2, (yEnds + yCtrl) / 2 - yEnds);
      return { d, fontPx, pathLen: (bw + legs) / 2, startFrac: 0, anchor: "start", fill: true };
    }
    case "textChevronInverted": {
      const fontPx = H * 0.46;
      const yBot = H * 0.66;
      const yTop = H * 0.34;
      const d = `M ${mx} ${yBot} L ${W / 2} ${yTop} L ${W - mx} ${yBot}`;
      return { d, fontPx, pathLen: 2 * Math.hypot(bw / 2, yBot - yTop), startFrac: 0, anchor: "start", fill: true };
    }
    case "textCircle":
    case "textCirclePour":
    case "textButton":
    case "textButtonPour":
    default: {
      // Text poured clockwise around a circle inscribed in the box. The
      // baseline radius leaves room for the OUTWARD-pointing glyphs so the ring
      // stays inside the box (Word's is a neat centred ring, not an overflowing
      // one). The path STARTS at ~10 o'clock (−60° from top) so the string's
      // first word rides centred across the top arc.
      const r = Math.min(W, H) * 0.3;
      const cx = W / 2;
      const cy = H / 2;
      const a = (-90 * Math.PI) / 180;
      const sx = cx + r * Math.sin(a);
      const sy = cy - r * Math.cos(a);
      const d = `M ${sx.toFixed(2)} ${sy.toFixed(2)} A ${r} ${r} 0 1 1 ${(sx - 0.01).toFixed(2)} ${sy.toFixed(2)} Z`;
      return { d, fontPx: 0, pathLen: 2 * Math.PI * r, startFrac: 0, anchor: "start", fill: false, pour: true };
    }
  }
}
