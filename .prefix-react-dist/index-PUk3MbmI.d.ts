import * as react from 'react';

/**
 * Minimal, fast XML parser producing a lightweight element tree.
 *
 * OOXML parts are machine-generated, namespace-prefixed, well-formed XML with
 * no DTDs. This parser handles exactly that subset: elements, attributes,
 * text, CDATA, comments, processing instructions, and the built-in + numeric
 * character entities. Namespace prefixes are kept verbatim (`w:p`), and
 * lookups match on either the full name or the local name, which is robust
 * for every mainstream producer (Word, LibreOffice, Google Docs export).
 */
interface XmlElement {
    name: string;
    attrs: Record<string, string>;
    children: XmlElement[];
    /** Concatenated character data directly inside this element. */
    text: string;
    /** Editor placeholder retained in memory but omitted from XML while empty. */
    omitWhenEmpty?: boolean;
}

/**
 * Typed document model produced by the parser and consumed by the layout
 * engine. All lengths are CSS px (see units.ts) unless a field name says
 * otherwise. Property bags are Partial-style: `undefined` means "not set at
 * this level" so the style-inheritance merge can distinguish absence from an
 * explicit value.
 *
 * Model nodes keep `src` references to the XML elements they were parsed
 * from: the XML tree is the source of truth for editing — commands mutate it
 * and the model is re-derived, which preserves round-trip fidelity for
 * everything untouched.
 */

type BorderStyle = "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave" | "dotDash" | "dotDotDash" | "thinThickSmallGap" | "triple";
interface Border {
    style: BorderStyle;
    /** Stroke width in px. */
    width: number;
    /** CSS color. */
    color: string;
    /** Gap between border and content in px (w:space, points in OOXML). */
    space: number;
    /** Declared stroke width in px BEFORE the 0.75px paint floor (w:sz/8 pt).
     * Word's LAYOUT math (table row border share) uses this true width — a
     * sz-4 rule contributes 0.5pt, not the floored 0.5625pt (phase23 p66:
     * 45 rows drift 2.5px down without it). */
    rawWidth?: number;
}
interface ParagraphBorders {
    top?: Border;
    bottom?: Border;
    left?: Border;
    right?: Border;
    /** Drawn between consecutive paragraphs that both specify it. */
    between?: Border;
}
type LineSpacingRule = "auto" | "atLeast" | "exact";
interface LineSpacing {
    rule: LineSpacingRule;
    /**
     * For "auto": multiple of single line height (240ths in OOXML, stored here
     * as the multiplier, e.g. 1.15). For atLeast/exact: px.
     */
    value: number;
}
type Alignment = "left" | "center" | "right" | "justify";
interface TabStop {
    /** px from the left text edge */
    pos: number;
    align: "left" | "center" | "right" | "decimal" | "bar";
    leader: "none" | "dot" | "hyphen" | "underscore" | "middleDot";
    /** w:val="clear": removes the inherited stop at this pos when tab lists
     * merge down the style chain (never a live stop itself). */
    clear?: boolean;
}
interface RunProps {
    bold?: boolean;
    italic?: boolean;
    underline?: string;
    strike?: boolean;
    doubleStrike?: boolean;
    /** Resolved primary font family name (ascii/hAnsi, theme-resolved). */
    font?: string;
    /** East Asian font (rFonts w:eastAsia): used for CJK codepoints. */
    fontEastAsia?: string;
    /** hAnsi font (rFonts w:hAnsi, theme-resolved): Word uses this channel for
     * non-ASCII, non-CJK, non-complex characters (curly quotes, ≤, dashes). A
     * run that declares only w:ascii leaves it inherited from the style chain. */
    fontHAnsi?: string;
    /** Complex-script font (rFonts w:cs): used for RTL/complex runs. */
    fontComplex?: string;
    /** w:rtl — this run's text is right-to-left (Arabic/Hebrew). */
    rtl?: boolean;
    /** Font size in px. */
    size?: number;
    /** CSS color; "auto" resolved to inherit/black by renderer. */
    color?: string;
    highlight?: string;
    /** Character shading fill (w:shd). */
    shading?: string;
    /** Run border (w:bdr): a box painted around the run. Word merges adjacent
     * runs with identical borders into one box; a run wrapping across lines
     * closes the box on every line segment. */
    border?: Border;
    verticalAlign?: "baseline" | "superscript" | "subscript";
    caps?: boolean;
    smallCaps?: boolean;
    /** Letter spacing px (w:spacing, twips in OOXML). */
    letterSpacing?: number;
    /** Minimum font size in px at which w:kern enables pair kerning; null is
     * the explicit w:kern=0 override. */
    kerningThreshold?: number | null;
    /** Paint-only letter spacing when layout advances intentionally differ. */
    paintLetterSpacing?: number;
    /** w:w horizontal character scaling as a fraction (1.5 = 150%). */
    textScale?: number;
    /** w:fitText target width in px: the run's advances scale so its text
     * occupies exactly this width (resolved to textScale at layout time). */
    fitText?: number;
    /** w:position baseline shift in px, positive = raised. */
    raise?: number;
    /** w:outline — hollow stroked glyphs. */
    outline?: boolean;
    /** w:emboss / w:imprint — Word triple-draws offset copies. */
    emboss?: boolean;
    imprint?: boolean;
    vanish?: boolean;
    /** Character style id (w:rStyle). */
    styleId?: string;
    lang?: string;
}
interface NumberingRef {
    numId: number;
    ilvl: number;
}
interface ParaProps {
    styleId?: string;
    alignment?: Alignment;
    /** The exact OOXML w:jc value when it is a "justify" flavor whose behavior
     * differs from plain "both": `distribute` also stretches the LAST line, and
     * the Arabic `*Kashida` levels elongate glyphs (kashida) so lines pack fewer
     * words and — at medium/high — wrap to more lines. Absent for plain "both".
     * `alignment` stays "justify" for all of these. */
    justifyKind?: "distribute" | "lowKashida" | "mediumKashida" | "highKashida";
    /** w:bidi — right-to-left paragraph: lines assemble RTL and default
     * alignment flips to the right edge. */
    bidi?: boolean;
    /** px */
    indentLeft?: number;
    indentRight?: number;
    /** Positive px; mutually exclusive with hanging in effect. */
    indentFirstLine?: number;
    /** Positive px hanging indent. */
    indentHanging?: number;
    spacingBefore?: number;
    spacingAfter?: number;
    /** w:beforeAutospacing/afterAutospacing: HTML-style automatic paragraph
     * spacing — Word ignores the literal before/after and inserts one blank
     * line's worth of space. */
    beforeAutospacing?: boolean;
    afterAutospacing?: boolean;
    /** Word's default when absent is auto/1.0 via docDefaults. */
    lineSpacing?: LineSpacing;
    /** w:snapToGrid: whether this paragraph's lines use the section docGrid. */
    snapToGrid?: boolean;
    contextualSpacing?: boolean;
    keepNext?: boolean;
    /** w:framePr w:dropCap: the paragraph is a drop-cap letter frame. */
    dropCap?: {
        mode: "drop" | "margin";
        lines: number;
        hSpace: number;
        pageAnchored?: boolean;
    };
    /** w:framePr positioned text frame (absolute-positioned paragraph the body
     * text wraps around). Geometry in px. Individual attributes are optional so a
     * framePr merges attribute-by-attribute across the style cascade (a direct
     * framePr that sets only h/x/y keeps the style framePr's width/anchor — IEEE
     * authors); the engine normalizes defaults at layout time. */
    frame?: {
        w?: number;
        h?: number;
        hRule?: "auto" | "atLeast" | "exact";
        x?: number;
        y?: number;
        hAnchor?: "page" | "margin" | "text" | "column";
        vAnchor?: "page" | "margin" | "text" | "paragraph";
        xAlign?: "left" | "center" | "right" | "inside" | "outside";
        yAlign?: "top" | "center" | "bottom" | "inside" | "outside" | "inline";
        wrap?: "around" | "auto" | "notBeside" | "through" | "tight" | "none";
        /** w:hSpace / w:vSpace: horizontal/vertical margin around the frame (px). */
        hSpace?: number;
        vSpace?: number;
    };
    keepLines?: boolean;
    pageBreakBefore?: boolean;
    widowControl?: boolean;
    borders?: ParagraphBorders;
    /** Paragraph shading fill as CSS color. */
    shading?: string;
    numbering?: NumberingRef | null;
    /** A numPr that carries ilvl but no numId (Heading3 basedOn Heading2)
     * overrides only the list LEVEL, keeping the inherited numId; the style
     * chain applies it to `numbering` on merge. */
    numberingLevelOverride?: number;
    tabs?: TabStop[];
    outlineLevel?: number;
    /** Run props declared on pPr/rPr — apply to the paragraph mark & numbering label. */
    markRunProps?: RunProps;
    /**
     * styleId of the enclosing table (set on cell paragraphs at parse time).
     * A table style's own pPr layers between docDefaults and the paragraph
     * style, e.g. TableGrid's `spacing after=0 line=240` overrides docDefaults'
     * `after=200 line=276` so list cells lay out compactly.
     */
    tableStyleId?: string;
    /**
     * Grid position of the enclosing cell plus the table's tblLook (set on cell
     * paragraphs at parse time, like tableStyleId). Lets run/paragraph
     * resolution layer the applicable w:tblStylePr conditional formats (e.g.
     * firstRow bold + white) between the table style's own rPr and the
     * paragraph style chain.
     */
    tableCellCond?: {
        look?: TableLook;
        rowIdx: number;
        nRows: number;
        colStart: number;
        colSpan: number;
        nCols: number;
    };
}
interface TextContent {
    kind: "text";
    text: string;
    /** Source w:t element inside the run, when this text came verbatim from one. */
    srcT?: XmlElement;
    /** When this text is the glyph of a modern checkbox content control, the
     * source `w14:checkbox` element (in the enclosing w:sdt's sdtPr). Its
     * presence makes the glyph a click-to-toggle target rather than plain text. */
    checkbox?: XmlElement;
}
interface BreakContent {
    kind: "break";
    breakType: "line" | "page" | "column";
}
interface TabContent {
    kind: "tab";
}
interface PTabContent {
    kind: "ptab";
    /** Absolute-position tab: jump to left/center/right of the base. */
    alignment: "left" | "center" | "right";
    relativeTo: "margin" | "indent";
}
interface ImageContent {
    kind: "image";
    /** Part path of the image inside the package. */
    part: string;
    width: number;
    height: number;
    /** True when the drawing is floating (wp:anchor); v1 lays it out inline. */
    anchored?: boolean;
    /** a:srcRect crop, fractions of the source bitmap. */
    crop?: {
        l: number;
        t: number;
        r: number;
        b: number;
    };
    /** a:xfrm rotation, degrees clockwise. */
    rotation?: number;
    /** a:ln picture outline (Word draws it just outside the image). */
    border?: {
        color: string;
        width: number;
    };
    /** Source w:drawing (or pict) element, for resize/move editing. */
    srcDrawing?: XmlElement;
    /** Native Office 3D model shown through its saved poster raster. */
    model3D?: Model3DReference;
    /** Word online-video metadata shown through its saved poster raster. */
    webVideo?: WebVideoReference;
    /** Embedded OLE package activated as a safe download in the browser. */
    embeddedObject?: EmbeddedObjectReference;
}
interface Model3DReference {
    part: string;
    posterPart: string;
    /** Native Office model orientation, in degrees around the X/Y/Z axes. */
    rotation?: {
        x: number;
        y: number;
        z: number;
    };
}
interface WebVideoReference {
    url: string;
    embeddedHtml: string;
    width: number;
    height: number;
}
interface EmbeddedObjectReference {
    part: string;
    filename: string;
    progId: string;
}
/** OMML equation node (subset: runs, scripts, fractions, radicals). */
type MathNode = {
    t: "run";
    text: string;
    normal?: boolean;
    /** m:func function name (cos, log, …): Word kerns a thin space on
     * both sides of the name. */
    fname?: boolean;
} | {
    t: "sup" | "sub";
    base: MathNode[];
    script: MathNode[];
} | {
    t: "frac";
    num: MathNode[];
    den: MathNode[];
    bar?: boolean;
}
/** Radical; deg is the optional m:deg index (∛), absent when m:degHide. */
 | {
    t: "rad";
    e: MathNode[];
    deg?: MathNode[];
}
/** n-ary operator (sum/integral); chr defaults to the integral sign. */
 | {
    t: "nary";
    chr: string;
    sub: MathNode[];
    sup: MathNode[];
    e: MathNode[];
}
/** Delimiters grown to the content height; beg/end default to parens.
 * An empty beg/end string means that side has no delimiter (cases "{"). */
 | {
    t: "dlm";
    beg: string;
    end: string;
    e: MathNode[][];
}
/** Matrix: rows x cells. */
 | {
    t: "mat";
    rows: MathNode[][][];
}
/** m:eqArr equation array (cases/piecewise rows). */
 | {
    t: "eqarr";
    rows: MathNode[][];
}
/** m:acc combining accent (hat/bar/vector) over the base. */
 | {
    t: "acc";
    chr: string;
    e: MathNode[];
}
/** m:groupChr horizontal group character (over/under brace). */
 | {
    t: "grp";
    chr: string;
    pos: "top" | "bot";
    vertJc: "top" | "bot";
    e: MathNode[];
}
/** m:limLow / m:limUpp limit stacked under/over a text operator. */
 | {
    t: "lim";
    pos: "low" | "upp";
    e: MathNode[];
    lim: MathNode[];
};
interface MathContent {
    kind: "math";
    nodes: MathNode[];
    /** Source m:oMath element (math editing). */
    src?: XmlElement;
    /** Display equation (wrapped in m:oMathPara): centered on its own line with
     * display-style layout - larger n-ary operators with limits stacked
     * above/below, and full-size fraction numerators/denominators. */
    display?: boolean;
    /** Display group with several direct equations. Word leaves a small
     * advance after script clusters in these grouped displays. */
    multiEquationDisplay?: boolean;
    /** m:oMathParaPr/m:jc - display-equation justification. Absent: the
     * document default (settings m:mathPr/m:defJc, itself "centerGroup" when
     * unset). */
    jc?: "left" | "right" | "center" | "centerGroup";
}
interface FieldContent {
    kind: "field";
    /** Raw field instruction, e.g. ` PAGE \\* MERGEFORMAT `. */
    instruction: string;
    /** Last cached result text from the file, used for unsupported fields. */
    cachedResult: string;
    /** When this field is a legacy FORMCHECKBOX, the source `w:checkBox` element
     * (in the fldChar begin's w:ffData). Its presence makes the rendered ballot
     * box a click-to-toggle target; the glyph derives from its w:checked child. */
    checkbox?: XmlElement;
}
/** What an anchored shape's coordinates are measured from. "char" (relH=
 * "character") measures from the anchor run's pen position; "line" (relV=
 * "line") from the top of the line holding the anchor — both resolved from
 * the anchor paragraph's first-pass layout. */
type AnchorRel = "page" | "margin" | "text" | "column" | "char" | "line";
interface ShapeLine {
    type: "line";
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    color: string;
    /** Stroke weight px. */
    weight: number;
    hRel: AnchorRel;
    vRel: AnchorRel;
}
interface ShapeTextbox {
    type: "textbox";
    x: number;
    y: number;
    width: number;
    height: number;
    hRel: AnchorRel;
    vRel: AnchorRel;
    blocks: Block[];
    /** Background fill (CSS color) painted behind the text. */
    fill?: string;
    /** Outline. */
    stroke?: {
        color: string;
        weight: number;
    };
    /** Alignment-based positioning (mso-position-horizontal/vertical). */
    hAlign?: "left" | "center" | "right";
    vAlign?: "top" | "center" | "bottom";
    /** Percent-of-base geometry, 0..1 (mso-*-percent / wp14 pct offsets).
     * Bases: pctWidthRel/pctHeightRel say page vs margin. */
    pctX?: number;
    pctY?: number;
    pctWidth?: number;
    pctHeight?: number;
    pctWidthRel?: "page" | "margin";
    pctHeightRel?: "page" | "margin";
    /** Vertical anchoring of the text INSIDE the box (v-text-anchor). */
    textAnchor?: "top" | "middle" | "bottom";
    /** How body text flows around a DrawingML textbox (wp:wrap*). */
    wrap?: WrapMode;
    /** behindDoc: paint under the body text, never displace it. */
    behind?: boolean;
    /** wp:anchor @allowOverlap="0": Word shifts this shape clear of earlier
     * overlapping floats rather than letting them overlap. Default true. */
    allowOverlap?: boolean;
    /** Wrap distances px (wp:anchor distT/B/L/R). */
    dist?: {
        t: number;
        b: number;
        l: number;
        r: number;
    };
    /** a:xfrm rotation, degrees clockwise (rotates the whole box). */
    rotation?: number;
    /** Text insets px (bodyPr lIns/tIns/rIns/bIns); default 9.6/4.8. For
     * non-rect preset geometries these INCLUDE the geometry's text-rectangle
     * insets (ellipse: inscribed rect; diamond: middle-half rect). */
    insets?: {
        l: number;
        t: number;
        r: number;
        b: number;
    };
    /** Non-rect preset geometry: paint this outline (in a viewW x viewH space
     * scaled to the shape box) instead of a rectangle. */
    geom?: {
        d: string;
        viewW: number;
        viewH: number;
    };
    /** bodyPr a:noAutofit: the box does NOT grow with its text — Word hides
     * whole lines that stick out past the shape bottom (phase23 p12 ovals:
     * "Was 0 / B" visible, the wrapped tail rows hidden). */
    clipText?: boolean;
    /** bodyPr a:spAutoFit: the box grows in height to exactly fit its laid text
     * plus the top/bottom insets (the stored cy is only Word's last cached
     * value). Width is fixed. */
    autofitHeight?: boolean;
    /** Linked text-box chain id (wps:txbx/@id or wps:linkedTxbx/@id): boxes with
     * the same id form one story. The seq-0 box (wps:txbx) holds the content;
     * later boxes (wps:linkedTxbx, chainSeq>0) are empty sinks that continue the
     * overflow. */
    chainId?: string;
    /** Position in the linked chain (0 = the content box). */
    chainSeq?: number;
    /** Source w:drawing element, attached at parse time so the shape's fill can
     * be an interactive hit target (select the shape instead of the body text
     * behind it). */
    srcDrawing?: XmlElement;
    /** Header/footer-owned textbox whose text is edited as an independent story. */
    textboxStory?: boolean;
    /** DrawingML WordArt, including the identity textNoShape preset. */
    wordArt?: boolean;
    /** bodyPr a:prstTxWarp preset name (textArchUp, textWave1, textChevron,
     * textCirclePour, …): the shape's text is bent onto the preset's envelope
     * rather than flowed as ordinary lines. "textNoShape" (no warp) is dropped. */
    warp?: string;
}
/** WordArt (VML v:textpath, e.g. a "CONFIDENTIAL" watermark): text scaled to
 * fill a box, optionally rotated, painted semi-transparent behind the body. */
interface ShapeWordArt {
    type: "wordart";
    text: string;
    fontFamily: string;
    bold?: boolean;
    italic?: boolean;
    /** CSS fill color. */
    fill: string;
    /** 0..1 alpha. */
    opacity: number;
    x: number;
    y: number;
    width: number;
    height: number;
    hRel: AnchorRel;
    vRel: AnchorRel;
    hAlign?: "left" | "center" | "right";
    vAlign?: "top" | "center" | "bottom";
    /** Clockwise degrees. */
    rotation: number;
    behind?: boolean;
    /** The source v:textpath font-size (px). */
    fontSize?: number;
    /** The referenced v:shapetype has a malformed text-guide path (missing
     * coordinates), so Word collapses its horizontal glyph outlines into a thin
     * vertical band instead of painting a box-filling watermark. */
    noFit?: boolean;
    /** Source VML shape element (v:shape / v:rect) for interactive editing:
     * the v:textpath string, fill/opacity, and rotation live under it. */
    src?: XmlElement;
}
interface ShapeArt {
    type: "art";
    srcDrawing?: XmlElement;
    /** A freehand stroke, selectable by the Draw ribbon's stroke eraser. */
    ink?: boolean;
    x: number;
    y: number;
    /** Percent-of-page offsets (wp14:pctPos*Offset), 0..1. */
    pctX?: number;
    pctY?: number;
    width: number;
    height: number;
    hRel: AnchorRel;
    vRel: AnchorRel;
    hAlign?: "left" | "center" | "right";
    vAlign?: "top" | "center" | "bottom";
    behind?: boolean;
    /** a:xfrm rotation, degrees clockwise. */
    rotation?: number;
    lines: DrawingLine[];
    images: DrawingImage[];
    paths: DrawingPath[];
    /** Positioned text bodies inside the group (wps textboxes, dsp:sp txBody). */
    texts?: DrawingTextShape[];
}
/** How text interacts with a floating image. */
type WrapMode = "square" | "topAndBottom" | "none";
interface ShapeImage {
    type: "image";
    part: string;
    /** Offset from the anchor origin, px. */
    x: number;
    y: number;
    width: number;
    height: number;
    hRel: AnchorRel;
    vRel: AnchorRel;
    /** Horizontal alignment when the file uses wp:align instead of an offset. */
    hAlign?: "left" | "center" | "right";
    /** Vertical alignment keyword (VML mso-position-vertical; watermarks center). */
    vAlign?: "top" | "center" | "bottom";
    wrap: WrapMode;
    /** behindDoc anchors render under the text and never displace it. */
    behind?: boolean;
    /** Wrap distances px (wp:anchor distT/B/L/R); text clears the image by these. */
    dist?: {
        t: number;
        b: number;
        l: number;
        r: number;
    };
    crop?: {
        l: number;
        t: number;
        r: number;
        b: number;
    };
    rotation?: number;
    /** Watermark "washout" recolor (VML v:imagedata gain/blacklevel, 0..1). */
    washout?: {
        gain: number;
        blacklevel: number;
    };
    /** Source w:drawing element (editing). */
    srcDrawing?: XmlElement;
    model3D?: Model3DReference;
    webVideo?: WebVideoReference;
    embeddedObject?: EmbeddedObjectReference;
}
type Shape = (ShapeLine | ShapeTextbox | ShapeImage | ShapeArt | ShapeWordArt) & {
    /** wp:anchor relativeHeight, used for object stacking order. */
    z?: number;
};
/**
 * Floating/anchored object: does not occupy inline space; positioned against
 * page/margin/paragraph during layout. (How classic pleading paper draws its
 * margin line numbers and vertical rules.)
 */
interface AnchorContent {
    kind: "anchor";
    shape: Shape;
}
/** A stroked segment inside a composite drawing, px relative to its box. */
interface DrawingLine {
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    color: string;
    weight: number;
    style?: "single" | "dotted" | "dashed";
}
interface DrawingImage {
    part: string;
    x: number;
    y: number;
    width: number;
    height: number;
    crop?: {
        l: number;
        t: number;
        r: number;
        b: number;
    };
    rotation?: number;
    /** a:ln picture outline (Word draws it just outside the image). */
    border?: {
        color: string;
        width: number;
    };
}
/**
 * Composite inline drawing (DrawingML groups): vector lines + placed images
 * inside a width×height box that flows like an image.
 */
interface DrawingContent {
    kind: "drawing";
    width: number;
    height: number;
    lines: DrawingLine[];
    images: DrawingImage[];
    /** Native ChartML chart data resolved from the related chart part. */
    chart?: ChartData;
    /** Native SmartArt diagram data resolved from its diagram parts. */
    smartArt?: SmartArtData;
    /** Freeform vector shapes (a:custGeom), as SVG path data. */
    paths?: DrawingPath[];
    /** An INLINE wps text box (wp:inline wps:txbx): a fixed-extent box that
     * flows in the text (occupying its width x height like an inline image) and
     * carries a fill/border + its own block content. Distinct from the floating
     * ShapeTextbox, which is absolutely placed via a wp:anchor. */
    textbox?: {
        blocks: Block[];
        fill?: string;
        stroke?: {
            color: string;
            weight: number;
        };
        insets?: {
            l: number;
            t: number;
            r: number;
            b: number;
        };
        textAnchor?: "top" | "middle" | "bottom";
    };
    /** Source w:drawing element (select/move as a group). */
    srcDrawing?: XmlElement;
    /** Positioned text bodies inside the drawing (multi-textbox groups and
     * SmartArt cached-drawing shapes). Painted above lines/paths. */
    texts?: DrawingTextShape[];
}
type ChartType = "column" | "bar" | "line" | "pie";
interface ChartSeries {
    name: string;
    values: number[];
}
interface ChartData {
    type: ChartType;
    title?: string;
    categories: string[];
    series: ChartSeries[];
}
type SmartArtLayout = "process" | "cycle" | "hierarchy" | "list";
interface SmartArtData {
    layout: SmartArtLayout;
    items: string[];
}
/** A text body positioned inside a composite drawing (a wps textbox in a
 * group, or a SmartArt dsp:sp txBody), px relative to the drawing's box. */
interface DrawingTextShape {
    x: number;
    y: number;
    width: number;
    height: number;
    blocks: Block[];
    /** Background fill painted behind the text box. */
    fill?: string;
    /** Outline. */
    stroke?: {
        color: string;
        weight: number;
    };
    /** Text insets px (bodyPr lIns/tIns/rIns/bIns). */
    insets: {
        l: number;
        t: number;
        r: number;
        b: number;
    };
    /** Vertical anchoring of the text inside the box (bodyPr anchor). */
    textAnchor?: "top" | "middle" | "bottom";
    /** DrawingML cached text paint-only baseline adjustment in px. */
    paintOffsetY?: number;
}
interface DrawingPath {
    /** Position/size inside the drawing, px. */
    x: number;
    y: number;
    width: number;
    height: number;
    /** SVG path in the `viewW x viewH` source coordinate space. */
    d: string;
    viewW: number;
    viewH: number;
    fill?: string;
    stroke?: {
        color: string;
        width: number;
        opacity?: number;
    };
    /** Zero-based visible SmartArt node this path paints. */
    smartArtNodeIndex?: number;
}
/**
 * Footnote/endnote reference mark. In body text (w:footnoteReference /
 * w:endnoteReference) `id` points into the notes part. Inside a note body,
 * w:footnoteRef / w:endnoteRef render the note's own mark: `self` is true and
 * `id` is meaningless.
 */
interface NoteRefContent {
    kind: "noteRef";
    noteType: "footnote" | "endnote";
    id: number;
    self?: boolean;
    /** The document supplies the visible mark in the following run content. */
    customMarkFollows?: boolean;
}
/** w:ruby — an East-Asian ruby (furigana) cluster: an annotation run (w:rt)
 * riding above a base run (w:rubyBase). Both sides render as their own runs;
 * the annotation is centered over the base and raised above its baseline. */
interface RubyContent {
    kind: "ruby";
    base: Run;
    rt: Run;
    /** w:rubyPr/w:hpsRaise — annotation baseline raise above the base baseline,
     * half-points. */
    hpsRaise?: number;
    /** w:rubyPr/w:rubyAlign — how the annotation distributes over the base. */
    align?: "center" | "distributeLetter" | "distributeSpace" | "left" | "right" | "rightVertical";
}
type RunContent = MathContent | TextContent | BreakContent | TabContent | PTabContent | ImageContent | FieldContent | AnchorContent | DrawingContent | NoteRefContent | RubyContent;
interface Run {
    type: "run";
    props: RunProps;
    content: RunContent[];
    /** Source w:r element. */
    src?: XmlElement;
    /** Element whose children array contains src (w:p, w:hyperlink, …). */
    srcParent?: XmlElement;
}
interface Hyperlink {
    type: "hyperlink";
    href?: string;
    anchor?: string;
    runs: Run[];
}
type ParaChild = Run | Hyperlink;
interface Paragraph {
    type: "paragraph";
    props: ParaProps;
    children: ParaChild[];
    /** Section break attached to this paragraph's pPr (ends a section). */
    sectionBreak?: SectionProps;
    /** w:bookmarkStart names in this paragraph (PAGEREF targets). */
    bookmarks?: string[];
    /** Final revision view: the paragraph mark AND all content are tracked
     * deletions — the paragraph does not exist (no line, no numbering). */
    revisionHidden?: boolean;
    /** Source w:p element. */
    src?: XmlElement;
}
interface TableCellProps {
    /** Preferred width px (from tcW when dxa). */
    width?: number;
    /** The source explicitly declares tcW type="auto" (distinct from no tcW). */
    widthAuto?: boolean;
    gridSpan: number;
    vMerge?: "restart" | "continue";
    borders?: {
        top?: Border;
        bottom?: Border;
        left?: Border;
        right?: Border;
        tl2br?: Border;
        tr2bl?: Border;
    };
    shading?: string;
    /** Cell margins px. */
    margins?: {
        top?: number;
        right?: number;
        bottom?: number;
        left?: number;
    };
    verticalAlign?: "top" | "center" | "bottom";
    textDirection?: "btLr" | "tbRl";
}
interface TableCell {
    props: TableCellProps;
    blocks: Block[];
}
interface TableRowProps {
    /** px; undefined = auto */
    height?: number;
    heightRule?: "auto" | "atLeast" | "exact";
    cantSplit?: boolean;
    /** Repeat as header row on each page. */
    tblHeader?: boolean;
}
interface TableRow {
    props: TableRowProps;
    cells: TableCell[];
}
interface TableProps {
    styleId?: string;
    /** px table indent from left margin. */
    indent?: number;
    alignment?: Alignment;
    /** w:bidiVisual — RTL table: column order reverses and the table hugs the
     * right margin. */
    bidiVisual?: boolean;
    borders?: {
        top?: Border;
        bottom?: Border;
        left?: Border;
        right?: Border;
        insideH?: Border;
        insideV?: Border;
    };
    /** Default cell margins px. */
    cellMargins?: {
        top?: number;
        right?: number;
        bottom?: number;
        left?: number;
    };
    /** w:tblCellSpacing px — old-style separated cell borders: every cell gets
     * its own border box inset by this much from the grid slot on all sides
     * (adjacent cells end up 2x apart), and the table outline is a separate box
     * around everything. */
    cellSpacing?: number;
    /** w:tblpPr — floating table: absolutely positioned against page/margin/
     * column, painted over an opaque white sheet, with body text (including
     * text EARLIER in the flow on the same page) wrapping square around it. */
    floating?: {
        hAnchor: "page" | "margin" | "text";
        vAnchor: "page" | "margin" | "text";
        /** tblpX/tblpY px (absent when an alignment keyword is used). */
        x?: number;
        y?: number;
        xAlign?: "left" | "center" | "right";
        yAlign?: "top" | "center" | "bottom";
        /** Wrap distances px (leftFromText, …). */
        dist: {
            l: number;
            r: number;
            t: number;
            b: number;
        };
        /** w:tblOverlap "never": shift clear of earlier floating tables. */
        allowOverlap: boolean;
    };
    /** Preferred total width px (tblW dxa) or undefined for auto. */
    width?: number;
    /** Preferred total width as a fraction of available width (tblW pct). */
    widthPct?: number;
    layout?: "fixed" | "autofit";
    /** w:tblLook flags controlling which conditional formats apply. */
    tblLook?: TableLook;
}
/** w:tblLook flags controlling which conditional table-style formats apply. */
interface TableLook {
    firstRow: boolean;
    lastRow: boolean;
    firstColumn: boolean;
    lastColumn: boolean;
    noHBand: boolean;
    noVBand: boolean;
}
interface Table {
    type: "table";
    props: TableProps;
    /** Column grid widths px (tblGrid). */
    grid: number[];
    rows: TableRow[];
    /** Source w:tbl element. */
    src?: XmlElement;
}
type Block = Paragraph | Table;
interface HeaderFooterRefs {
    default?: string;
    first?: string;
    even?: string;
}
interface ColumnSpec {
    count: number;
    /** Space between columns px. */
    space: number;
    /** Explicit widths px when equalWidth=false. */
    widths?: number[];
    /** Per-column trailing space px (w:col w:space), when explicit w:col
     * entries are present. spaces[i] separates column i from column i+1. */
    spaces?: number[];
    /** w:cols w:sep: paint a vertical rule centered in each inter-column gap. */
    sep?: boolean;
}
interface SectionProps {
    pageWidth: number;
    pageHeight: number;
    marginTop: number;
    marginRight: number;
    marginBottom: number;
    marginLeft: number;
    /** Distance of header top from page top, px (w:headerReference distance). */
    headerDistance: number;
    footerDistance: number;
    gutter: number;
    headerRefs: HeaderFooterRefs;
    footerRefs: HeaderFooterRefs;
    /** Different first-page header/footer enabled. */
    titlePage: boolean;
    /** Page numbering restart value, if set. */
    pageNumberStart?: number;
    pageNumberFormat?: string;
    columns: ColumnSpec;
    /** w:docGrid type=lines/linesAndChars: minimum single-line height (px) that
     * the line-spacing multiplier is applied over. Word snaps every line's font
     * height up to this grid pitch (CJK documents). */
    docGridLinePitch?: number;
    /** Present w:docGrid type; omitted w:type means "default". */
    docGridType?: "default" | "lines" | "linesAndChars" | "snapToChars";
    /** w:docGrid type="charsAndLines": a combined char+line grid that (in compat
     * 15) keeps natural East-Asian line pitch. Drives the CJK line-height
     * correction during measurement. */
    docGridCharGrid?: boolean;
    type?: "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";
    /** Vertical alignment of page content. */
    vAlign?: "top" | "center" | "both" | "bottom";
    /** w:pgBorders. Offsets (border.space, px) measure from text or page edge. */
    pageBorders?: {
        top?: Border;
        bottom?: Border;
        left?: Border;
        right?: Border;
        offsetFrom: "text" | "page";
    };
    /** w:lnNumType: margin line numbering. distance px from the text edge. */
    lineNumbering?: {
        countBy: number;
        start: number;
        distance: number;
        restart: "continuous" | "newPage" | "newSection";
    };
    /** Footnote/endnote mark numbering (w:footnotePr / w:endnotePr). */
    footnoteNumFmt?: string;
    footnoteNumStart?: number;
    endnoteNumFmt?: string;
    endnoteNumStart?: number;
    /** w:textDirection tbRl: the whole section flows as East-Asian vertical
     * writing — lines run top-to-bottom, progressing right-to-left. */
    textDirection?: "tbRl";
}
interface Section {
    props: SectionProps;
    blocks: Block[];
}
/** A review comment from word/comments.xml. */
interface DocComment {
    id: string;
    author: string;
    initials?: string;
    /** ISO timestamp from w:date, verbatim. */
    date?: string;
    /** Plain text of the comment body (paragraphs joined with newlines). */
    text: string;
    /** w14:paraId of the comment's last body paragraph (threading key). */
    paraId?: string;
    /** Parent comment id when this comment is a reply (commentsExtended). */
    parentId?: string;
}
interface HeaderFooter {
    blocks: Block[];
}
/** One w:tblStylePr conditional-formatting block (firstRow, band1Horz, …). */
interface TableCondFormat {
    /** Cell shading fill as CSS color. */
    shd?: string;
    /** Conditional cell borders. */
    borders?: {
        top?: Border;
        bottom?: Border;
        left?: Border;
        right?: Border;
        insideH?: Border;
        insideV?: Border;
    };
    /** Conditional run bold (firstRow/firstCol headers). */
    bold?: boolean;
    /** Full conditional run props (w:tblStylePr > w:rPr). */
    rPr?: RunProps;
}
type TableCondType = "wholeTable" | "band1Vert" | "band2Vert" | "band1Horz" | "band2Horz" | "firstRow" | "lastRow" | "firstCol" | "lastCol" | "nwCell" | "neCell" | "swCell" | "seCell";
interface Style {
    id: string;
    type: "paragraph" | "character" | "table" | "numbering";
    name?: string;
    basedOn?: string;
    isDefault?: boolean;
    pPr?: ParaProps;
    rPr?: RunProps;
    tblPr?: TableProps;
    /** w:tblStylePr conditional formats, by type (table styles only). */
    condFormats?: Map<TableCondType, TableCondFormat>;
    /** w:tblStyleRowBandSize / ColBandSize (default 1). */
    rowBandSize?: number;
    colBandSize?: number;
}
interface Styles {
    byId: Map<string, Style>;
    defaultParagraphStyle?: string;
    defaultRPr: RunProps;
    defaultPPr: ParaProps;
}
interface NumberingLevel {
    ilvl: number;
    start: number;
    format: string;
    /** Template like "%1." */
    text: string;
    alignment: Alignment;
    /** Paragraph props contributed by the level (indents). */
    pPr?: ParaProps;
    rPr?: RunProps;
    suffix: "tab" | "space" | "nothing";
    restartAfter?: number;
}
interface AbstractNum {
    id: number;
    levels: Map<number, NumberingLevel>;
    numStyleLink?: string;
}
interface NumInstance {
    numId: number;
    abstractNumId: number;
    overrides: Map<number, {
        startOverride?: number;
        level?: NumberingLevel;
    }>;
}
interface Numbering {
    abstract: Map<number, AbstractNum>;
    instances: Map<number, NumInstance>;
}
interface Theme {
    majorFont: string;
    minorFont: string;
    majorBidiFont?: string;
    minorBidiFont?: string;
    /** East Asian faces from theme <a:ea> (fontScheme major/minorFont). Word's
     * eastAsiaTheme="minor/majorEastAsia" resolves through these, NOT the Latin
     * minor/major font — routing a CJK run's theme reference to the Latin face
     * (Calibri) leaves East Asian text in a glyphless Latin font. */
    majorEastAsiaFont?: string;
    minorEastAsiaFont?: string;
    colors: Map<string, string>;
}

/** A read-only view over the OPC (zip) package inside a .docx file. */
declare class Package {
    private files;
    constructor(data: Uint8Array);
    static from(data: ArrayBuffer | Uint8Array): Package;
    has(name: string): boolean;
    binary(name: string): Uint8Array | undefined;
    text(name: string): string | undefined;
    names(): string[];
    /** Raw entry map (shared, do not mutate) — used for write-back. */
    raw(): Record<string, Uint8Array>;
    /** OPC part names never start with '/' inside the zip; tolerate both. */
    private normalize;
}

/**
 * Stable node identity for replicated editing.
 *
 * Positions on the wire cannot be object references, and numeric child paths
 * shift under concurrent structural edits — so blocks (w:p, w:tbl) and runs
 * (w:r) get stable numeric ids held in an in-memory side table. The table is
 * identity-keyed and NEVER serialized into the XML: writing ids as attributes
 * would dirty every part on open and break the byte-identical round-trip
 * guarantee for untouched parts.
 *
 * Assignment is deterministic: a document-order walk over the given roots,
 * so two replicas that parse the same document derive the same table. Nodes
 * created later by edits get ids explicitly (locally via `assign`, or from a
 * replicated edit that carries the originating client's allocation) — a
 * re-walk after edits only fills gaps and never renumbers survivors, because
 * the table is keyed by element identity and in-place XML mutation preserves
 * identity across `refresh()`.
 */
type StableId$1 = number;
interface EncodedCaret {
    blockId: StableId$1;
    runId: StableId$1;
    /** Offset in the run's WIRE space: cumulative through the run's inline
     * content in document order, where each w:t contributes its text length
     * and each inline separator (w:tab / w:br / w:cr) contributes ONE unit.
     * Counting separators makes encoding ONE-TO-ONE: a caret at the end of
     * the w:t before a tab and one at the start of the w:t after it are
     * different physical positions and get different wire offsets — without
     * the separator unit both collapsed to the same number and the decoder
     * had to guess (review bug: boundary intents applied into the wrong w:t
     * on remote replicas, and boundary-starting ranges always rejected). */
    offset: number;
}
declare class StableIds {
    private byEl;
    private byId;
    private next;
    /** Document-order assignment over the given roots (typically
     * `doc.editableRoots()`). Idempotent: elements that already have ids keep
     * them; untracked ids are assigned in walk order. Call after load and
     * after any edit that created tracked nodes without explicit ids. */
    assignFromRoots(roots: XmlElement[]): void;
    /**
     * Assign ids inside the given subtrees only — the scoped counterpart of
     * assignFromRoots, for an edit that is known to have created nodes in one
     * or two paragraphs (dirty-scoped reconciliation, perf B9/B10).
     *
     * EQUIVALENCE (the property the scoped hot path rests on): this produces
     * the same table a full assignFromRoots would, PROVIDED every tracked node
     * outside the subtrees already has an id. Assignment is order-dependent —
     * fresh nodes take sequential numbers from `next` in walk order — and a
     * subtree walk visits exactly a contiguous slice of the document-order
     * walk, in the same order. So the only way the two can disagree is an
     * un-id'd node elsewhere that the full walk would have numbered first.
     * That invariant holds inductively: load/importSidecar ids everything, and
     * every apply either ids the nodes it creates (carried ids), reports
     * document scope (full refresh + full assign), or creates nodes only
     * inside the subtrees passed here. Callers that cannot establish it must
     * fall back to assignFromRoots.
     */
    assignFromSubtrees(subtrees: XmlElement[]): void;
    private walk;
    /** Explicitly assign an id to a newly created node. Local edits allocate
     * (`assign(el)`); applying a replicated edit installs the carried value
     * (`assign(el, carriedId)`). Carried ids must not collide. */
    assign(el: XmlElement, id?: StableId$1): StableId$1;
    /** Force `el` to carry `id`, replacing any id it already has (e.g. one
     * auto-assigned by a refresh() before a carried id could be installed).
     * The previous id is retired. Throws if `id` is held by a different element. */
    reassign(el: XmlElement, id: StableId$1): void;
    private install;
    idOf(el: XmlElement): StableId$1 | undefined;
    elOf(id: StableId$1): XmlElement | undefined;
    /** Drop mappings for elements no longer reachable from the given roots
     * (deleted content). Keeps the table from growing across long sessions.
     * Ids of dropped elements are retired, never reused. */
    prune(roots: XmlElement[]): void;
    size(): number;
    /** Snapshot the table as [id, element] pairs against a parallel clone of
     * the tree: `pairs` maps originals→clones (as produced by walking original
     * and cloned roots in step). Used by the reconciliation snapshot (plan
     * doc 03) so a restore re-keys the table to the restored elements. */
    captureForClone(originalToClone: Map<XmlElement, XmlElement>): Map<StableId$1, XmlElement>;
    /** Replace the table wholesale from a captured mapping (restore path). */
    restore(fromCapture: Map<StableId$1, XmlElement>, nextId: StableId$1): void;
    nextId(): StableId$1;
    /**
     * Export the table as `(id, path)` pairs, where a path is the structural
     * location of the id'd element within the given roots: `[rootIndex,
     * childIndex, childIndex, …]`. This is the checkpoint ID sidecar (plan doc
     * 03): it lets a replica that reloads a snapshot reproduce the exact id
     * table — parse-order re-derivation cannot, because split-created nodes
     * carry non-sequential ids. Also carries `next` so id allocation resumes
     * consistently.
     */
    exportSidecar(roots: XmlElement[]): {
        next: StableId$1;
        entries: [StableId$1, number[]][];
    };
    /** Rebuild the table from a sidecar against freshly parsed roots (the
     * snapshot the sidecar was exported from). Replaces any current mappings. */
    importSidecar(roots: XmlElement[], sidecar: {
        next: StableId$1;
        entries: [StableId$1, number[]][];
    }): void;
    /** Encode a caret as wire-stable addresses. `t` is the w:t (or other text
     * holder) the caret sits in; the run is its nearest tracked ancestor run,
     * the block the nearest tracked block above that. Returns null when the
     * position isn't inside id-tracked content (e.g. math internals) — such
     * positions are not yet addressable on the wire. */
    encodeCaret(t: XmlElement, offset: number, parentOf: (el: XmlElement) => XmlElement | null): EncodedCaret | null;
    /** Resolve an encoded caret back to a concrete (w:t, offset) in THIS
     * table's tree — the inverse of encodeCaret, for restoring a caret across
     * a reconciliation reload (the ids survive via the sidecar; the node
     * references do not). When the exact run is gone (deleted or replaced by
     * a concurrent edit) the caret falls back to the block's first text; null
     * when nothing addressable remains. */
    decodeCaret(pos: EncodedCaret): {
        t: XmlElement;
        offset: number;
    } | null;
}

interface Relationship {
    id: string;
    type: string;
    target: string;
    external: boolean;
}
type Relationships = Map<string, Relationship>;

/**
 * A fully parsed .docx: sections of blocks, styles, numbering, theme, and
 * header/footer parts, with helpers to resolve effective formatting.
 */
declare class DocxDocument {
    /** Changes whenever refresh() rebuilds the parsed model. Plain in-place text
     * edits can keep this stable so incremental layout reuses model-only caches. */
    private _modelVersion;
    private _packageResourceVersion;
    get modelVersion(): number;
    /** Invalidate layout derived from related parts such as ChartML. */
    markPackageResourceChanged(): void;
    readonly pkg: Package;
    readonly theme: Theme;
    styles: Styles;
    numbering: Numbering;
    sections: Section[];
    /** Header/footer parts keyed by relationship id from document.xml.rels. */
    readonly headers: Map<string, HeaderFooter>;
    readonly footers: Map<string, HeaderFooter>;
    /** Note content by note id (render-only; sources stripped). */
    readonly footnotes: Map<number, Block[]>;
    readonly endnotes: Map<number, Block[]>;
    /** The separator paragraph controls the gap between its rule and the first
     * footnote. */
    readonly footnoteSeparator: Block[];
    /** `_Ref` cross-reference bookmark ranges (name → captured runs). REF
     * fields re-render the referenced text from these — Word recomputes REF on
     * open, so the cached field result in the file is stale. */
    refBookmarks: Map<string, Run[]>;
    readonly documentRels: Relationships;
    /** settings.xml w:evenAndOddHeaders — enables the "even" header/footer variants. */
    readonly evenAndOddHeaders: boolean;
    /** settings.xml w:mirrorMargins — facing-page (book fold) margins: even
     * (verso) pages swap the left/right margins and place the gutter on the
     * inside (right) edge so the binding margin stays on the inner side of
     * each spread. */
    readonly mirrorMargins: boolean;
    /** settings.xml w:defaultTabStop in px (Word default 0.5"). */
    readonly defaultTabStop: number;
    /** settings.xml w:compat compatibilityMode (12=Word2007, 14=Word2010,
     * 15=Word2013+). Word 2013 (mode 15) introduced suppressing a paragraph's
     * space-before when it lands at the top of a page; mode 14 and earlier keep
     * it (nccih: a Heading1/2 after a page break sits at margin + its before).
     * Absent → treated as current (15). */
    readonly compatibilityMode: number;
    /** settings.xml m:mathPr/m:defJc — default justification for display
     * equations whose m:oMathParaPr carries no explicit m:jc (Word default:
     * centerGroup — the rows of a broken equation left-align to each other and
     * the group is centered in the column). */
    readonly mathDefJc: "left" | "right" | "center" | "centerGroup";
    /** settings.xml m:mathPr/m:wrapIndent in px (Word default 1440tw = 1"):
     * indent of auto-wrapped display-equation continuation rows from the
     * equation group's left edge (dense p13: the "+Dc(...)" continuations sit
     * exactly 72pt right of the explicit rows). */
    readonly mathWrapIndent: number;
    /** Review comments from word/comments.xml (empty when the part is absent).
     * Re-derived from the retained comments XML on every refresh(). */
    comments: DocComment[];
    /** Retained comments.xml tree (editing + save round-trip), when present. */
    private commentsPart;
    private commentsRoot;
    /** Retained commentsExtended.xml tree (comment threading), when present. */
    private commentsExtPart;
    private commentsExtRoot;
    private commentsExtDirty;
    /** Conditional table formats per table style id, keyed by the Styles object
     * so re-parsing styles.xml (edits) naturally invalidates the cache. */
    private tableCondCache;
    /** Retained styles.xml tree (built-in style injection + save). */
    private stylesPart;
    private stylesRoot;
    /** Retained numbering.xml tree (list creation + save round-trip). */
    private numberingPart;
    private numberingRoot;
    private numberingDirty;
    /** Retained footnotes.xml tree (footnote insertion + save round-trip). */
    private footnotesPart;
    private footnotesRoot;
    private footnotesDirty;
    private footnotesRels;
    /** Serialize retained optional parts only once actually mutated, keeping
     * untouched parts byte-identical through save(). */
    private stylesDirty;
    private commentsDirty;
    /** Retained XML roots — source of truth for editing and save(). */
    private readonly docPart;
    /** Retained settings.xml tree. A synthetic empty root keeps history root
     * indices stable for documents that did not originally contain the part. */
    private readonly settingsPart;
    private readonly settingsRoot;
    private settingsDirty;
    /** Parsed document.xml root (read-only outside the class; the layout engine
     * scans it for incremental-reuse eligibility, tests walk it). */
    readonly docRoot: XmlElement;
    private readonly hfParts;
    private readonly ctxBase;
    /** Tracked-changes display mode; refresh() re-derives after changes. */
    revisionView: "final" | "markup";
    private readonly relsPath;
    private relsRoot;
    private contentTypesRoot;
    /** Canonical XML as first parsed from each always-modeled package part.
     * If the retained tree still matches on save, keep the part's original
     * bytes instead of replacing producer formatting such as CRLF line ends. */
    private readonly originalModeledXml;
    private docPrIdCounter;
    /** Transient layout state: set by the engine while laying out a docGrid
     * type="charsAndLines" section so line measurement can give East-Asian
     * glyphs their true (uninflated) grid line height. Off outside such a
     * section. Safe as document-scoped mutable state because layout is
     * single-threaded and sequential per section. */
    charGridEa: boolean;
    private constructor();
    /** Resolve content controls mapped to standard package core properties.
     * Word refreshes these bindings on open, so the serialized sdtContent can
     * be stale even though the visible value comes from docProps/core.xml. */
    private hydrateCorePropertyControls;
    /** Repair only objects emitted by older WordInWeb builds that Word rejects. */
    private repairLegacyWordInWebObjects;
    /**
     * Re-derive the document model from the retained XML trees. Called after
     * edit commands mutate the XML.
     */
    /** Switch tracked-changes display and re-derive the model. */
    setRevisionView(view: "final" | "markup"): void;
    /** Invalidated on refresh; see layoutGlobalSig. */
    private _layoutGlobalSig;
    /** Signature of everything OUTSIDE a paragraph's own XML that affects how it
     * breaks into lines: style + numbering definitions, doc-level layout scalars,
     * and the tracked-changes view mode. The line-break cache (layout/inline.ts)
     * combines this with a paragraph's own content signature so a style/numbering/
     * settings edit invalidates cached breaks even though the paragraph XML is
     * unchanged. Memoized until the next refresh() (styles/numbering trees are
     * stable across a plain text edit). */
    layoutGlobalSig(): string;
    refresh(): void;
    /** Stable node-id side table for replicated editing. Null (and zero cost)
     * for local-only documents; call `enableStableIds()` to populate and
     * maintain it. Kept in memory only — never serialized into the XML. */
    stableIds: StableIds | null;
    /** Populate the stable-id table from current content and keep it updated
     * on every subsequent refresh(). Idempotent. */
    enableStableIds(): StableIds;
    /** Reparse the two sibling body-story paragraphs created by Enter without
     * rebuilding the complete document model. Paragraphs nested in table cells
     * are included because legal documents spend most of their body inside
     * tables; revisions, bookmarks, fields, and section breaks use refresh(). */
    reparseDirectBodyParagraphSplit(beforeSource: XmlElement, afterSource: XmlElement): {
        before: Paragraph;
        after: Paragraph;
    } | null;
    /**
     * The parsed body-story block list holding the paragraph parsed from
     * `source`, and its index in it — searching sections and, recursively,
     * table cells. Null when the model doesn't hold that paragraph (it lives in
     * a header/footer/footnote part, or the model predates its creation), which
     * is every targeted-reparse helper's signal to fall back to a full
     * refresh(). Shared by those helpers and by paragraphBySource.
     */
    private locateParagraph;
    private _paraLoc;
    /** The parsed Paragraph for a retained w:p in the body story (table cells
     * included), or null when the model doesn't hold it. Lets a caller that
     * already knows which paragraph an edit addresses reach that paragraph's
     * runs without walking the whole model. */
    paragraphBySource(source: XmlElement): Paragraph | null;
    /** Insert a new paragraph immediately before a retained body paragraph
     * without rebuilding the complete document model. Used by Enter at the
     * exact paragraph start, where the existing paragraph itself is unchanged. */
    insertDirectBodyParagraphBefore(referenceSource: XmlElement, insertedSource: XmlElement): Paragraph | null;
    /** Reparse a body paragraph plus several new siblings created by
     * click-and-type without rebuilding the complete document model. */
    reparseDirectBodyParagraphSplits(beforeSource: XmlElement, afterSources: XmlElement[]): Paragraph[] | null;
    /** Reparse one retained body-story paragraph after a structural edit that
     * leaves its surrounding block list unchanged, such as inserting ink. */
    reparseBodyParagraph(source: XmlElement): Paragraph | null;
    /** Reparse two sibling body-story paragraphs after Backspace/Delete merged
     * their XML into one. Keeps the parsed model generation stable so long
     * documents can use incremental layout instead of repaginating in full. */
    reparseDirectBodyParagraphMerge(beforeSource: XmlElement, afterSource: XmlElement, survivorSource: XmlElement): Paragraph | null;
    private deriveComments;
    /** Retained comments tree for edit commands (null when the doc has none). */
    /**
     * Retained comments tree. With create=true, a missing comments.xml part is
     * created and registered (content type + document relationship) so newly
     * added comments serialize and round-trip through Word.
     */
    commentsTree(create?: boolean): XmlElement | null;
    /**
     * Make sure a paragraph style is usable: Word ships built-in definitions
     * for Heading 1-6/Title even when a file doesn't declare them, so applying
     * one to such a file must inject a standard definition (otherwise the
     * paragraph would reference an undefined style and render as Normal).
     */
    ensureParagraphStyle(styleId: string): boolean;
    /**
     * Retained numbering tree. With create=true, a missing numbering.xml part
     * is created and registered (content type + document relationship) so list
     * definitions added by editing serialize and round-trip.
     */
    numberingTree(create?: boolean): XmlElement | null;
    markNumberingChanged(): void;
    /**
     * Retained footnotes tree. With create=true, a missing footnotes.xml part
     * is created and registered (with Word's required separator footnotes) so
     * inserted footnotes serialize and round-trip.
     */
    footnotesTree(create?: boolean): XmlElement | null;
    /**
     * Create an empty header/footer part (with a default-type reference in
     * every sectPr) when the document has none - Word does this implicitly the
     * first time you edit the header area. Returns the part's root.
     */
    /** Whether the document declares any section properties at all. A blank
     * minimal document has none; see ensureHfPart. */
    private hasSectPr;
    /** Whether a header (or footer) part already exists — the precondition
     * `ensureHfPart` tests internally, exposed so a caller can tell a real
     * creation from a no-op before mutating (the collab apply needs to know
     * whether an ensureHeaderFooter intent has anything to do). */
    hasHfPart(kind: "header" | "footer"): boolean;
    ensureHfPart(kind: "header" | "footer"): XmlElement;
    markFootnotesChanged(): void;
    /** Called by comment edit commands after mutating the comments tree. */
    markCommentsChanged(): void;
    /**
     * Retained commentsExtended tree (threading). With create=true, a missing
     * part is created and registered (content type + document relationship) so
     * Word picks up reply threading.
     */
    commentsExtendedTree(create?: boolean): XmlElement | null;
    markCommentsExtendedChanged(): void;
    /**
     * The w:t elements covered by each comment's range, in document order.
     * Point comments (a bare commentReference with no range) anchor to the
     * nearest preceding w:t.
     */
    commentAnchors(): Map<string, XmlElement[]>;
    /** Flag the footnotes part dirty when `t` lives inside it, so save()
     * re-serializes footnotes.xml. Called by the editor after a text edit; a
     * no-op for body/header/footer targets. */
    markDirtyIfFootnote(t: XmlElement): void;
    /** The mutable XML roots (document body, related modeled parts, settings).
     * settingsRoot is always second and always present so its history snapshot
     * index stays stable even when optional related roots are created later. */
    editableRoots(): XmlElement[];
    /** Toggle the document-global facing-page margin mode in settings.xml. */
    setMirrorMargins(enabled: boolean): void;
    /**
     * Find the parent element of `target` in any modeled XML tree (document
     * body, headers, footers). Linear scan — documents are small and this only
     * runs on structural edits (Enter, paragraph merge).
     */
    /** XML roots that can carry tracked changes: body, headers/footers, footnotes. */
    revisionRoots(): XmlElement[];
    /**
     * Memoized child→parent links. Nothing invalidates this map — the tree is
     * spliced constantly — so every answer is RE-DERIVED before it is returned:
     * the memoized parent must still list the target among its children, and
     * must still hang off a root (see memoIsLive). A stale entry can then only
     * cost a cache miss, never a wrong answer.
     *
     * Both halves are needed. Containment alone is not proof of a live parent:
     * a run split for formatting is spliced OUT of its paragraph while its own
     * `children` array still lists the elements that moved into the replacement
     * runs, so a containment-only check would hand back the detached run where a
     * walk from the roots finds the new one.
     */
    private _parentMemo;
    private memoIsLive;
    findParentOf(target: XmlElement): XmlElement | undefined;
    /** Record a parent link the caller just created (a splice), so the next
     * lookup skips the full-document walk. Advisory only — findParentOf
     * verifies every memo against the live tree before trusting it. */
    noteParent(child: XmlElement, parent: XmlElement): void;
    /**
     * Serialize the (possibly edited) document back to .docx bytes. Only the
     * XML parts we model are re-serialized; every other part round-trips
     * byte-for-byte.
     */
    private rememberOriginalXml;
    private writeModeledXml;
    /**
     * Canonicalize producer shorthand that Google Docs otherwise interprets as
     * a fixed, few-twip table. Word treats `tblW="100%"` plus a placeholder
     * grid as autofit; Google needs the standard pct value and a usable cached
     * grid. The cached widths follow the same content-dominant shape and do not
     * change Word's autofit result.
     */
    /** Undo closures recorded by save-time fixups while `saveJournal` is
     * active, so `save()` can revert every live-tree mutation and remain
     * side-effect-free. Collaborative checkpoints re-serialize the
     * authoritative document repeatedly; a save that mutated the tree would
     * change its hash outside the intent stream and desync the fleet. */
    private saveJournal;
    private journalSetAttr;
    private normalizePercentageTableGrids;
    save(): Uint8Array;
    saveAsync(): Promise<Uint8Array>;
    private buildPackage;
    private buildPackageFiles;
    /** Fresh unique docPr id for inserted drawings. Seeded once past the
     * highest id already present in any editable root (floor 1000, matching
     * the historical seed for documents with no drawings) so a document that
     * already carries drawings never gets a colliding id. */
    nextDrawingId(): number;
    /** Next unused revision id (w:id on w:ins/w:del). Seeded once past the
     * highest id already present in any editable root so a document that
     * already has tracked changes never collides. */
    private revIdCounter;
    nextRevisionId(): number;
    /**
     * Add image bytes as a new media part + relationship (+ content-type
     * default). Returns the relationship id for use in a w:drawing.
     */
    private ensureRelsRoot;
    /** Register an external hyperlink relationship and return its rId. */
    addHyperlinkRel(url: string): string;
    /** Retarget an existing external relationship (hyperlink href edit). */
    setRelTarget(relId: string, url: string): boolean;
    addImageResource(bytes: Uint8Array, ext: string): string;
    /** Add a GLB model part and its Office 2019 model3d relationship. */
    addModel3DResource(bytes: Uint8Array): {
        relId: string;
        part: string;
    };
    /** Add an OLE package part used by a Word w:object. */
    addEmbeddedObjectResource(bytes: Uint8Array): {
        relId: string;
        part: string;
    };
    /** Add a DOCX package embedded as an activatable Word.Document.12 object. */
    addEmbeddedWordDocumentResource(bytes: Uint8Array): {
        relId: string;
        part: string;
    };
    /** Add a native ChartML part and its embedded editable workbook. */
    addChartResource(chartXml: string, workbook: Uint8Array): {
        relId: string;
        part: string;
    };
    /** Add the SmartArt data/layout/style/color parts and its cached diagram drawing. */
    addSmartArtResources(layoutXml: string, styleXml: string, colorsXml: string, drawingXml: string, dataXml: (drawingRelId: string) => string): {
        dataRelId: string;
        layoutRelId: string;
        styleRelId: string;
        colorsRelId: string;
        drawingRelId: string;
    };
    static load(data: ArrayBuffer | Uint8Array): DocxDocument;
    media(part: string): Uint8Array | undefined;
    /**
     * Pending-media registry (plan doc 05 change 1 / doc 16 §6): parts whose
     * XML registration (rels + content-type + extent geometry) exists but
     * whose BYTES have not arrived yet — the out-of-band media model. Layout
     * needs no bytes (extents live in the XML), so a doc with pending parts
     * lays out pixel-identically; the renderer shows a placeholder until
     * `installMedia`. Keyed by part name; the value is the doc-16 §5.3
     * metadata (declared sha + optional E2EE iv/epoch for re-supply).
     */
    readonly pendingMedia: Map<string, {
        sha: string;
        iv?: string;
        genesisId?: string;
    }>;
    /** DISPLAY-ONLY transfer state for a pending part (doc 16 §5.2 step 4), so
     * the skeleton can say "fetching" versus "nobody online has this yet".
     * Written by the media transfer layer, read by the renderer; never part of
     * the document's identity and never serialized. */
    readonly mediaTransferState: Map<string, "unavailable" | "fetching" | "waiting">;
    /** Per-part media metadata that PERSISTS after install (doc 16 §5.3):
     * holder duty needs the sha (and E2EE iv/epoch) of READY parts to answer
     * re-supply requests; pendingMedia above only tracks not-yet-arrived. */
    readonly mediaMeta: Map<string, {
        sha: string;
        iv?: string;
        genesisId?: string;
    }>;
    /** "ready" = bytes present; "pending" = registered, bytes absent (doc 05).
     * Unregistered parts report "pending" too — a rel pointing at a missing
     * zip entry after a save/parse round-trip IS the pending state (a hole is
     * detectable, never corrupt — doc 16 §6 round-trip rule). */
    mediaStatus(part: string): "ready" | "pending";
    /**
     * Register an image part + relationship WITHOUT bytes (doc 16 §2: the
     * intent carries the sha; bytes travel out of band). Same deterministic
     * naming/rId scan as addImageResource so every replica applying the same
     * canonical intent derives identical registration.
     */
    registerPendingImage(sha: string, ext: string, meta?: {
        iv?: string;
        genesisId?: string;
    }): string;
    /** Install fetched bytes into a pending part (doc 05). The caller has
     * ALREADY verified the sha against the intent's declaration (the
     * reservation is the hash commitment — doc 16 §1.1); this installs and
     * clears the pending record. Returns false for unknown parts. */
    installMedia(part: string, bytes: Uint8Array): boolean;
    /** Effective paragraph properties: docDefaults → table style → style chain → direct. */
    effectiveParaProps(para: Paragraph): ParaProps;
    /**
     * Run props contributed by the enclosing table style's conditional
     * w:tblStylePr blocks for this paragraph's cell (undefined when the
     * paragraph isn't in a styled table cell or nothing applies).
     */
    private tableCondRPr;
    /** Effective run properties for a run inside a paragraph. */
    effectiveRunProps(para: Paragraph, runProps: RunProps): RunProps;
    numberingLevel(numId: number, ilvl: number): NumberingLevel | undefined;
    numberingInstance(numId: number): NumInstance | undefined;
    private findDocumentPart;
    private readXmlOptional;
}

/** Maps a rendered text item back to its source XML for editing. */
interface TextSource {
    run: Run;
    /** Source w:t; null when the text is synthetic (fields, symbols). */
    t: XmlElement | null;
    /** Char offset of this item's first character within t's text. */
    offset: number;
}
/**
 * Layout output: pages of absolutely positioned primitives (px, page-relative).
 * The renderer maps these 1:1 to DOM/canvas/SVG — no further layout happens
 * downstream, which is what guarantees pagination fidelity.
 */
interface FontSpec {
    family: string;
    size: number;
    bold: boolean;
    italic: boolean;
    /** Pair kerning is active because the run met its w:kern size threshold. */
    kerning?: boolean;
    /** Optional PAINT-only family tried before `family` in the CSS stack, without
     * affecting metrics or width lookups (keyed by `family`). Used for CJK: the
     * line-pitch profile stays keyed by the macOS substitute face while the real
     * Windows glyphs (MS Mincho, Microsoft JhengHei, …) paint when available. */
    paintFamily?: string;
    /** PAINT-only vertical nudge (CSS px, positive = down) applied to the glyphs
     * without touching layout metrics or advances. Used when a substitute face
     * (Latha for Word's Tamil Vijaya) sits on a different baseline than the face
     * Word rasterized; the shift lands the glyphs on Word's baseline. */
    paintDY?: number;
}
interface PathItem {
    kind: "path";
    x: number;
    y: number;
    width: number;
    height: number;
    /** SVG path data in a `viewW x viewH` coordinate space. */
    d: string;
    viewW: number;
    viewH: number;
    fill?: string;
    stroke?: {
        color: string;
        width: number;
        opacity?: number;
    };
    rotate?: {
        deg: number;
        ox: number;
        oy: number;
    };
    /** Paint under / over the body text (anchored-shape layering). */
    behind?: boolean;
    front?: boolean;
    z?: number;
}
interface TextItem {
    kind: "text";
    x: number;
    /** Baseline y. */
    baseline: number;
    width: number;
    text: string;
    props: RunProps;
    font: FontSpec;
    /** Footnote/endnote id referenced by this run (registration happens when
     * the item lands on a real page - split table rows carry it across). */
    noteId?: number;
    /** Vertical stretch for tall delimiter glyphs (Word's glyph variants). */
    mathScaleY?: number;
    /** Horizontal stretch for wide brace glyphs (Word's over/under brace
     * variants), applied about the glyph box's horizontal center. */
    mathScaleX?: number;
    /** Stretch anchor above the baseline, px. */
    mathScaleAnchor?: number;
    /** Line box for selection/highlight backgrounds. */
    lineTop: number;
    lineHeight: number;
    /** Exact glyph box for baseline-shifted runs (superscript/subscript):
     * the renderer anchors these instead of bottoming on the line box. */
    glyphTop?: number;
    glyphBoxH?: number;
    /** Small-caps reduced segment: the base run font that must supply the
     * strut so the painted baseline matches neighboring full-size spans (the
     * renderer sizes the outer span with this font and shrinks the text via
     * a baseline-aligned inner span). */
    strutFont?: FontSpec;
    /** PAGEREF bookmark name: the final pass rewrites this item's text with
     * the bookmark's page number (Word recomputes PAGEREF on open; the docx
     * cached result is stale in real TOCs). */
    pageRef?: string;
    /** DOCX bookmarks that begin on this text item. */
    bookmarks?: string[];
    /** Source m:oMath element when this text is a piece of an equation. */
    mathSrc?: XmlElement;
    href?: string;
    /** Present for editable text (absent on numbering labels etc.). */
    src?: TextSource;
    /** Source drawing when this text belongs to an independently editable
     * text-box story. */
    textboxStory?: XmlElement;
    /** Rotate about a point (px, relative to this item's top-left). */
    rotate?: {
        deg: number;
        ox: number;
        oy: number;
    };
    /** Paint under the body text (behindDoc textbox content). */
    behind?: boolean;
    /** Paint over the body text (non-behindDoc anchored shape content: Word
     * layers in-front shapes above the text layer). */
    front?: boolean;
    z?: number;
    /** Right-to-left run: renderer sets direction:rtl so the browser shapes and
     * orders the (Arabic/Hebrew) glyphs within the span box. */
    rtl?: boolean;
    /** Ordinal of the source paragraph within its cell/frame. Set only for
     * widow-controlled paragraphs; row splitting uses it to scope Word's
     * widow/orphan rules to the paragraph straddling the cut. */
    paraSeq?: number;
    /** Editor-only: maximum caret x for this item — a trailing hanging space
     * confined to its table cell pins the caret at the cell's content edge
     * (Word behavior) while the span keeps its true hanging layout x. */
    caretClampX?: number;
}
interface RectItem {
    kind: "rect";
    x: number;
    y: number;
    width: number;
    height: number;
    fill: string;
    /** Semantic paint provenance used by render-layer comparison. */
    role?: "table-fill";
    /** Rotate about a point (px, relative to this item's top-left). */
    rotate?: {
        deg: number;
        ox: number;
        oy: number;
    };
    behind?: boolean;
    /** Paint over the body text (non-behindDoc anchored shape fill). */
    front?: boolean;
    z?: number;
}
interface LineEdgeItem {
    kind: "edge";
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    border: Border;
    /** Semantic paint provenance used by render-layer comparison. */
    role?: "table-rule";
    /** Rotate about a point (px, relative to this item's top-left). */
    rotate?: {
        deg: number;
        ox: number;
        oy: number;
    };
    /** Paint over the body text (non-behindDoc anchored shape border). */
    front?: boolean;
    /** Paint below the body text (behindDoc anchored shape border). */
    behind?: boolean;
    z?: number;
}
interface ImageItem {
    kind: "image";
    x: number;
    y: number;
    width: number;
    height: number;
    /** Package part path; renderer resolves bytes via DocxDocument.media(). */
    part: string;
    /** a:srcRect crop (fractions) and a:xfrm rotation (degrees). */
    crop?: {
        l: number;
        t: number;
        r: number;
        b: number;
    };
    rotation?: number;
    /** Whole-group rotation about an external origin. */
    rotate?: {
        deg: number;
        ox: number;
        oy: number;
    };
    /** Picture-watermark "washout" (VML v:imagedata gain/blacklevel, 0..1
     * fractions). Per-channel linear recolor, measured against Word's PDF of
     * probe2-picture-watermark (gain 0.3, blacklevel 0.35: source 32 -> 215,
     * 74 -> 227, 135 -> 246, 210 -> clamped 255):
     *   out = in * gain + 255 * (blacklevel * (1 + gain) + (1 - gain) / 2)  */
    washout?: {
        gain: number;
        blacklevel: number;
    };
    /** a:ln picture outline, drawn just outside the image (Word hairline). */
    border?: {
        color: string;
        width: number;
    };
    /** behindDoc: paint under the text layer. */
    behind?: boolean;
    /** wrapNone + !behindDoc ("in front of text"): paint ABOVE the text layer
     * (Word's z-order; without this, later-emitted text spans cover the image
     * and it neither shows in front nor receives clicks/drags). */
    front?: boolean;
    z?: number;
    /** Source w:drawing element (for interactive resize/move). */
    src?: XmlElement;
    model3D?: Model3DReference;
    webVideo?: WebVideoReference;
    embeddedObject?: EmbeddedObjectReference;
}
interface ChartItem {
    kind: "chart";
    x: number;
    y: number;
    width: number;
    height: number;
    data: ChartData;
}
/** Interactive resize zone over a table boundary (column or row). */
interface DrawingHitItem {
    kind: "drawingHit";
    x: number;
    y: number;
    width: number;
    height: number;
    /** Source w:drawing element (select/move the whole drawing). */
    src: XmlElement;
    /** Anchored drawings drag by offset; inline ones re-anchor into text. */
    anchored: boolean;
    /** Freehand stroke that may be removed with the Draw ribbon's eraser. */
    ink?: boolean;
    /** Native SmartArt diagram selection target. */
    smartArt?: boolean;
    /** Node hit boxes inside the SmartArt drawing, relative to this item. */
    smartArtNodes?: Array<{
        index: number;
        x: number;
        y: number;
        width: number;
        height: number;
    }>;
    chartData?: ChartData;
    smartArtData?: SmartArtData;
    /** A shape fill hit that sits UNDER the shape's own text spans (same z): a
     * click on bare fill selects or drags the shape, while clicks on its text
     * glyphs still reach the text editor. */
    belowText?: boolean;
    /** Keep a behindDoc drawing target under the body text. */
    behind?: boolean;
    /** The drawing owns an independently editable text-box story. */
    textboxStory?: boolean;
    rotate?: {
        deg: number;
        ox: number;
        oy: number;
    };
    z?: number;
}
interface GripItem {
    kind: "grip";
    /** "col": vertical zone at x spanning y1..y2. "row": horizontal zone at y1
     * spanning x..x2. "move": table bounds x..x2, y1..y2. */
    axis: "col" | "row" | "move";
    x: number;
    y1: number;
    y2: number;
    /** Right edge for row grips. */
    x2?: number;
    /** Source w:tbl element. */
    tbl: XmlElement;
    /** col: boundary 1..n (n = right edge). row: row index above the boundary. */
    boundary: number;
    /** Laid-out height of the row above (row grips), px. */
    rowHeightPx?: number;
    /** Rendered column widths px (col grips) — resize works in this space. */
    renderedWidths?: number[];
}
/** WordArt/watermark text scaled to fill a box, rotated about its center. */
interface WordArtItem {
    kind: "wordart";
    /** Box top-left, px. */
    x: number;
    y: number;
    width: number;
    height: number;
    text: string;
    fontFamily: string;
    bold?: boolean;
    italic?: boolean;
    fill: string;
    opacity: number;
    /** Clockwise degrees. */
    rotation: number;
    behind?: boolean;
    /** Source v:textpath font-size (px). */
    fontSize?: number;
    /** Malformed shapetype guide path: Word collapses the fitted glyph outlines
     * into a thin vertical band. */
    noFit?: boolean;
    /** Source VML shape element (v:shape / v:rect), for interactive select/edit. */
    src?: XmlElement;
}
/** DrawingML a:prstTxWarp text: the shape's single text line bent onto a
 * preset envelope (arch/wave/chevron/circle-pour), filling the shape box. */
interface WarpTextItem {
    kind: "warptext";
    /** Shape box top-left, px. */
    x: number;
    y: number;
    width: number;
    height: number;
    text: string;
    fontFamily: string;
    /** Run font size px (the text's natural size; presets that don't fill the
     * box — textArchUp — ride the text at this size). */
    fontSize: number;
    bold?: boolean;
    italic?: boolean;
    /** CSS text color. */
    fill: string;
    /** Preset name (textArchUp, textWave1, textChevron, textCirclePour, …). */
    warp: string;
    rotate?: {
        deg: number;
        ox: number;
        oy: number;
    };
    behind?: boolean;
    front?: boolean;
    z?: number;
}
type PageItem = TextItem | PathItem | RectItem | LineEdgeItem | ImageItem | ChartItem | DrawingHitItem | GripItem | WordArtItem | WarpTextItem;
interface LaidOutPage {
    width: number;
    height: number;
    /** 1-based physical index. */
    index: number;
    /** Display page number after pgNumType.start is applied. */
    number: number;
    items: PageItem[];
    /** Body box (for header/footer editing chrome). */
    bodyTop: number;
    bodyBottom: number;
    /** Items from this index on belong to the header/footer parts. */
    hfStart: number;
    /** Column geometry for editor hit-testing. A page can contain more than one
     * band when a continuous section changes the column layout mid-page. */
    columnBands: Array<{
        top: number;
        colXs: number[];
        colWidths: number[];
    }>;
}
interface LayoutResult {
    pages: LaidOutPage[];
    totalPages: number;
    /** Opaque incremental-reuse state (engine-internal); pass a previous result
     * back as LayoutOptions.prev to enable incremental relayout. Not part of the
     * rendered output and ignored by the equivalence harness. */
    _incr?: unknown;
    /** Internal marker: unchanged pages in this result retain exact object
     * identity from the previous layout, so the DOM diff can use identity
     * instead of structurally walking a dense changed page. */
    _incremental?: boolean;
    /** Exclusive global page index through which an incremental render may run
     * pageEq before the page containing the exact dirty text source. */
    _incrementalStructuralPrefixEnd?: number;
    /** Opaque page state used to refresh headers/footers without repaginating
     * the body when their measured geometry is unchanged. */
    _hf?: unknown;
}

interface FontMetrics {
    /** px above baseline (raw font ascent - anchors CSS glyph boxes). */
    ascent: number;
    /** px below baseline, raw (positive - anchors CSS glyph boxes). */
    descent: number;
    /** Natural single-spaced line height px (what Word calls single spacing). */
    lineHeight: number;
    /** Word's below-baseline share of lineHeight (quantized); use for
     * baseline placement. Falls back to `descent` when absent. */
    lineDescent?: number;
}
interface TextMeasurer {
    width(text: string, font: FontSpec, letterSpacing?: number): number;
    metrics(font: FontSpec): FontMetrics;
    /** The browser's own font box (fontBoundingBoxAscent/Descent) for the
     * resolved first face of this font's CSS stack, when the host can know it.
     * Used to anchor paint-routed CJK spans: the DOM renderer centers glyphs
     * by the browser strut, which differs from the engine's calibrated line
     * profile (Hiragino/PingFang) whenever the real Windows face paints. */
    paintBox?(font: FontSpec): {
        ascent: number;
        descent: number;
    } | undefined;
    /** Ink extents (actualBoundingBox*) of `text` in the RESOLVED paint face,
     * px above/below the baseline. The radical join needs the √ glyph's real
     * ink — which face paints it (real Cambria Math vs the STIX substitute)
     * changes the extents, so a baked table can't cover both. */
    inkBox?(text: string, font: FontSpec): {
        ascent: number;
        descent: number;
    } | undefined;
}

interface LayoutOptions {
    measurer?: TextMeasurer;
    /** Previous layout result (from an earlier layoutDocument call on the same
     * document). Enables incremental relayout: pages whose input blocks and
     * lead-in state are unchanged are reused instead of re-laid. The engine falls
     * back to a full layout whenever it cannot prove reuse is byte-identical. */
    prev?: LayoutResult;
    /** The top-level block XML element (w:p / w:tbl) the editor mutated IN PLACE
     * since `prev` — the paragraph the caret sits in for a single-character
     * type/delete. Lets the incremental scan skip re-hashing every block: it
     * re-hashes only the hinted block and its two neighbours and reuses prev's
     * per-block signatures for the rest. Purely an optimisation — the fast path
     * is gated on block identity/count and neighbour-signature checks, so a stale
     * or wrong hint falls through to the full block scan. */
    dirtyHint?: XmlElement;
    /** Exact retained text element changed by a local edit. Used only to bound
     * structural page comparison before the first actually dirty page. */
    dirtySource?: XmlElement;
}
declare function layoutDocument(doc: DocxDocument, options?: LayoutOptions): LayoutResult;

interface RenderOptions {
    /** Zoom factor (1 = 100%). */
    zoom?: number;
    /** Gap between pages, px. */
    pageGap?: number;
    /** Page drop shadow / chrome. */
    pageShadow?: boolean;
    /** Materialize interactive affordances (table resize grips). */
    interactive?: boolean;
    /** Keep only nearby page contents mounted. Page shells always remain so
     * scroll geometry is unchanged. Intended for long interactive documents. */
    virtualize?: boolean;
    /** Called after the mounted page window changes so editor chrome can be
     * restored on newly-mounted pages. */
    onViewportChange?: () => void;
    /** Show review comments (highlight + margin balloons). Default true. */
    comments?: boolean;
    /** Called when the user deletes a comment from its balloon. The balloon
     * shows a delete button only when this is provided. */
    onDeleteComment?: (id: string) => void;
    /** Called when the user submits a reply from a balloon's reply box. The
     * reply box only renders when this is provided. */
    onReplyComment?: (id: string, text: string) => void;
}
interface TextBinding {
    el: HTMLElement;
    item: TextItem;
}
interface GripBinding {
    el: HTMLElement;
    item: GripItem;
}
interface ImageBinding {
    el: HTMLElement;
    item: ImageItem;
}
interface DrawingBinding {
    el: HTMLElement;
    item: DrawingHitItem;
}
interface WordArtBinding {
    el: HTMLElement;
    item: WordArtItem;
}
interface RenderHandle {
    /** Root element containing all pages. */
    root: HTMLElement;
    /** Rendered text elements in paint order, for selection mapping. */
    bindings: TextBinding[];
    /** Mounted text bindings indexed by their retained source w:t element. */
    bindingsByText: Map<XmlElement, TextBinding[]>;
    /** Table resize grips (only when options.interactive). */
    grips: GripBinding[];
    /** Rendered images, for interactive select/resize/move. */
    images: ImageBinding[];
    /** Transparent hit targets over vector drawings/icons (select/move). */
    drawings: DrawingBinding[];
    /** WordArt / text watermarks, for interactive select/edit. */
    wordarts: WordArtBinding[];
    /** Revoke object URLs etc. */
    destroy: () => void;
    /** Temporarily mount every page, returning a function that restores the
     * viewport window. Used by the synchronous print-clone path. */
    materializeAll?: () => () => void;
    /** Recompute the mounted page window after an external scroll/resize. */
    updateViewport?: () => void;
    /** Per-page render records, retained so the next render can reuse the DOM of
     * pages whose layout is unchanged (see renderToDom's `prev` parameter). */
    _pages?: PageRender[];
    /** Whether this render drew the comments overlay. Comment highlights live
     * INSIDE page surfaces, so the next render must sweep them out of any page
     * DOM it adopts before re-running the overlay. */
    _hadComments?: boolean;
    /** Zoom this render painted at — adoption is only valid at the same zoom. */
    _zoom?: number;
    /** Whether this handle keeps only a viewport-sized page window mounted. */
    _virtualized?: boolean;
    /** Parsed-model generation whose source bindings this handle owns. */
    _modelVersion?: number;
    /** Remove viewport listeners before this handle is replaced. */
    _stopVirtualizer?: () => void;
}
/** One page's DOM element plus the editor bindings it owns and the object URLs
 * it created. Retained on the handle so an incremental re-render can adopt an
 * unchanged page wholesale instead of tearing it down and rebuilding it. */
interface PageRender {
    el: HTMLElement;
    page: LaidOutPage;
    bindings: TextBinding[];
    grips: GripBinding[];
    images: ImageBinding[];
    drawings: DrawingBinding[];
    wordarts: WordArtBinding[];
    urls: string[];
    mounted: boolean;
}
declare function renderToDom(doc: DocxDocument, layout: LayoutResult, container: HTMLElement, options?: RenderOptions, prev?: RenderHandle): RenderHandle;
/**
 * Print the rendered pages (browser print -> paper or PDF): clones the page
 * DOM into a hidden same-origin iframe sized to the document's page, strips
 * screen chrome (shadows, gaps), and invokes the print dialog.
 */
declare function printPages(root: HTMLElement, pageWidthPx: number, pageHeightPx: number): void;

interface RunFormatPatch {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    /** "#RRGGBB"; null removes the direct color. */
    color?: string | null;
    /** Word highlight name ("yellow", "cyan", …); null removes. */
    highlight?: string | null;
    /** Font size in points. */
    fontSizePt?: number;
    fontFamily?: string;
    /** Superscript/subscript; null returns the run to the baseline. */
    verticalAlign?: "superscript" | "subscript" | null;
    /** Remove all direct character formatting (and character style). */
    clear?: boolean;
}
interface SelectionFormat {
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strike: boolean;
    /** Common font size in points, if uniform. */
    fontSizePt?: number;
    color?: string;
    highlight?: string;
    fontFamily?: string;
    verticalAlign?: "superscript" | "subscript";
}

type ShapePreset = "line" | "verticalLine" | "rectangle" | "roundedRectangle" | "ellipse" | "diamond" | "textBox";
type WordArtPreset = "plain" | "archUp" | "archDown" | "wave" | "chevron";
type DrawingLineDash = "solid" | "dashed" | "dotted";
type DrawingTool = {
    kind?: "pen";
    color: string;
    width: number;
} | {
    kind: "highlighter";
    color: string;
    width: number;
} | {
    kind: "eraser";
    size: number;
} | {
    kind: "lasso";
};

interface SmartArtTextFormat {
    fontFamily: string;
    fontSizePt: number;
    color: string;
    bold: boolean;
    italic: boolean;
    alignment: "left" | "center" | "right";
}

type ObjectArrangeAction = "alignLeft" | "alignCenter" | "alignRight" | "alignTop" | "alignMiddle" | "alignBottom" | "rotateLeft" | "rotateRight" | "bringToFront" | "sendToBack";
type SelectedObjectKind = "shape" | "line" | "smartArt" | "chart" | "image" | "model3d";
type SelectedObjectCommand = "altText" | "editText" | "fill" | "outline" | "lineStyle" | "rotate" | "size" | "position" | "wrapInline" | "wrapSquare" | "wrapTopAndBottom" | "wrapFront" | "wrapBehind" | "bringForward" | "sendBackward" | "reset3d" | "delete";
/** A local edit expressed as a replicable intent, emitted by the editor for
 * the collab layer. Positions are stable ids; the connection adds wire
 * bookkeeping (clientId/clientSeq/base). Mirrors the @wordinweb/collab intent
 * shapes without a dependency on that package. */
type EditorIntent = {
    kind: "insertText";
    at: {
        blockId: number;
        runId: number;
        offset: number;
    };
    text: string;
    suggest?: {
        author: string;
        date: string;
    };
} | {
    kind: "deleteText";
    blockId: number;
    runId: number;
    start: number;
    end: number;
} | {
    kind: "splitParagraph";
    at: {
        blockId: number;
        runId: number;
        offset: number;
    };
    newBlockId: number;
    newRunId: number;
} | {
    kind: "formatRun";
    blockId: number;
    runId: number;
    patch: Record<string, unknown>;
} | {
    kind: "formatRange";
    blockId: number;
    runId: number;
    start: number;
    end: number;
    patch: Record<string, unknown>;
    beforeId?: number;
    middleId: number;
    afterId?: number;
} | {
    kind: "formatParagraph";
    blockId: number;
    align?: "left" | "center" | "right" | "justify";
    styleId?: string | null;
} | {
    kind: "setListType";
    blockId: number;
    listKind: "bullet" | "number" | null;
} | {
    kind: "mergeParagraph";
    blockId: number;
} | {
    kind: "suggestRevision";
    ranges?: {
        blockId: number;
        runId: number;
        start: number;
        end: number;
    }[];
    marks?: {
        blockId: number;
        glyph: "ins" | "del";
    }[];
    suggest: {
        author: string;
        date: string;
    };
} | {
    kind: "setImageWrap";
    runId: number;
    mode: "inline" | "square" | "topAndBottom" | "none" | "behind";
} | {
    kind: "setFloatingPagePosition";
    runId: number;
    xPx: number;
    yPx: number;
} | {
    kind: "resizeDrawing";
    runId: number;
    widthPx: number;
    heightPx: number;
} | {
    kind: "setDrawingRotation";
    runId: number;
    degrees: number;
} | {
    kind: "setDrawingOrder";
    runId: number;
    order: "front" | "back";
} | {
    kind: "setDrawingLineStyle";
    runId: number;
    color: string;
    widthPx: number;
    dash: "solid" | "dashed" | "dotted";
} | {
    kind: "setImageAltText";
    runId: number;
    alt: string;
} | {
    kind: "setDrawingFill";
    runId: number;
    color: string | null;
} | {
    kind: "setSmartArtFill";
    runId: number;
    color: string | null;
    nodeIndex?: number;
} | {
    kind: "setSmartArtNodeText";
    runId: number;
    index: number;
    text: string;
} | {
    kind: "setDrawingWordArtText";
    runId: number;
    text: string;
} | {
    kind: "insertBreak";
    runId: number;
    breakKind: "page" | "column";
    nodeIds: number[];
} | {
    kind: "resizeTableColumn";
    cellParagraphId: number;
    boundary: number;
    deltaPx: number;
    renderedWidths?: number[];
} | {
    kind: "resizeTableRow";
    cellParagraphId: number;
    rowIdx: number;
    heightPx: number;
} | {
    kind: "moveTable";
    cellParagraphId: number;
    xPx: number;
    yPx: number;
    preservePageStart: boolean;
    pageDelta: number;
} | {
    kind: "removeDrawing";
    runId: number;
} | {
    kind: "setMathLinear";
    blockId: number;
    mathText: string;
} | {
    kind: "deleteMath";
    blockId: number;
} | {
    kind: "moveMath";
    blockId: number;
    at: {
        blockId: number;
        runId: number;
        offset: number;
    };
    nodeIds: number[];
} | {
    kind: "ensureHeaderFooter";
    hfKind: "header" | "footer";
    nodeIds: number[];
};

type ParagraphAlignment = "left" | "center" | "right" | "justify";
interface PageLayoutPatch {
    /** Margins in inches. */
    margins?: {
        top?: number;
        right?: number;
        bottom?: number;
        left?: number;
    };
    /** Document-global facing-page mode (settings.xml w:mirrorMargins). */
    mirrorMargins?: boolean;
    /** Page size in inches (before orientation). */
    size?: {
        width: number;
        height: number;
    };
    orientation?: "portrait" | "landscape";
    /** Equal-width text columns (1 removes w:cols). */
    columns?: number;
    /** Draw a vertical rule between text columns. Used with columns > 1. */
    columnSeparator?: boolean;
    /** Box page border (null removes). sz in eighth-points, color hex. */
    pageBorders?: {
        sz?: number;
        color?: string;
        offsetFrom?: "text" | "page";
    } | null;
}

/**
 * Table manipulation: add/remove rows and columns relative to the cell
 * containing the caret. Mutates source XML; callers checkpoint + relayout.
 *
 * v1 scope: column operations refuse tables using gridSpan (merged cells)
 * rather than corrupt them.
 */
type TableOp = "rowAbove" | "rowBelow" | "deleteRow" | "colLeft" | "colRight" | "deleteCol" | "deleteTable" | "mergeRight" | "mergeDown" | "splitCell" | {
    kind: "cellShading";
    fill: string | null;
} | {
    kind: "cellVAlign";
    v: "top" | "center" | "bottom";
};

type ParagraphDividerStyle = "single" | "double" | "dotted" | "dashed" | "thinThickSmallGap";
interface ParagraphDivider {
    style: ParagraphDividerStyle;
    color: string;
    widthPt: number;
    spacePt: number;
}

interface CoverPageContent {
    title: string;
    subtitle?: string;
    author?: string;
}
interface LineNumberingPatch {
    /** Turn margin line numbering on/off for the target section(s). */
    enabled: boolean;
    /** Number every Nth line (1/5/10). Default 1 when first enabled. */
    countBy?: number;
    /** When the count resets. Default newPage when first enabled. */
    restart?: "continuous" | "newPage" | "newSection";
    /** First line number. Default 1. */
    start?: number;
}

interface MissingFont {
    /** The face the document asked for (first name in the item's stack). */
    family: string;
    /** A short sample of text that will render in a substitute. */
    sample: string;
}

/**
 * Wire intents for replicated editing.
 *
 * Positions address content by stable id (plan doc 02): `{ blockId, runId,
 * offset }`, where offset is a character offset within the run's text. Ids
 * come from the core StableIds side table; the server validates every
 * client-supplied id before applying (validate-then-install, doc 11) — the
 * types here are the transport shape, not a trust boundary.
 *
 * `base` is the last server sequence number the originating client had
 * applied when it produced the intent; the server transforms the intent
 * against everything sequenced in `(base, seq)` before applying and
 * broadcasting the canonical form (doc 03).
 */
type StableId = number;
/** A caret/endpoint position within a run. */
interface Position {
    blockId: StableId;
    runId: StableId;
    offset: number;
}
interface IntentBase {
    /** Originating client (authenticated identity; the server never trusts a
     * claimed id — see doc 11). */
    clientId: string;
    /** Per-client monotonic counter; `(clientId, clientSeq)` is the idempotency
     * key that dedups re-sends after reconnect (doc 03). */
    clientSeq: number;
    /** Last applied server seq at production time. */
    base: number;
}
/**
 * Tracked-change marks (plan doc 14 §3 L2): strike text ranges as w:del
 * revisions and/or mark paragraph glyphs (split "ins" marks, merge "del"
 * suggestions) — the SUGGESTING counterparts of deleteText/mergeParagraph.
 * Nothing is removed: Word keeps both sides until accept/reject (the
 * already-collab-wired acceptRevision/rejectRevision/acceptAllRevisions).
 * Author/date travel IN the intent (doc 05 rule a: nondeterministic values
 * are generated once by the originator); revision w:ids are scan-derived
 * from identical tree state on every replica. Positional offsets were
 * client-resolved (rule b — no Intl re-derivation at apply).
 */
interface SuggestRevisionIntent extends IntentBase {
    kind: "suggestRevision";
    ranges?: {
        blockId: StableId;
        runId: StableId;
        start: number;
        end: number;
    }[];
    marks?: {
        blockId: StableId;
        glyph: "ins" | "del";
    }[];
    suggest: {
        author: string;
        date: string;
    };
}
/** Insert `text` at a position. With `suggest`, the insertion is recorded as
 * a tracked change (w:ins) carrying the author + date (revision tracking /
 * suggesting mode), rather than a plain insert. */
interface InsertTextIntent extends IntentBase {
    kind: "insertText";
    at: Position;
    text: string;
    /** Tracked-change (suggesting) metadata; omit for a plain insert. */
    suggest?: {
        author: string;
        date: string;
    };
}
/** Delete `[start, end)` characters within a single run. */
interface DeleteTextIntent extends IntentBase {
    kind: "deleteText";
    blockId: StableId;
    runId: StableId;
    start: number;
    end: number;
}
/**
 * Split a paragraph at a position. Text after `at.offset` in run `at.runId`
 * (and following runs) moves to a new paragraph. The originating client
 * allocates the ids for the new paragraph and the new run holding the moved
 * tail; every replica installs these recorded values (doc 03), and they also
 * let the transform remap concurrent positions that fall in the moved tail.
 */
interface SplitParagraphIntent extends IntentBase {
    kind: "splitParagraph";
    at: Position;
    newBlockId: StableId;
    newRunId: StableId;
}
/**
 * Character-format an entire run (bold/italic/underline/…). Whole-run only:
 * it mutates the run's w:rPr in place, so no run splits, no new ids, and the
 * run id is preserved — which makes its transform identity (formatting moves
 * no text). Sub-range formatting splits the run into up to three, needing the
 * run-split id-inheritance + position remapping (plan doc 03 F3); that is a
 * documented extension, not implemented here.
 */
interface FormatRunIntent extends IntentBase {
    kind: "formatRun";
    blockId: StableId;
    runId: StableId;
    /** RunFormatPatch (bold/italic/underline/strike/color/…). Structural shape
     * mirrors @wordinweb/core's RunFormatPatch; carried verbatim. */
    patch: Record<string, unknown>;
}
/**
 * Paragraph-level formatting: alignment and/or paragraph style. Block-level —
 * it mutates the paragraph's w:pPr, creates no tracked nodes, and moves no
 * text, so the block id is preserved and its transform is identity.
 */
interface FormatParagraphIntent extends IntentBase {
    kind: "formatParagraph";
    blockId: StableId;
    align?: "left" | "center" | "right" | "justify";
    /** Paragraph style id; null clears to Normal. Omit to leave unchanged. */
    styleId?: string | null;
}
/**
 * Turn a paragraph into a bullet/numbered list item, or clear its list
 * formatting (kind null). Block-level: mutates w:pPr/numbering, preserves the
 * block id, moves no text — transform identity.
 */
interface SetListTypeIntent extends IntentBase {
    kind: "setListType";
    blockId: StableId;
    listKind: "bullet" | "number" | null;
}
/**
 * Format a character sub-range of a single run (plan doc 03 F3). The run is
 * split into up to three pieces — before [0,start), middle [start,end)
 * (formatted), after [end,len) — replacing the original run. The originating
 * client allocates ids for the pieces that exist and carries them so every
 * replica addresses the pieces identically; the transform remaps any
 * concurrent position in the old run into the correct piece.
 */
interface FormatRangeIntent extends IntentBase {
    kind: "formatRange";
    blockId: StableId;
    runId: StableId;
    start: number;
    end: number;
    patch: Record<string, unknown>;
    /** Piece ids, present iff the piece exists: before when start>0, after when
     * end<runLen. middle always. */
    beforeId?: StableId;
    middleId: StableId;
    afterId?: StableId;
}
/**
 * A table operation that does NOT create new tracked nodes (so no carried ids
 * are needed and the transform is identity): delete row/column/table, cell
 * shading, cell vertical align. The target cell is addressed by the stable id
 * of a paragraph inside it. Row/column INSERTION (which creates cells with new
 * paragraphs and runs needing carried ids) is a documented harder extension.
 */
interface TableOpIntent extends IntentBase {
    /** Stable id of a paragraph inside the target cell. */
    cellParagraphId: StableId;
    kind: "tableOp";
    op: "deleteRow" | "deleteCol" | "deleteTable" | "rowAbove" | "rowBelow" | "colLeft" | "colRight" | {
        kind: "cellShading";
        fill: string | null;
    } | {
        kind: "cellVAlign";
        v: "top" | "center" | "bottom";
    };
    /** For INSERT ops (rowAbove/rowBelow/colLeft/colRight): carried ids for the
     * new tracked nodes (p / r) in document order, so replicas address them
     * alike. */
    nodeIds?: StableId[];
}
/**
 * Merge a paragraph into its predecessor (Backspace at paragraph start). The
 * paragraph's runs MOVE into the previous paragraph (element identity
 * preserved), so run-addressed positions survive unchanged — an identity
 * transform. The merged paragraph's block id is retired.
 */
interface MergeParagraphIntent extends IntentBase {
    kind: "mergeParagraph";
    /** The paragraph to merge into the one before it. */
    blockId: StableId;
}
/**
 * Add a review comment anchored to an entire run. Comment markers
 * (commentRangeStart/End + a commentReference run) are inserted as run
 * siblings and the comment body goes into comments.xml — no commented run's
 * text moves, so the transform is identity. The nondeterministic values
 * (w14:paraId, w:date) are generated once by the originating client and
 * carried here (plan doc 05) so every replica writes identical XML. Sub-range
 * comments (splitting the run) are the documented harder extension.
 */
interface CommentRunIntent extends IntentBase {
    kind: "commentRun";
    runId: StableId;
    text: string;
    author: string;
    initials?: string;
    /** Carried provenance for deterministic XML across replicas. */
    date: string;
    paraId: string;
}
/**
 * Rich paste: splice validated OOXML paragraph blocks after a target
 * paragraph. The pasting client converts clipboard HTML to OOXML locally
 * (browser-only, engine-dependent — doc 02 M4) and carries the serialized
 * blocks; the server VALIDATES them against the positive allowlist before
 * applying (doc 11 gate 2). `nodeIds` are carried ids for the new tracked
 * nodes (p / r) in document order, so every replica addresses them alike.
 * Inserting separate blocks shifts no existing run's offsets — identity
 * transform.
 */
interface PasteBlocksIntent extends IntentBase {
    kind: "pasteBlocks";
    /** Paragraph after which to insert the pasted blocks. */
    afterBlockId: StableId;
    /** Serialized OOXML block list (w:p elements), validated at apply. */
    blocksXml: string;
    /** Carried ids for the new p/r nodes, in document order. */
    nodeIds: StableId[];
}
/**
 * Insert an image at a run (as a sibling drawing run — no text split, identity
 * transform). The plan's media design carries bytes out-of-band via presigned
 * upload with the intent carrying (part, extents, sha) — modeled here with the
 * bytes inline (base64), which is correct for small images and keeps the whole
 * flow testable headlessly; the out-of-band path is a transport optimization
 * over the same convergence. Client-measured dimensions (widthPx/heightPx) are
 * carried so layout reserves space deterministically. `nodeIds` are carried
 * ids for the new drawing run(s).
 */
interface InsertImageIntent extends IntentBase {
    kind: "insertImage";
    runId: StableId;
    /**
     * sha256 hex of the BLOB (plan doc 16 §2): plaintext bytes in plaintext
     * mode, ciphertext in E2EE mode. This is the reservation's hash
     * COMMITMENT — every replica registers the part pending and verifies any
     * later-arriving bytes against it (the swap-proofing trust chain, doc 16
     * §1.1). Bytes NEVER ride the intent (doc 06: sequencer bandwidth is
     * proportional to typing, not content).
     */
    blobSha: string;
    /** Blob length in bytes — validate.ts bounds it (doc 13 blocker 3). */
    bytesLen: number;
    /** File extension without dot (png/jpg/gif…) — drives the content type. */
    ext: string;
    /** E2EE only: base64 12-byte GCM IV, recorded so any holder can re-seal
     * byte-identically for re-supply (doc 16 §5.3). */
    iv?: string;
    widthPx: number;
    heightPx: number;
    nodeIds: StableId[];
}
/**
 * Insert a page/column break at the END of a run (inserts sibling break + tail
 * runs — no text split, identity transform). Mid-run breaks (which split the
 * run) follow the carried-id-split pattern (formatRange) and are the
 * documented extension. `nodeIds` are carried ids for the new runs.
 */
interface InsertBreakIntent extends IntentBase {
    kind: "insertBreak";
    runId: StableId;
    breakKind: "page" | "column";
    nodeIds: StableId[];
}
/**
 * Insert a math (OMML) formula from a linear expression at the end of a run
 * (sibling insertion — identity transform). `nodeIds` for any new tracked
 * nodes.
 */
interface InsertMathIntent extends IntentBase {
    kind: "insertMath";
    runId: StableId;
    mathText: string;
    nodeIds: StableId[];
}
/** Insert a shape/textbox drawing at the end of a run (sibling — identity). */
interface InsertShapeIntent extends IntentBase {
    kind: "insertShape";
    runId: StableId;
    preset: "line" | "verticalLine" | "rectangle" | "roundedRectangle" | "ellipse" | "diamond" | "textBox";
    text?: string;
    nodeIds: StableId[];
}
/**
 * Reply to an existing comment (threading). Addressed by the parent comment's
 * id (a deterministic string). Carries provenance (w14:paraId candidates in
 * consumption order + the w:date) so every replica writes identical XML. No
 * document run's text moves — identity transform.
 */
interface ReplyCommentIntent extends IntentBase {
    kind: "replyComment";
    parentId: string;
    text: string;
    author: string;
    initials?: string;
    date: string;
    /** paraId candidates in the order replyToComment consumes them: the reply's
     * paraId, preceded by the parent's if the parent lacks one. */
    paraIds: string[];
}
/** Adjust a paragraph's indent (Tab/Shift-Tab). Block-level, identity. */
interface AdjustIndentIntent extends IntentBase {
    kind: "adjustIndent";
    blockId: StableId;
    direction: 1 | -1;
}
/** Set paragraph line/before/after spacing. Block-level, identity. */
interface SetSpacingIntent extends IntentBase {
    kind: "setSpacing";
    blockId: StableId;
    /** ParagraphSpacingPatch (before/after/line/lineRule) — carried verbatim. */
    patch: Record<string, unknown>;
}
/** Insert a page-number field at the end of a run (sibling insertion,
 * identity). Page fields are deterministic (no clock). */
interface InsertPageFieldIntent extends IntentBase {
    kind: "insertPageField";
    runId: StableId;
    fieldKind: "page" | "pageOfTotal";
    nodeIds: StableId[];
}
/** Wrap a run in a hyperlink to `url`. The url is scheme-validated at apply
 * (rejects javascript:/data: etc. — doc 11 gate 1). Run-level, identity. */
interface SetLinkIntent extends IntentBase {
    kind: "setLink";
    runId: StableId;
    url: string;
    nodeIds: StableId[];
}
/** Insert a footnote reference at the end of a run + its footnote text. */
interface InsertFootnoteIntent extends IntentBase {
    kind: "insertFootnote";
    runId: StableId;
    text: string;
    nodeIds: StableId[];
}
/** Set/clear a drop cap on a paragraph. Block-level. */
interface SetDropCapIntent extends IntentBase {
    kind: "setDropCap";
    blockId: StableId;
    mode: "drop" | "margin" | null;
    nodeIds: StableId[];
}
/** Set/clear a paragraph divider (horizontal rule). Block-level, identity. */
interface SetDividerIntent extends IntentBase {
    kind: "setDivider";
    blockId: StableId;
    divider: Record<string, unknown> | null;
}
/** Insert a bookmark anchor at the end of a run. */
interface InsertBookmarkIntent extends IntentBase {
    kind: "insertBookmark";
    runId: StableId;
    name: string;
}
type Intent = InsertTextIntent | DeleteTextIntent | SplitParagraphIntent | FormatRunIntent | FormatParagraphIntent | SetListTypeIntent | FormatRangeIntent | TableOpIntent | MergeParagraphIntent | CommentRunIntent | PasteBlocksIntent | InsertImageIntent | SuggestRevisionIntent | InsertBreakIntent | InsertMathIntent | InsertShapeIntent | ReplyCommentIntent | AdjustIndentIntent | SetSpacingIntent | InsertPageFieldIntent | SetLinkIntent | InsertFootnoteIntent | SetDropCapIntent | SetDividerIntent | InsertBookmarkIntent | InsertBlankPageIntent | InsertSectionBreakIntent | InsertCrossRefIntent | InsertCoverPageIntent | SetPageLayoutIntent | SetListLevelIntent | InsertWordArtIntent | InsertChartIntent | InsertSmartArtIntent | SetLineNumberingIntent | InsertDateTimeFieldIntent | InsertFieldIntent | SetDrawingRotationIntent | SetDrawingFillIntent | SetChartDataIntent | SetSmartArtNodeTextIntent | SetDrawingWordArtTextIntent | SetDrawingLineStyleIntent | SetImageAltTextIntent | RemoveLinkIntent | SetImageWrapIntent | SetDrawingOrderIntent | SetSmartArtDataIntent | SetSmartArtFillIntent | SetSmartArtTextFormatIntent | SetFloatingPagePositionIntent | ResizeDrawingIntent | ResizeTableColumnIntent | ResizeTableRowIntent | MoveTableIntent | RemoveDrawingIntent | SetMathLinearIntent | DeleteMathIntent | MoveMathIntent | EnsureHeaderFooterIntent | DeleteCommentIntent | InsertBookmarkRangeIntent | ToggleCheckboxIntent | AcceptRevisionIntent | RejectRevisionIntent | AcceptAllRevisionsIntent | InsertTableIntent;
/** Insert a blank page at the end of a run. */
interface InsertBlankPageIntent extends IntentBase {
    kind: "insertBlankPage";
    runId: StableId;
    nodeIds: StableId[];
}
/** Insert a section break at the end of a run. */
interface InsertSectionBreakIntent extends IntentBase {
    kind: "insertSectionBreak";
    runId: StableId;
    breakType: "nextPage" | "continuous";
    nodeIds: StableId[];
}
/** Insert a cross-reference to a bookmark at the end of a run. */
interface InsertCrossRefIntent extends IntentBase {
    kind: "insertCrossRef";
    runId: StableId;
    bookmark: string;
    refKind: "text" | "page";
    nodeIds: StableId[];
}
/** Insert a cover page (document-level). */
interface InsertCoverPageIntent extends IntentBase {
    kind: "insertCoverPage";
    /** CoverPageContent ({title, subtitle?, author?, …}) — carried verbatim. */
    content: Record<string, unknown>;
    nodeIds: StableId[];
}
/** Update page setup: margins/size/orientation/columns/borders. When target is
 * omitted it applies to every section; document-level, identity transform. */
interface SetPageLayoutIntent extends IntentBase {
    kind: "setPageLayout";
    /** PageLayoutPatch ({margins?, size?, orientation?, columns?, …}) — carried
     * verbatim; validated against a shape+range allowlist. */
    patch: Record<string, unknown>;
}
/** Indent (+1) or outdent (-1) a paragraph's list nesting level. Block-level. */
interface SetListLevelIntent extends IntentBase {
    kind: "setListLevel";
    blockId: StableId;
    delta: 1 | -1;
}
/** Insert decorative WordArt (a text drawing) at the end of a run. */
interface InsertWordArtIntent extends IntentBase {
    kind: "insertWordArt";
    runId: StableId;
    text: string;
    preset: "plain" | "archUp" | "archDown" | "wave" | "chevron";
    nodeIds: StableId[];
}
/** Insert a data chart (column/bar/line/pie) at the end of a run. The chart
 * data is carried verbatim; the workbook part is generated deterministically. */
interface InsertChartIntent extends IntentBase {
    kind: "insertChart";
    runId: StableId;
    chart: {
        type: "column" | "bar" | "line" | "pie";
        title?: string;
        categories: string[];
        series: {
            name: string;
            values: number[];
        }[];
    };
    nodeIds: StableId[];
}
/** Insert a SmartArt diagram (process/cycle/hierarchy/list) at end of a run. */
interface InsertSmartArtIntent extends IntentBase {
    kind: "insertSmartArt";
    runId: StableId;
    smartArt: {
        layout: "process" | "cycle" | "hierarchy" | "list";
        items: string[];
    };
    nodeIds: StableId[];
}
/** Toggle/configure margin line numbering for the section(s). Document-level. */
interface SetLineNumberingIntent extends IntentBase {
    kind: "setLineNumbering";
    patch: {
        enabled: boolean;
        countBy?: number;
        restart?: "continuous" | "newPage" | "newSection";
        start?: number;
    };
}
/** Insert an auto-updating DATE or TIME field at the end of a run. The picture
 * (format string) is carried so every replica renders identical field XML. */
interface InsertDateTimeFieldIntent extends IntentBase {
    kind: "insertDateTimeField";
    runId: StableId;
    dtKind: "date" | "time";
    picture: string;
    nodeIds: StableId[];
}
/** Insert a Word field (e.g. PAGE, NUMPAGES, REF) at the end of a run. The
 * instruction is restricted to a safe positive allowlist (no INCLUDETEXT/DDE/
 * LINK external-content fields). */
interface InsertFieldIntent extends IntentBase {
    kind: "insertField";
    runId: StableId;
    instruction: string;
    cachedResult?: string;
    nodeIds: StableId[];
}
/** Rotate the drawing carried by a run. Identity transform (run-addressed). */
interface SetDrawingRotationIntent extends IntentBase {
    kind: "setDrawingRotation";
    runId: StableId;
    degrees: number;
}
/** Set/clear the solid fill of the drawing carried by a run. `color` is a
 * 6-hex-digit RGB (no #) or null to clear. */
interface SetDrawingFillIntent extends IntentBase {
    kind: "setDrawingFill";
    runId: StableId;
    color: string | null;
}
/** Replace the data of the chart carried by a run (same shape as insertChart's
 * chart). Edits an existing chart in place. */
interface SetChartDataIntent extends IntentBase {
    kind: "setChartData";
    runId: StableId;
    chart: {
        type: "column" | "bar" | "line" | "pie";
        title?: string;
        categories: string[];
        series: {
            name: string;
            values: number[];
        }[];
    };
}
/** Set the text of one node of the SmartArt diagram carried by a run. */
interface SetSmartArtNodeTextIntent extends IntentBase {
    kind: "setSmartArtNodeText";
    runId: StableId;
    index: number;
    text: string;
}
/** Replace the text of the WordArt drawing carried by a run. */
interface SetDrawingWordArtTextIntent extends IntentBase {
    kind: "setDrawingWordArtText";
    runId: StableId;
    text: string;
}
/** Set the outline (line) style of the drawing carried by a run. */
interface SetDrawingLineStyleIntent extends IntentBase {
    kind: "setDrawingLineStyle";
    runId: StableId;
    color: string;
    widthPx: number;
    dash: "solid" | "dashed" | "dotted";
}
/** Set/clear the alt text (accessibility description) of the image/drawing
 * carried by a run. Empty string clears it. */
interface SetImageAltTextIntent extends IntentBase {
    kind: "setImageAltText";
    runId: StableId;
    alt: string;
}
/** Remove the hyperlink wrapping a run's text (inverse of setLink). */
interface RemoveLinkIntent extends IntentBase {
    kind: "removeLink";
    runId: StableId;
}
/** Set the text-wrap mode of the drawing carried by a run. A floating mode
 * converts an inline drawing to an anchor. */
interface SetImageWrapIntent extends IntentBase {
    kind: "setImageWrap";
    runId: StableId;
    mode: "inline" | "square" | "topAndBottom" | "none" | "behind";
}
/** Bring the (floating) drawing carried by a run to front/back. */
interface SetDrawingOrderIntent extends IntentBase {
    kind: "setDrawingOrder";
    runId: StableId;
    order: "front" | "back";
}
/** Replace the whole SmartArt diagram (layout + items) carried by a run. */
interface SetSmartArtDataIntent extends IntentBase {
    kind: "setSmartArtData";
    runId: StableId;
    smartArt: {
        layout: "process" | "cycle" | "hierarchy" | "list";
        items: string[];
    };
}
/** Set/clear the fill of a SmartArt node (or all nodes when nodeIndex omitted).
 * `color` is 6-hex RGB (no #) or null to clear. */
interface SetSmartArtFillIntent extends IntentBase {
    kind: "setSmartArtFill";
    runId: StableId;
    color: string | null;
    nodeIndex?: number;
}
/** Set the text format of a SmartArt node (or all nodes). */
interface SetSmartArtTextFormatIntent extends IntentBase {
    kind: "setSmartArtTextFormat";
    runId: StableId;
    format: {
        fontFamily: string;
        fontSizePt: number;
        color: string;
        bold: boolean;
        italic: boolean;
        alignment: "left" | "center" | "right";
    };
    nodeIndex?: number;
}
/** Position the (floating) drawing carried by a run relative to the page. */
interface SetFloatingPagePositionIntent extends IntentBase {
    kind: "setFloatingPagePosition";
    runId: StableId;
    xPx: number;
    yPx: number;
}
/** Resize the drawing carried by a run (extent in px; the mutation clamps to
 * a 1px-EMU floor and keeps line geometry degenerate on its minor axis). */
interface ResizeDrawingIntent extends IntentBase {
    kind: "resizeDrawing";
    runId: StableId;
    widthPx: number;
    heightPx: number;
}
/** Drag-resize a table column boundary. Addressed like tableOp: a paragraph
 * inside the table. `renderedWidths` (the dragger's measured column widths)
 * rides as DATA so every replica applies the identical mutation. */
interface ResizeTableColumnIntent extends IntentBase {
    kind: "resizeTableColumn";
    cellParagraphId: StableId;
    boundary: number;
    deltaPx: number;
    renderedWidths?: number[];
}
/** Drag-resize a table row's height (trHeight atLeast). */
interface ResizeTableRowIntent extends IntentBase {
    kind: "resizeTableRow";
    cellParagraphId: StableId;
    rowIdx: number;
    heightPx: number;
}
/** Position a table at page coordinates (converts inline to float). */
interface MoveTableIntent extends IntentBase {
    kind: "moveTable";
    cellParagraphId: StableId;
    xPx: number;
    yPx: number;
    preservePageStart: boolean;
    pageDelta: number;
}
/** Delete the drawing carried by a run (removes the carrier run). */
interface RemoveDrawingIntent extends IntentBase {
    kind: "removeDrawing";
    runId: StableId;
}
/** Replace the equation of the math object in a paragraph (linear syntax).
 * Math (m:oMath) lives at paragraph level, so this is block-addressed. */
interface SetMathLinearIntent extends IntentBase {
    kind: "setMathLinear";
    blockId: StableId;
    mathText: string;
}
/** Delete the math object in a paragraph. */
interface DeleteMathIntent extends IntentBase {
    kind: "deleteMath";
    blockId: StableId;
}
/** Drag-move the equation out of `blockId` to the text position `at`. The
 * equation is addressed like the other math intents (block + firstMathIn); the
 * destination is an ordinary wire caret. Dropping mid-text splits the
 * destination run, so the split-off tail run takes a carried id. */
interface MoveMathIntent extends IntentBase {
    kind: "moveMath";
    blockId: StableId;
    at: Position;
    nodeIds: StableId[];
}
/** Create the document's header (or footer) part the way Word does on first
 * entry into the margin band: a new part + relationship + content-type
 * override, referenced from every sectPr (one is materialized when a minimal
 * document has none). Applies as a clean no-op when the part already exists,
 * so it is safe for two participants to open the header at the same moment.
 * The created paragraph and run take carried ids — every later edit in the
 * header addresses them. */
interface EnsureHeaderFooterIntent extends IntentBase {
    kind: "ensureHeaderFooter";
    hfKind: "header" | "footer";
    nodeIds: StableId[];
}
/** Delete a comment (and its reply thread) by id. */
interface DeleteCommentIntent extends IntentBase {
    kind: "deleteComment";
    commentId: string;
}
/** Wrap a sub-range of a run's text in a named bookmark. */
interface InsertBookmarkRangeIntent extends IntentBase {
    kind: "insertBookmarkRange";
    runId: StableId;
    name: string;
    start: number;
    end: number;
}
/** Toggle the checkbox content control carried by a run. */
interface ToggleCheckboxIntent extends IntentBase {
    kind: "toggleCheckbox";
    runId: StableId;
}
/** Accept one tracked change, addressed by its position in document order
 * (collectRevisions order — deterministic across replicas). */
interface AcceptRevisionIntent extends IntentBase {
    kind: "acceptRevision";
    index: number;
}
/** Reject one tracked change, addressed by document-order index. */
interface RejectRevisionIntent extends IntentBase {
    kind: "rejectRevision";
    index: number;
}
/** Accept every tracked change in the document. */
interface AcceptAllRevisionsIntent extends IntentBase {
    kind: "acceptAllRevisions";
}
/** Insert a rows×cols table after the paragraph containing the anchor run. */
interface InsertTableIntent extends IntentBase {
    kind: "insertTable";
    runId: StableId;
    rows: number;
    cols: number;
    nodeIds: StableId[];
}
/** A sequenced log entry: an applied intent with its assigned seq, or a
 * rejection no-op (doc 03) that still occupies a position in the total order
 * so every replica agrees where a drop took effect. */
type LogEntry = {
    seq: number;
    kind: "applied";
    intent: Intent;
} | {
    seq: number;
    kind: "rejected";
    clientId: string;
    clientSeq: number;
    reason: string;
};

/** The ID sidecar carried in a checkpoint bundle (plan doc 03). */
type IdSidecar = ReturnType<StableIds["exportSidecar"]>;
/**
 * What the top of a client's undo stack offers. The three cases are
 * deliberately distinct because the UI says something different for each:
 * nothing to undo (button disabled), the last action can't be reversed yet
 * (button disabled WITH a reason), or here is the inverse to submit.
 */
/**
 * What a collaborative undo attempt did, as the UI needs to distinguish it:
 *
 *  undone         the inverse was submitted and painted optimistically
 *  cannot-undo    the last action has no inverse yet — a HARD STOP; undo does
 *                 not reach past it (the cure is more inverses, not skipping)
 *  changed-since  the target is already gone in canonical state, so the undo
 *                 could never land. The entry IS consumed: what the user
 *                 asked for has effectively happened, so their next press
 *                 moving to the previous action cannot surprise them
 *  nothing-to-undo  the stack is empty
 *  unavailable    this connection has no collaborative undo (plaintext rooms,
 *                 whose authority is the server)
 */
type UndoOutcome = "undone" | "cannot-undo" | "changed-since" | "nothing-to-undo" | "unavailable";

/**
 * E2EE primitives for the blind-sequencer mode (plan doc 13). Pure WebCrypto
 * (available in browsers and Node ≥16 globals) — no dependencies, usable
 * from both the client and tests.
 *
 * Key model (doc 13 §1): a per-DOCUMENT master key `K_doc` minted at
 * go-live, carried only in the share link's URL fragment. Per-epoch subkeys
 * are DERIVED — `K_epoch = HKDF(K_doc, genesisId [+ share code])` — then
 * domain-separated into K_content (intents/checkpoints) and K_media
 * (blobs). Epoch derivation kills cross-epoch replay (round-4 F5a) and
 * resets the GCM IV budget per epoch (F21); nobody ever redistributes keys,
 * because anyone with the link + the (public) epoch id derives the same
 * values.
 *
 * AAD discipline (round-4 F5): every ciphertext is bound to its exact
 * position in the protocol — intents to (docId, genesisId, clientId,
 * clientSeq, base); checkpoints to (docId, genesisId, seq). `base` is
 * authenticated because it is a TRANSFORM INPUT: an unauthenticated base
 * would let a keyless server shift where an edit applies.
 */
/** The plaintext bookkeeping + opaque body the blind sequencer handles. */
interface IntentEnvelope {
    clientId: string;
    clientSeq: number;
    base: number;
    /** base64 12-byte GCM IV (fresh per envelope). */
    iv: string;
    /** base64 AES-256-GCM ciphertext (+tag) of the JSON-serialized intent. */
    ciphertext: string;
}
interface EpochKeys {
    kContent: CryptoKey;
    kMedia: CryptoKey;
    /**
     * Presence (carets/selections). A THIRD domain-separated key rather than
     * kContent, because presence is high-frequency, low-value material with a
     * completely different exposure profile: it is sealed and opened on every
     * caret move, so it burns IVs fastest and is the likeliest place for a
     * future mistake. Key separation costs one HKDF call at join and means a
     * presence-side compromise cannot touch document content.
     */
    kPresence: CryptoKey;
}
declare function bytesToB64(bytes: Uint8Array): string;
/** Mint a fresh 256-bit document master key (go-live, doc 13 §1). Returned
 * base64url — the exact string that rides the URL fragment (`#k=`). */
declare function mintDocKey(): string;
/** Parse `#k=<key>` out of a URL fragment; null if absent. The PRESENCE of
 * this is what fixes the doc's mode client-side (doc 13 §6: `#k` present ⇒
 * encrypted, always — a welcome claiming otherwise is hard-refused). */
declare function docKeyFromFragment(fragment: string): string | null;
/**
 * Derive this epoch's working keys. Share code (doc 13 §7) mixes in when
 * set: a leaked link without the code cannot decrypt anything. The code is
 * stretched by the CALLER via stretchShareCode (PBKDF2) — this function
 * just mixes; keeping stretching separate lets the UI do it once.
 */
declare function deriveEpochKeys(docKeyB64url: string, genesisId: string, stretchedCode?: Uint8Array): Promise<EpochKeys>;
/** PBKDF2-SHA256 stretch of the share code (doc 13 §7): 600k iterations,
 * WebCrypto-native (no WASM dep), salt bound to the docId so a precomputed
 * table for one doc is useless for another. Deliberately NOT epoch-salted:
 * the hello must carry the proof BEFORE the welcome reveals the epoch id
 * (chicken-and-egg otherwise); epoch-binding of the KEYS happens in
 * deriveEpochKeys' HKDF salt instead, which is where it matters. */
declare function stretchShareCode(code: string, docId: string): Promise<Uint8Array>;
/** Seal a checkpoint bundle (doc 13 §3): AAD binds (docId, genesisId, seq)
 * so a stored checkpoint cannot be replayed at another position (F5c). The
 * body carries the docx + sidecar + the canonical docHash for joiner
 * cross-checking (blocker-2 verification). */
/**
 * The sealed checkpoint's plaintext body (doc 13 §3). `mediaMeta` is the
 * doc-16 §6 late-join map: the declared sha (+ IV) of every registered media
 * part, so a joiner whose snapshot contains the REGISTRATIONS can also fetch
 * the bytes — the addresses live only in already-folded intents otherwise,
 * which is what made late-join media unfetchable.
 *
 * It rides INSIDE the sealed body on purpose: in an encrypted room the server
 * must not learn the document's part structure. It learns which shas exist
 * only when someone PUTs or GETs one, and nothing more.
 *
 * OPTIONAL and tolerantly read — a checkpoint sealed by an older build simply
 * has no map, and such a joiner degrades to reserving the box without being
 * able to fill it (the prior behavior). Additive, so no engine bump.
 */
interface CheckpointBody {
    docx: string;
    sidecar: unknown;
    docHash: string;
    mediaMeta?: {
        part: string;
        sha: string;
        iv?: string;
    }[];
}
declare function sealCheckpoint(kContent: CryptoKey, docId: string, genesisId: string, seq: number, body: CheckpointBody): Promise<{
    iv: string;
    ciphertext: string;
}>;
/** A presence payload as it crosses a BLIND relay: an opaque blob. The
 * server sees two base64 strings and the sender's id — never a coordinate. */
interface SealedPresence {
    iv: string;
    ciphertext: string;
}

/** A sequenced envelope in an encrypted room's log (doc 13 §2): the
 * server-assigned seq plus the opaque envelope. The server can read only
 * the bookkeeping; clients derive the canonical form themselves. */
type EnvelopeEntry = IntentEnvelope & {
    seq: number;
};
/**
 * One head in a document's LINEAGE chain (doc 15 §1): the identity of the
 * confirmed state a session epoch ended at. Ancestry is decided by hash
 * membership — never by dates (clocks are self-asserted). A rejoiner whose
 * own head appears in the seed's lineage is a strict ancestor and can
 * FAST-FORWARD silently; anything else is true divergence (draft/fork).
 */
interface LineageHead {
    genesisId: string;
    seq: number;
    /** sha256 hex of the confirmed docx bytes at (genesisId, seq). */
    docHash: string;
}
/** A sealed checkpoint blob as the wire carries it (doc 13 §3). */
interface SealedCheckpoint {
    seq: number;
    iv: string;
    ciphertext: string;
}
/**
 * Wire protocol between a collab client and the server host. Transport-
 * agnostic: the hub speaks these messages; a WebSocket (or in-memory test)
 * adapter serializes them. Kept deliberately small (plan doc 06).
 */
/** One highlighted stretch of a remote participant's selection: a half-open
 * `[start, end)` inside ONE run, in the same WIRE offset basis carets and
 * suggestRevision use (cumulative within the run, inline separators counting
 * one unit each — see core `EncodedCaret`). The sender emits one entry per
 * selection segment, and a segment never straddles a w:t, so a range stays
 * inside a single run by construction. */
interface PresenceRange {
    blockId: number;
    runId: number;
    start: number;
    end: number;
}
/** An ephemeral cursor/selection position on the presence channel: stable-id
 * addresses (plan doc 03). Never logged, never persisted. */
interface PresencePosition {
    anchor: {
        blockId: number;
        runId: number;
        offset: number;
    };
    focus?: {
        blockId: number;
        runId: number;
        offset: number;
    };
    /** Selection highlight (Google-Docs style): the stretches the participant
     * has selected, drawn in their presence color under the remote caret.
     * OPTIONAL and additive — a payload from an older client carries only the
     * anchor and still renders a caret. Absent/empty ⇒ no selection. */
    ranges?: PresenceRange[];
}
/**
 * A participant's self-asserted display profile (plan doc 14 §2). No
 * accounts: anyone with the link can claim any name — stated in the UI
 * ("names are chosen by participants") — but the F4 clientId binding
 * guarantees CONTINUITY: one name maps to one actor per session and nobody
 * can impersonate an existing participant mid-session. The server sanitizes
 * (length, control chars, color shape) and clients still render text-node-
 * only with palette-validated colors (doc 11 XSS vector 7 — defense in
 * depth, since a hostile server or E2EE mode moves sanitization client-side).
 */
interface ParticipantProfile {
    /** 1–40 chars after server sanitization. */
    name: string;
    /** Hex color `#rrggbb`; invalid values are replaced server-side by a
     * palette color derived from the clientId hash. */
    color: string;
}
/** One roster row (doc 14 §2): keyed by the BOUND clientId — the same
 * identity intents carry, so attribution, presence, and roster share one
 * keyspace. Disconnected participants stay listed (greyed) for the session's
 * lifetime; a reconnect under the same clientId resumes the entry. */
/**
 * Whether a participant's edits will be SEQUENCED, and if not, why.
 *
 * A POSITIVE signal, which is the entire point. Until this existed a client
 * could only discover it was blocked by making an edit and being refused —
 * which cost the user their first keystrokes every time (optimistic local
 * mutation, refusal, heal: a burst of text that appears and vanishes), and
 * left them stuck in viewer mode after a lock was LIFTED, because nothing
 * announced the good news either.
 *
 * The three blocked states arrive as one `read-only` refusal on the wire, so
 * they are separated here — the difference is what a UI can honestly say:
 *  - `owner-lock`  the owner paused editing for the whole room; it may lift
 *  - `demoted`     the owner made THIS participant a viewer
 *  - `viewer-role` the access token grants read access only; nothing the
 *                  owner does in-session changes it
 */
type WriteStatus = "allowed" | "owner-lock" | "demoted" | "viewer-role";
/** One roster row (doc 14 §2): keyed by the BOUND clientId — the same
 * identity intents carry, so attribution, presence, and roster share one
 * keyspace. Disconnected participants stay listed (greyed) for the session's
 * lifetime; a reconnect under the same clientId resumes the entry. */
interface RosterEntry {
    clientId: string;
    profile: ParticipantProfile;
    connected: boolean;
    /**
     * Whether this participant may write, re-fanned on every transition.
     *
     * OPTIONAL because it is additive: an older server omits it. A client that
     * sees `undefined` must NOT assume "allowed" — it should fall back to the
     * refusal-driven behaviour it had before, exactly as the published media
     * limit's `null` means "skip the pre-check". Inventing a permissive default
     * would put the user back in the first-edit-lost loop this removes.
     *
     * PER-CLIENT, not a room flag: an owner keeps writing through their own
     * room-wide lock, so a single broadcast value would be wrong for them.
     */
    write?: WriteStatus;
}
/** Client → server. */
type ClientMessage = {
    t: "hello";
    protocolVersion: number;
    docId: string;
    /**
     * The identity every subsequent submit/presence message is bound to.
     * The hub registers it for this socket at hello and REFUSES any submit
     * whose intent.clientId differs (plan doc 11 decision 8 / round-4 F4):
     * `(clientId, clientSeq)` is the idempotency key, so an unbound claimed
     * id would let one client poison another's dedup space (silent edit
     * loss) and forge attribution. Registration-on-first-hello is
     * sufficient — the id needs continuity, not global meaning.
     */
    clientId: string;
    /**
     * Single-live-connection rule (plan doc 12 §7): a second hello for the
     * same (docId, clientId) is refused `already-open` — same-profile tabs
     * share localStorage identity and would collide clientSeq counters.
     * With `takeover: true` the NEW connection wins instead: the hub kicks
     * the old socket (it may be a zombie tab) and admits this one.
     */
    takeover?: boolean;
    token?: string;
    sinceSeq: number;
    /** The epoch id from the client's stored bundle, when resuming (plan
     * doc 12 §5). Informational to the server today — epoch comparison is
     * client-side (welcome.genesisId vs the bundle's) — but carried so the
     * server can log resume-vs-fresh joins and future-proof case handling. */
    genesisId?: string;
    /** Display profile (doc 14 §2) — optional; omitted = anonymous default
     * derived server-side. In E2EE mode this becomes an opaque encrypted
     * blob (doc 13/14) — the shape change rides that protocol bump. */
    profile?: ParticipantProfile;
    /** Engine (transform/apply semantics) version — REQUIRED for encrypted
     * rooms, where the fence prevents mixed-semantics divergence. */
    engineVersion?: string;
    /**
     * Share-code proof (doc 13 §7): the PBKDF2-stretched code, base64.
     * The server compares against the verifier registered at seed time and
     * enforces an attempt budget — 10^6 combinations is plenty against 5
     * online tries, and in E2EE mode the code ALSO mixes into key
     * derivation, so even a server bypass yields undecryptable content.
     */
    codeProof?: string;
    /**
     * Owner-capability proof (doc 14 §2.5): the token minted at seed and
     * returned ONLY to the seeder — never part of the shared link. A
     * matching token flags this connection as the epoch's owner
     * (read-only bypass + admin channel). Roles are capabilities, not
     * accounts: lose the bundle, lose the crown; re-seed to reclaim.
     */
    ownerToken?: string;
}
/** Update this connection's profile mid-session (rename / recolor). */
 | {
    t: "profile";
    profile: ParticipantProfile;
}
/**
 * Owner admin channel (doc 14 §2.5) — only honored from a connection
 * whose hello proved the owner token. All enforceable by a BLIND server
 * (identity/role are plaintext bookkeeping). Honest limits: kick/demote
 * key by clientId (a determined user can mint a fresh id — a nudge, not
 * a wall, without accounts), and in E2EE mode an owner controls WRITES,
 * never reads (re-key is the only read fence).
 */
 | {
    t: "admin";
    action: {
        op: "kick";
        clientId: string;
    } | {
        op: "readOnly";
        on: boolean;
    } | {
        op: "setRole";
        clientId: string;
        role: "editor" | "viewer";
    };
} | {
    t: "submit";
    intent: Intent;
}
/** Encrypted-mode submit (doc 13 §2): opaque body, plaintext bookkeeping.
 * The hub refuses this on plaintext rooms and refuses plaintext `submit`
 * on encrypted rooms — no mixed-mode documents, enforced both ways. */
 | {
    t: "submit-enc";
    envelope: IntentEnvelope;
}
/** Upload a sealed checkpoint (doc 13 §3) — accepted ONLY from the
 * connection the server currently designates as checkpointer. */
 | {
    t: "checkpoint";
    checkpoint: SealedCheckpoint;
}
/** Media re-supply (doc 16 §3): "I need blob <sha>" (relay miss). */
 | {
    t: "media-need";
    sha: string;
}
/** "I hold these blobs" — reply to media-request, or unsolicited right
 * after a welcome whose mediaNeeded intersects local holdings (§5.4). */
 | {
    t: "media-have";
    shas: string[];
}
/** Encrypted hash gossip (doc 13 §2): an OPAQUE sealed {seq, hash} blob
 * the server relays without reading — divergence detection must not leak
 * even document hashes to a blind server (a hash is a stable content
 * fingerprint: confirmable by anyone holding a guess). */
 | {
    t: "gossip";
    iv: string;
    ciphertext: string;
}
/** Presence. In an ENCRYPTED room this is a `SealedPresence` blob and the
 * hub relays it without looking (#20): the server learns who is pointing,
 * never where. Plaintext rooms keep the structured position, which the hub
 * still clamps at the relay. */
 | {
    t: "presence";
    position: PresencePosition | SealedPresence | null;
}
/**
 * APPLICATION-LEVEL LIVENESS PROBE (the reconnect arc). The server answers
 * every one with {@link ServerMessage} `pong` carrying the same nonce.
 *
 * WHY THIS EXISTS AT ALL, given TCP and the WebSocket close handshake: a
 * phone that sleeps mid-session leaves a HALF-OPEN connection. The peer is
 * gone, but nothing tells this end — `onclose` never fires, `onerror` never
 * fires, `readyState` still reads OPEN, and every `send()` succeeds into a
 * socket whose bytes will never arrive. The only way to learn the truth is
 * to ask for a reply and notice it did not come.
 *
 * `nonce` is an opaque round-trip token, echoed back unmodified. It exists
 * so a pong can be matched to the ping it answers — without it a stale pong
 * from before a reconnect could silently satisfy the CURRENT probe and mask
 * a genuinely dead socket. It carries no meaning and is never persisted.
 *
 * DELIBERATELY NOT ACTIVITY. The hub answers this WITHOUT calling
 * `noteActivity`, so a heartbeat cannot hold an abandoned room open — the
 * exact hazard the idle-timeout notes call out ("an activity-based clock is
 * what a keepalive script defeats"). A ping proves the socket is alive; it
 * proves nothing about a human being present, and those are different
 * questions with different consequences.
 */
 | {
    t: "ping";
    nonce: number;
};
/** Server → client. */
type ServerMessage = {
    t: "welcome";
    docId: string;
    seq: number;
    snapshot: string;
    /**
     * The stable-id sidecar for the snapshot (plan doc 12 §2, round-4 F10).
     * A snapshot NEVER travels without it: the joiner cannot re-derive the
     * id table from parse order once history contains split-created carried
     * ids (round-2 F1) — it would silently mis-address every later intent.
     */
    sidecar: IdSidecar;
    tail: LogEntry[];
    /**
     * The session EPOCH id (plan doc 12): minted fresh every time a doc is
     * seeded/re-seeded. A resuming client compares it against its bundle's
     * stored genesisId — same ⇒ seamless rejoin (case 1); different ⇒
     * someone re-seeded while it was away (case 2: take server state, keep
     * the local copy per doc 15 lineage — NEVER silently merge epochs).
     */
    genesisId: string;
    /** Session encryption mode (plan doc 13 §6). Clients derive the truth
     * from their link (#k present ⇒ encrypted) and HARD-REFUSE a welcome
     * that contradicts it — the wire value must never downgrade a client. */
    mode: "plaintext" | "encrypted";
    /** The seed's lineage chain (doc 15): lets a rejoining holder decide
     * fast-forward vs fork CLIENT-side. Absent for provider-created and
     * legacy rooms (⇒ every epoch mismatch is a fork, the safe default). */
    lineage?: LineageHead[];
    /** Shas with waiters / outstanding unavailability (doc 16 §3): a
     * joining holder intersects with its local media and volunteers —
     * the mechanism behind "reappears when a holder rejoins". */
    mediaNeeded?: string[];
    /** Per-part media ADDRESSES (doc 16 §6 late-join): the declared sha
     * (+ E2EE iv) for every registered media part, so a joiner can fetch
     * pixels its snapshot references but does not contain — without this
     * the address lives only in the already-folded insertImage intent and
     * late-join media is unfetchable. Plaintext rooms only; encrypted
     * rooms carry the same map INSIDE the sealed checkpoint body (the
     * server must not learn part structure beyond what PUT addresses
     * already reveal). Parse-derived holes with unknown sha are omitted. */
    media?: {
        part: string;
        sha: string;
        iv?: string;
    }[];
    /**
     * The relay's configured per-blob size limit, in bytes, so a client can
     * CHECK A FILE BEFORE UPLOADING IT rather than discovering the limit
     * from a 413 afterwards.
     *
     * The number flows FORWARD — server to client, once per session — which
     * is why it lives here rather than being threaded back through the
     * upload's return type. A refusal has to travel up through every layer
     * between the relay and the file picker, and each of those layers has
     * its own idea of what failure looks like; a value published at join
     * time has no such path to get lost on.
     *
     * NOT A SECRET, and nothing is gated on it: the same number already
     * appears in a 413 body, and anyone can discover it by uploading
     * something large. It is per-SERVER-CONFIG, not per-room, so it carries
     * nothing about the room's contents or participants.
     *
     * Optional and additive: an older server omits it, and a client that
     * gets nothing should skip the pre-check rather than invent a limit —
     * the server still enforces the real one.
     */
    mediaMaxBlobBytes?: number;
} | {
    t: "broadcast";
    entries: LogEntry[];
}
/**
 * Encrypted-mode welcome (doc 13 §3): the seed checkpoint (sealed by the
 * epoch's seeder) + the epoch's envelope log after it. BASE-COMPLETE by
 * construction (round-4 blocker 1): the demo server retains the WHOLE
 * epoch log in RAM (doc 12 — rooms are session-scoped), so every entry a
 * joiner needs to re-derive canonical forms is present; client-produced
 * mid-session checkpoints (doc 13 item 6) are a RAM optimization on top,
 * not a correctness requirement, and land with the retention rule.
 */
 | {
    t: "welcome-enc";
    docId: string;
    genesisId: string;
    checkpoint: SealedCheckpoint;
    tail: EnvelopeEntry[];
    mode: "encrypted";
    mediaNeeded?: string[];
    /** Relay per-blob size limit — see the `welcome` field of the same
     * name. Published identically in encrypted rooms: it is server config,
     * not room content, so a blind sequencer can state it in the clear. */
    mediaMaxBlobBytes?: number;
}
/** Encrypted-mode broadcast: sequenced opaque envelopes. */
 | {
    t: "broadcast-enc";
    entries: EnvelopeEntry[];
}
/** Relayed gossip blob; `from` is the sender's BOUND clientId. */
 | {
    t: "gossip";
    from: string;
    iv: string;
    ciphertext: string;
}
/** Media re-supply control (doc 16 §4). request = broadcast "who has
 * it"; upload = to ONE chosen holder; ready/unavailable = to waiters. */
 | {
    t: "media-request";
    sha: string;
} | {
    t: "media-upload";
    sha: string;
} | {
    t: "media-ready";
    sha: string;
} | {
    t: "media-unavailable";
    sha: string;
}
/** Checkpointer designation (doc 13 §3): the SERVER assigns the role —
 * the v1 lowest-clientId election was riggable (round-4 blocker 2); an
 * assigned role can only be held by an authenticated connection the
 * server picked. Rotated on disconnect. */
 | {
    t: "checkpointer";
    active: boolean;
}
/** `participant` is the sender's bound clientId (round-4 F14) — the same
 * identifier intents carry — so presence joins the roster/attribution
 * keyspace and survives the sender reconnecting on a new socket. */
 | {
    t: "presence";
    participant: string;
    position: PresencePosition | SealedPresence | null;
}
/** Full roster snapshot, fanned out on every join/leave/profile change
 * (rooms are small; a snapshot beats delta bookkeeping). Ephemeral like
 * presence: never logged, never persisted, dies with the room. */
 | {
    t: "roster";
    roster: RosterEntry[];
}
/**
 * The session is going to END, and this is the grace period before it does
 * (server lifecycle arc). Sent once per armed deadline to every connection
 * in the room so the UI can run a countdown — a session that vanishes
 * without warning is indistinguishable from a crash, and the whole point of
 * these timeouts is that they are POLICY, visibly applied.
 *
 *  - `idle`     nobody has done anything for the idle window. CANCELLABLE:
 *               any qualifying activity (a submit, an admin action, a media
 *               transfer, someone joining — never presence) resets the clock
 *               and produces `session-warning-cleared`.
 *  - `lifetime` the room has reached its absolute age cap. NOT cancellable
 *               by anything; the deadline was fixed when the room was
 *               created. Ending it does not end the document — any holder
 *               re-seeds from their bundle into a NEW epoch.
 *
 * `inMs` is the ACTUAL remaining time when the frame was built, not the
 * configured lead time: warnings are emitted by a periodic sweep, so the
 * real remainder is up to one sweep interval less than the configured
 * value, and a client counting down from the constant would lie.
 */
 | {
    t: "session-warning";
    reason: "idle" | "lifetime";
    inMs: number;
}
/**
 * A previously warned deadline is no longer approaching — the countdown
 * should disappear. Only ever `idle`: the lifetime cap cannot be cancelled,
 * and narrowing the type here means a consumer cannot write a branch the
 * server is incapable of producing.
 */
 | {
    t: "session-warning-cleared";
    reason: "idle";
}
/**
 * The answer to a client `ping`, echoing its nonce. Proof that the socket is
 * end-to-end alive RIGHT NOW — not merely that the local OS still believes
 * in it, which is all `readyState === OPEN` ever establishes.
 *
 * Emitted unconditionally and statelessly: no room lookup, no auth check, no
 * activity bookkeeping. A liveness answer that could itself be refused would
 * be indistinguishable from the death it is supposed to detect.
 */
 | {
    t: "pong";
    nonce: number;
} | {
    t: "refused";
    reason: string;
};

interface ToolbarMenuSelectOption {
    value: string;
    label: React.ReactNode;
    group?: string;
    disabled?: boolean;
    fontFamily?: string;
}
interface ToolbarMenuSelectProps {
    value: string;
    options: ToolbarMenuSelectOption[];
    onChange: (value: string) => void;
    placeholder?: React.ReactNode;
    title?: string;
    ariaLabel?: string;
    triggerAriaLabel?: string;
    width?: number | string;
    menuWidth?: number;
    className?: string;
    style?: React.CSSProperties;
}
/** CSS-overridable replacement for visible native selects. An inert,
 * transparent select is kept as an event bridge for existing integrations;
 * every user-facing part is a button/listbox rendered by us. */
declare function ToolbarMenuSelect({ value, options, onChange, placeholder, title, ariaLabel, triggerAriaLabel, width, menuWidth, className, style, }: ToolbarMenuSelectProps): react.JSX.Element;
/** Toolbar control groups a host can disable via the `features` prop. */
type ToolbarFeature = "history" | "styles" | "font" | "size" | "format" | "color" | "highlight" | "alignment" | "indent" | "spacing" | "link" | "lists" | "table" | "image" | "icon" | "screenshot" | "model3D" | "media" | "object" | "chart" | "smartArt" | "comment" | "footnote" | "bookmark" | "crossReference" | "dateTime" | "field" | "equation" | "symbol" | "shape" | "divider" | "textBox" | "wordArt" | "drawing" | "arrange" | "dropCap" | "headerFooter" | "coverPage" | "pageNumber" | "break" | "layout" | "help" | "download";
/**
 * Every `DocxViewApi` command that INSERTS content at the caret, paired with
 * the toolbar feature group that gates its control.
 *
 * A single source of truth in the same spirit as core's
 * `SELECTED_OBJECT_COMMANDS`: the collab audit (react/test/insert-commands
 * INVARIANT C) reads this list and the real `COLLAB_TOOLBAR_DEFAULTS` gate
 * rather than restating either, so "what the toolbar offers in a room" and
 * "what the audit checks" cannot drift apart.
 *
 * THE RULE a new entry signs up to: in collab mode an insert command must
 * EMIT an intent, or its feature must be gated off so no control exists to
 * press. An absent button is honest; a present button that mutates the local
 * document without emitting forks the room silently — the exact shape that
 * shipped in `insertImage` and stayed green through the whole capability
 * matrix, because the matrix's collab mount withholds `submitOp` and so
 * never exercises an insert command's collab path.
 */
interface InsertCommandSpec {
    /** The `DocxViewApi` method name. */
    command: string;
    /** Toolbar group whose gate decides whether a control is offered. */
    feature: ToolbarFeature;
}
declare const INSERT_COMMANDS: readonly InsertCommandSpec[];
type ToolbarMode = "simple" | "advanced";
interface DocxToolbarProps {
    api: DocxViewApi | null;
    onSave?: (bytes: Uint8Array) => void;
    /** Simple shows basic Home editing; advanced adds the Insert, Draw, and Layout ribbons. */
    mode?: ToolbarMode;
    /** Per-group overrides; every group defaults to enabled. */
    features?: Partial<Record<ToolbarFeature, boolean>>;
    /** Extra class on the toolbar root (e.g. a scope for CSS-variable overrides). */
    className?: string;
    /** Inline overrides merged onto the toolbar root; wins over the defaults. */
    style?: React.CSSProperties;
}
declare function DocxToolbar({ api, onSave, mode, features, className, style, }: DocxToolbarProps): react.JSX.Element;

interface DocxViewApi {
    /** Apply character formatting to the current browser selection. */
    applyFormat(patch: RunFormatPatch): void;
    /** Create a review comment on the current selection. False if no selection. */
    addComment(text: string): boolean;
    /** Insert a footnote at the caret. False without a caret. */
    addFootnote(text: string): boolean;
    /** Insert a dynamic page-number field at the caret (body, header or footer). */
    insertPageNumber(kind?: "page" | "pageOfTotal"): boolean;
    /** Insert any Word field instruction supported by the renderer. */
    insertField(instruction: string, cachedResult?: string): boolean;
    /** Insert a live DATE or TIME field with an optional Word date picture. */
    insertDateTime(kind: "date" | "time", picture?: string): boolean;
    /** Named bookmarks in document order. */
    listBookmarks(): string[];
    /** Add a bookmark around the selection, or a zero-length bookmark at the caret. */
    addBookmark(name: string): boolean;
    /** Insert a live text or page reference to a bookmark. */
    insertCrossReference(bookmark: string, kind: "text" | "page"): boolean;
    /** Insert an editable inline equation from WordInWeb's linear math syntax. */
    insertEquation(linear: string): boolean;
    /** Insert a Unicode symbol through the normal undo/suggestion-aware typing path. */
    insertSymbol(symbol: string): boolean;
    /** Insert a floating editable DrawingML shape at the caret. */
    insertShape(preset: ShapePreset, text?: string, lineStyle?: {
        color: string;
        width: number;
        dash: DrawingLineDash;
    }): boolean;
    /** Insert editable DrawingML WordArt at the caret. */
    insertWordArt(text: string, preset?: WordArtPreset): boolean;
    /** Insert a native editable ChartML chart at the caret. */
    insertChart(data: ChartData): boolean;
    /** Replace the selected native chart's type and data. */
    updateSelectedChart(data: ChartData): boolean;
    /** Data for the selected native chart, or null when another object is selected. */
    getSelectedChart(): ChartData | null;
    /** Insert a native editable SmartArt diagram at the caret. */
    insertSmartArt(data: SmartArtData): boolean;
    /** Replace the selected native SmartArt diagram's layout and node text. */
    updateSelectedSmartArt(data: SmartArtData): boolean;
    /** Data for the selected native SmartArt diagram, or null when another object is selected. */
    getSelectedSmartArt(): SmartArtData | null;
    /** Text formatting for the selected SmartArt node, or the first node when the group is selected. */
    getSelectedSmartArtTextFormat(): SmartArtTextFormat | null;
    /** Format the selected SmartArt node, or every node when the group is selected. */
    setSelectedSmartArtTextFormat(patch: Partial<SmartArtTextFormat>): boolean;
    /** Insert a native Office 3D model with an optional custom poster image. */
    insertModel3D(file: Blob, poster?: Blob): Promise<boolean>;
    /** Insert Word online-video metadata with a browser-safe poster. */
    insertOnlineVideo(url: string): Promise<boolean>;
    /** Embed an arbitrary file as a native OLE Package object. */
    insertEmbeddedObject(file: Blob, filename?: string): Promise<boolean>;
    /** Activate a freehand pen, or return to selection mode with null. */
    setDrawingTool(tool: DrawingTool | null): void;
    /** Current freehand pen, or null while in selection mode. */
    getDrawingTool(): DrawingTool | null;
    /** Align, rotate, or reorder the selected image, shape, or ink group. */
    arrangeObject(action: ObjectArrangeAction): boolean;
    /** True while an image, shape, or ink group is selected. */
    hasSelectedObject(): boolean;
    /** Kind of the selected object, used by contextual formatting controls. */
    getSelectedObjectContext(): {
        kind: SelectedObjectKind;
        canEditText: boolean;
        smartArtNodeSelected?: boolean;
        smartArtNodeIndex?: number;
    } | null;
    /** Run a formatting command against the current selected object. */
    runSelectedObjectCommand(command: SelectedObjectCommand): boolean;
    undo(): void;
    redo(): void;
    canUndo(): boolean;
    canRedo(): boolean;
    /** Insert a rows×cols table at the caret's paragraph. */
    insertTable(rows: number, cols: number): void;
    /** Row/column/table operations on the table containing the caret. */
    tableOp(op: TableOp): void;
    /** Current table-cell fill, undefined when the caret is outside a table. */
    getTableCellFill(): string | null | undefined;
    /** Insert an image file at the caret (inline, natural size clamped to column).
     * The result is REPORTED rather than swallowed: a picker that accepts a file
     * and then does nothing is indistinguishable from a broken button. */
    insertImage(file: Blob): Promise<ImageInsertResult>;
    /** What this document's image picker should ADVERTISE, as a file-input
     * `accept` string — narrower in a shared document, where the format has to
     * survive the wire. Derived from the same list the insert guard consults, so
     * the picker cannot offer something the insert will refuse. */
    imageAccept(): string;
    /** The relay's published per-image byte limit, or null when there is none to
     * show (a local document, or a server that publishes no limit). Null means
     * "say nothing about size" — never substitute a default. */
    imageMaxBytes(): number | null;
    /** Capture a screen, window, or browser tab and insert the current frame as a PNG picture. */
    insertScreenshot(): Promise<ScreenshotInsertResult>;
    /** Align the paragraph(s) under the caret or selection. */
    setAlignment(align: ParagraphAlignment): void;
    /** Apply a named paragraph style (null clears back to Normal). */
    setParagraphStyle(styleId: string | null): void;
    /** Toggle bulleted/numbered list on the paragraph(s) under the selection. */
    toggleList(kind: "bullet" | "number"): void;
    /** Current list kind at the caret ("bullet" | "number" | null). */
    getListType(): "bullet" | "number" | null;
    /** Link the selection to a URL; null removes the link at the caret. */
    setLink(url: string | null): void;
    /** URL of the hyperlink at the caret/selection, or null. */
    getLinkAt(): string | null;
    /** Step paragraph indent by half an inch (Word's indent buttons). */
    adjustIndent(direction: 1 | -1): void;
    /** Line spacing multiple or exact point height, and/or space before/after (points). */
    setParagraphSpacing(patch: {
        lineMultiple?: number;
        exactLinePt?: number;
        beforePt?: number | null;
        afterPt?: number | null;
    }): void;
    /** Create, customize, or remove the bottom-border divider on the selected paragraph(s). */
    setParagraphDivider(divider: ParagraphDivider | null): boolean;
    /** Direct bottom-border divider on the caret paragraph. */
    getParagraphDivider(): ParagraphDivider | null;
    /** Apply or remove a native Word drop cap on the caret paragraph. */
    setDropCap(mode: "drop" | "margin" | null, lines?: number): boolean;
    /** Remove direct character formatting from the selection. */
    clearFormatting(): void;
    /** Change the selection's case. */
    changeCase(mode: "upper" | "lower" | "title"): void;
    /** Find matches for a query; selects the first and returns the count. */
    find(query: string, opts?: {
        matchCase?: boolean;
    }): number;
    /** Select the next/previous match; returns 1-based index or 0. */
    findStep(delta: 1 | -1): number;
    /** Replace the current match; returns remaining match count. */
    replaceCurrent(replacement: string): number;
    /** Replace every match; returns how many were replaced. */
    replaceAll(query: string, replacement: string): number;
    /** Paragraph styles for the style menu (declared + Word built-ins). */
    listParagraphStyles(): {
        id: string;
        name: string;
    }[];
    /** pStyle id of the caret paragraph (null = Normal). */
    getParagraphStyleId(): string | null;
    /** Change margins / page size / orientation (inches). */
    setPageLayout(patch: PageLayoutPatch, scope?: "document" | "section"): void;
    /** One-based logical section containing the caret or selection. */
    getSectionContext(): {
        index: number;
        count: number;
    } | null;
    /** Insert a page/column break or a section break at the caret. */
    insertBreak(kind: "page" | "column" | "sectionNextPage" | "sectionContinuous"): boolean;
    /** Insert a full blank page at the caret (two consecutive page breaks). */
    insertBlankPage(): boolean;
    /** Insert an editable cover page before the current document. */
    insertCoverPage(content: CoverPageContent): boolean;
    /** Toggle/configure margin line numbers (Word's Layout > Line Numbers).
     * scope "section" targets the caret's section; "document" every section. */
    setLineNumbering(patch: LineNumberingPatch, scope?: "document" | "section"): void;
    /** Current line-numbering settings for the caret's section, or null (off). */
    getLineNumbering(): {
        countBy: number;
        restart: "continuous" | "newPage" | "newSection";
        start: number;
    } | null;
    /** Leave header/footer editing mode. */
    closeHeaderFooter(): void;
    /** Enter and, if needed, create the header or footer on the caret's page. */
    openHeaderFooter(kind: "header" | "footer"): boolean;
    /** Effective formatting of the current selection (toolbar state), or null. */
    getSelectionFormat(): SelectionFormat | null;
    /** Print the rendered pages (browser print dialog / save as PDF). */
    print(): void;
    /** Serialize the (edited) document back to .docx bytes. */
    save(): Uint8Array;
    /** Page count after the latest layout. */
    pageCount(): number;
    /**
     * Suggesting mode: when on, edits record as OOXML tracked changes (w:ins /
     * w:del) instead of mutating text directly, and the view switches to markup
     * so the suggestion shows live. `author` stamps the revision (defaults to the
     * commentAuthor prop). Turning it off restores the prior revision view.
     */
    setSuggesting(on: boolean, author?: string): void;
    isSuggesting(): boolean;
    /** Accept the tracked change at the caret (keep insertion / apply deletion).
     * Refuses (false, no mutation) in a live collab session — resolving a
     * suggestion emits no intent yet, so it would fork the room. */
    acceptRevisionAtCaret(): boolean;
    /** Reject the tracked change at the caret (drop insertion / restore
     * deletion). Refuses in a live collab session, like accept. */
    rejectRevisionAtCaret(): boolean;
    /** How many tracked changes (suggestions) the document currently holds. */
    revisionCount(): number;
    /** Accept every tracked change (one undo step). Returns how many applied.
     * Returns 0 without mutating in a live collab session, like accept. */
    acceptAllRevisions(): number;
    /** Reject every tracked change (one undo step). Returns how many applied.
     * Returns 0 without mutating in a live collab session, like accept. */
    rejectAllRevisions(): number;
    /** Current caret as stable-id addresses (collab), or null. The encoding
     * survives a reconciliation reload, so it can be captured from a view
     * about to remount and restored into its replacement. */
    getEncodedCaret(): EncodedCaret | null;
    /** Restore a caret captured by getEncodedCaret. False when the position no
     * longer resolves (or outside collab mode). */
    setCaretFromEncoded(pos: EncodedCaret): boolean;
    /**
     * Scroll the view so `participant`'s live presence caret is on screen —
     * "jump to that person". A pure VIEW operation: it never relayouts, marks
     * nothing dirty, and mutates no document state. The jump is instant, never
     * smooth: the target can be hundreds of virtualized pages away, and a
     * smooth scroll across that distance forces continuous page mounting;
     * instant also satisfies prefers-reduced-motion by construction.
     *
     *  - "revealed"    scrolled to the caret (or, when the exact run was
     *                  deleted since it was reported, to the start of its
     *                  paragraph — the same fallback caret restore uses).
     *  - "no-position" that participant has no broadcast cursor right now
     *                  (never placed one, cleared it, or left).
     *  - "unresolved"  the position no longer maps to content in this replica.
     */
    revealPresence(participant: string): "revealed" | "no-position" | "unresolved";
    document: DocxDocument;
}
type ScreenshotInsertResult = "inserted" | "unsupported" | "cancelled" | "error" | "no-caret";
/**
 * Image formats a SHARED document accepts, and the single source of truth for
 * both halves of that promise: the guard that declines an intent and the
 * `accept` attribute of the file picker that offers one.
 *
 * They were separately written constants once, and drifted — the picker
 * advertised SVG while the guard refused it, so choosing an SVG in a shared
 * document did nothing at all, with no skeleton, no message and no log. The
 * user's report was "it just didn't work and then disappeared". Deriving the
 * offer from the acceptance is what makes that drift unrepresentable.
 *
 * Raster-only because the wire allowlist is (collab/src/validate.ts): widening
 * it is an intent-shape change with an ENGINE_VERSION bump, not a UI edit.
 */
declare const COLLAB_IMAGE_EXTS: readonly ["png", "jpg", "jpeg", "gif", "bmp", "webp"];
/**
 * Why an image insert did not happen, so the caller can SAY so. Every one of
 * these was a bare `return` until a user hit the `unsupported-format` path and
 * had no way to tell a rejected file from a broken button.
 */
type ImageInsertResult = "inserted" | "no-caret"
/** Not a format a shared document can carry (SVG today). */
 | "unsupported-format"
/** Bigger than the relay's published per-blob limit. Refused BEFORE any
 * bytes were read, sealed, hashed or sent — see `imageMaxBytes()` for the
 * number to show the user. */
 | "too-large"
/** Collab is wired but no media relay is configured — nothing can be
 * uploaded, and inserting locally would fork the room. */
 | "no-relay"
/** The relay refused the bytes; nothing was reserved, nothing forked. */
 | "upload-failed"
/** The bytes did not decode as an image at all. */
 | "error";
interface DocxViewProps {
    /** The document: raw bytes, a File/Blob, or a URL to fetch. */
    source: ArrayBuffer | Uint8Array | Blob | string;
    /** Zoom factor, 1 = 100%. */
    zoom?: number;
    /**
     * Fit-to-width (Google-Docs mobile behavior): when the page is wider than the
     * viewport, auto-scale down so it fits with a small gutter and never scrolls
     * horizontally. The computed scale drives the real `zoom` (crisp text, not a
     * blurry transform) and is capped at `zoom`, so a wide desktop viewport is
     * unchanged. Recomputed on container resize. Default true.
     */
    fitWidth?: boolean;
    /**
     * Container width (px) at or below which the chrome switches to its compact
     * phone/tablet treatment — the comment rail collapses to tap-to-open cards so
     * balloons never force horizontal scroll. Default 820.
     */
    narrowWidth?: number;
    /**
     * Enable editing commands (selection-based formatting, save-back).
     * Default false: pure render-only viewer.
     */
    editable?: boolean;
    className?: string;
    style?: React.CSSProperties;
    onLoad?: (info: {
        pageCount: number;
        document: DocxDocument;
    }) => void;
    /** Fires whenever editing changes the rendered page count. */
    onPageCountChange?: (pageCount: number) => void;
    /** Fires when the document is ready; the api is only usable while mounted. */
    onReady?: (api: DocxViewApi) => void;
    onError?: (error: Error) => void;
    /**
     * Optional collaborative session (from `wordinweb/collab`'s `useCollab`).
     * When provided, the editor forwards each local edit as an intent via
     * `collab.submit`. Typed structurally so the main `wordinweb` bundle carries
     * no runtime dependency on the collab engine (plan doc 07 tree-shaking) —
     * the app imports the session from the separate `wordinweb/collab` entry and
     * injects it here.
     */
    collab?: {
        /** clientId → display name for caret flags (doc 14 §2); text-node rendered. */
        participantNames?: Record<string, string>;
        submit: (intent: EditorIntent) => void;
        /** Remote participants' cursor/selection positions, drawn as colored
         * carets over the page (see presence-cursors). */
        presence?: Record<string, PresencePosition | null>;
        /** Allocate `n` fresh carried node ids (for sub-range format / split /
         * insert intents). Injected from the collab connection. */
        allocIds?: (n: number) => number[];
        /** The live reconciled document object to render DIRECTLY (skip the
         * bytes → parse round-trip). The collab replica mutates this same instance
         * in place on each broadcast; DocxView repaints it when `renderSignal`
         * bumps, so a remote edit costs one repaint — no re-serialize, no re-parse,
         * no caret reset. When present, `source` is only a placeholder. */
        doc?: DocxDocument;
        /** Monotonic counter that bumps whenever `doc` was mutated in place; a
         * change triggers an in-place repaint of `doc`. */
        renderSignal?: number;
        /** Drain the union of the dirty scopes behind `renderSignal` since the
         * last take. A narrow scope lets the repaint relayout one paragraph
         * incrementally (the same path local typing takes) instead of the whole
         * document; `doc` scope (or an absent method) keeps the whole-document
         * repaint; null means nothing is dirty and the repaint is skipped.
         * Consumed at the repaint so a coalesced repaint covers every batched
         * remote intent. */
        takeRenderScope?: () => {
            kind: "doc";
        } | {
            kind: "block";
            blocks: XmlElement[];
        } | {
            kind: "split";
            before: XmlElement;
            after: XmlElement;
        } | null;
        /** Broadcast the local caret so remote participants draw this user's
         * cursor. Called with the caret's stable-id address on every caret move
         * (null when the caret leaves id-tracked content). */
        setPresence?: (pos: PresencePosition | null) => void;
        /** Submit a toolbar/API operation NOT yet applied locally. The connection
         * applies it optimistically through the same canonical code the server
         * runs, so the local result is byte-identical to every replica. When set,
         * DocxView routes its imperative commands (insert chart/table/equation,
         * set link/page layout, comments, ...) through this instead of mutating
         * the document itself. */
        submitOp?: (intent: {
            kind: string;
        } & Record<string, unknown>) => void;
        /** Upload image bytes to the media relay and return the address fields
         * the insertImage intent must carry (plan doc 16 §5.1). Null means the
         * relay REFUSED — the caller must then not reserve anything, or the room
         * gets a skeleton nobody can ever fill. Absent when the app supplied no
         * relay origin, in which case images stay a local-only feature. */
        uploadMedia?: (bytes: Uint8Array) => Promise<{
            blobSha: string;
            bytesLen: number;
            iv?: string;
        } | null>;
        /**
         * Largest single upload the RELAY will accept, in bytes, as published in
         * the welcome. Lets the insert refuse an oversized file locally instead of
         * discovering it after sealing, hashing and a full upload.
         *
         * NULL MEANS SKIP THE CHECK — not "no limit" and not "use a default". An
         * older server publishes nothing, and a client that invents a number
         * either blocks uploads the server would have taken or promises the user
         * one it will refuse. The server enforces the real limit either way, so
         * skipping is safe and guessing is not.
         */
        mediaMaxBlobBytes?: number | null;
        /** Reverse this user's last SEQUENCED action (plan doc 03 Phase 8). The
         * editor routes Cmd+Z here in a room, because replaying the LOCAL history
         * stack would edit this replica with nothing on the wire. Absent ⇒ undo
         * declines rather than mutating. */
        undoLast?: () => void;
    };
    /** Author name stamped on comment replies (default "You"). */
    commentAuthor?: string;
    /** Render review comments (range highlights + margin balloons). Default true. */
    showComments?: boolean;
    /** Tracked-changes display: "final" (default) or "markup". */
    revisions?: "final" | "markup";
    /** Fires after render with document-requested font faces the browser cannot
     * render (unavailable, or lacking the document's script) — the page is
     * silently substituting and may differ from Word. Empty array = all good. */
    onMissingFonts?: (missing: MissingFont[]) => void;
}
/**
 * High-fidelity paginated DOCX viewer (and, with `editable`, editor).
 *
 * ```tsx
 * <DocxView source="/report.docx" />                          // render-only
 * <DocxView source="/report.docx" editable onReady={setApi} /> // editing
 * ```
 */
declare function DocxView({ source, zoom, fitWidth, narrowWidth, editable, className, style, onLoad, onPageCountChange, onReady, onError, commentAuthor, showComments, revisions, onMissingFonts, collab, }: DocxViewProps): react.JSX.Element;

export { type ShapePreset as A, ToolbarMenuSelect as B, type ClientMessage as C, DocxDocument as D, type ToolbarMenuSelectOption as E, type ToolbarMenuSelectProps as F, layoutDocument as G, printPages as H, type Intent as I, renderToDom as J, type LineageHead as L, type PresencePosition as P, type RosterEntry as R, type ServerMessage as S, type ToolbarFeature as T, type UndoOutcome as U, type WriteStatus as W, type XmlElement as X, type ParticipantProfile as a, type IdSidecar as b, type ToolbarMode as c, type DocxViewApi as d, bytesToB64 as e, deriveEpochKeys as f, docKeyFromFragment as g, stretchShareCode as h, COLLAB_IMAGE_EXTS as i, type CoverPageContent as j, DocxToolbar as k, type DocxToolbarProps as l, mintDocKey as m, DocxView as n, type DocxViewProps as o, type DrawingTool as p, INSERT_COMMANDS as q, type ImageInsertResult as r, sealCheckpoint as s, type InsertCommandSpec as t, type LineNumberingPatch as u, type PageLayoutPatch as v, type ParagraphAlignment as w, type RunFormatPatch as x, type ScreenshotInsertResult as y, type SelectionFormat as z };
