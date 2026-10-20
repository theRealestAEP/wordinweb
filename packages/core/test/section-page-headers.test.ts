import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, W_NS } from "./helpers.js";

// A page prints the header and footer of the section at its top. A section
// that holds only its empty break paragraph still opens a page; govdocs
// 068399 is the case found: an empty titlePg section 1 carries the letterhead
// in its first-page header, and the continuous section 2 holds the text.

const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const KINDS = ["first", "default", "even"] as const;

type Kind = (typeof KINDS)[number];

interface SectionSpec {
  /** Header and footer variants this section declares, named S<n><KIND>HEADER/FOOTER. */
  parts?: readonly Kind[];
  type?: "continuous" | "nextPage";
  titlePg?: boolean;
  /** Body paragraphs before the section's break; zero makes an empty section. */
  paragraphs?: number;
  landscape?: boolean;
  margin?: number;
}

function sectPr(n: number, s: SectionSpec): string {
  const refs = (s.parts ?? []).flatMap((kind) => [
    `<w:headerReference w:type="${kind}" r:id="h${n}${kind}"/>`,
    `<w:footerReference w:type="${kind}" r:id="f${n}${kind}"/>`,
  ]);
  const size = s.landscape ? `<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>` : `<w:pgSz w:w="12240" w:h="15840"/>`;
  const side = s.margin ?? 1440;
  return (
    `<w:sectPr>${refs.join("")}` +
    (s.type ? `<w:type w:val="${s.type}"/>` : "") +
    size +
    `<w:pgMar w:top="1440" w:right="${side}" w:bottom="1440" w:left="${side}" w:header="720" w:footer="720" w:gutter="0"/>` +
    (s.titlePg ? "<w:titlePg/>" : "") +
    `</w:sectPr>`
  );
}

function build(sections: SectionSpec[], evenAndOdd = false) {
  const parts: Record<string, string> = {};
  const rels: string[] = [];
  let body = "";
  sections.forEach((s, i) => {
    const n = i + 1;
    for (const kind of s.parts ?? []) {
      for (const [tag, prefix, part] of [["hdr", "h", "header"], ["ftr", "f", "footer"]] as const) {
        const name = `${part}${n}${kind}.xml`;
        const marker = `S${n}${kind.toUpperCase()}${part.toUpperCase()}`;
        parts[`word/${name}`] = `<w:${tag} ${W_NS}><w:p><w:r><w:t>${marker}</w:t></w:r></w:p></w:${tag}>`;
        rels.push(`<Relationship Id="${prefix}${n}${kind}" Type="${REL}/${part}" Target="${name}"/>`);
      }
    }
    for (let k = 1; k <= (s.paragraphs ?? 0); k++) body += `<w:p><w:r><w:t>Body ${k} of section ${n}.</w:t></w:r></w:p>`;
    if (i < sections.length - 1) body += `<w:p><w:pPr>${sectPr(n, s)}</w:pPr></w:p>`;
    else body += sectPr(n, s);
  });
  parts["word/document.xml"] = `<w:document ${W_NS} ${R_NS}><w:body>${body}</w:body></w:document>`;
  parts["word/_rels/document.xml.rels"] =
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
  if (evenAndOdd) parts["word/settings.xml"] = `<w:settings ${W_NS}><w:evenAndOddHeaders/></w:settings>`;
  const doc = DocxDocument.load(makeDocx(parts));
  return layoutDocument(doc, { measurer: new ApproxMeasurer() });
}

/** Per page: the header and footer markers, then the first body line, if any. */
function pages(result: ReturnType<typeof layoutDocument>): string[] {
  return result.pages.map((page) => {
    const texts = page.items.map((item) => (item.kind === "text" ? item.text.trim() : "")).filter((t) => t);
    const markers = texts.filter((t) => /^S\d[A-Z]+(HEADER|FOOTER)$/.test(t));
    const body = /Body (\d+)/.exec(texts.join(" "));
    return [...markers, body ? body[0] : "blank"].join(" ");
  });
}

const BODY = 120;

describe("headers and footers of a page opened by an empty section", () => {
  it("an empty titlePg section before a continuous section keeps its first-page header on page 1", () => {
    const result = build([
      { parts: ["first", "default"], titlePg: true },
      { parts: ["first", "default"], type: "continuous", paragraphs: BODY },
    ]);
    const got = pages(result);
    expect(got[0]).toBe("S1FIRSTHEADER S1FIRSTFOOTER Body 1");
    // Later pages are section 2's own, and not its first page.
    expect(got.length).toBeGreaterThan(1);
    for (const page of got.slice(1)) expect(page).toMatch(/^S2DEFAULTHEADER S2DEFAULTFOOTER Body \d+$/);
  });

  it("with text in section 1, page 1 is the same (control)", () => {
    const result = build([
      { parts: ["first", "default"], titlePg: true, paragraphs: 1 },
      { parts: ["first", "default"], type: "continuous", paragraphs: BODY },
    ]);
    const got = pages(result);
    expect(got[0]).toBe("S1FIRSTHEADER S1FIRSTFOOTER Body 1");
    for (const page of got.slice(1)) expect(page).toMatch(/^S2DEFAULTHEADER S2DEFAULTFOOTER Body \d+$/);
  });

  it("a section that inherits the headers keeps the first page's variant", () => {
    const result = build([
      { parts: ["first", "default"], titlePg: true },
      { type: "continuous", paragraphs: BODY },
    ]);
    const got = pages(result);
    expect(got[0]).toBe("S1FIRSTHEADER S1FIRSTFOOTER Body 1");
    for (const page of got.slice(1)) expect(page).toMatch(/^S1DEFAULTHEADER S1DEFAULTFOOTER Body \d+$/);
  });

  it("an empty section before a next-page section prints a blank page 1 with its header", () => {
    const result = build([
      { parts: ["first", "default"], titlePg: true },
      { parts: ["first", "default"], type: "nextPage", paragraphs: BODY },
    ]);
    const got = pages(result);
    expect(got[0]).toBe("S1FIRSTHEADER S1FIRSTFOOTER blank");
    expect(got[1]).toBe("S2DEFAULTHEADER S2DEFAULTFOOTER Body 1");
  });

  it("a continuous section with another paper size starts page 2, after the empty section's page", () => {
    const result = build([
      { parts: ["first", "default"], titlePg: true },
      { parts: ["first", "default"], type: "continuous", paragraphs: BODY, landscape: true },
    ]);
    const got = pages(result);
    expect(got[0]).toBe("S1FIRSTHEADER S1FIRSTFOOTER blank");
    expect(result.pages[0].width).toBe(816);
    expect(got[1]).toBe("S2DEFAULTHEADER S2DEFAULTFOOTER Body 1");
    expect(result.pages[1].width).toBe(1056);
  });

  it("a continuous section with other margins shares page 1; the header keeps section 1's margins", () => {
    const result = build([
      { parts: ["first", "default"], titlePg: true },
      { parts: ["first", "default"], type: "continuous", paragraphs: BODY, margin: 2880 },
    ]);
    const got = pages(result);
    expect(got[0]).toBe("S1FIRSTHEADER S1FIRSTFOOTER Body 1");
    const first = result.pages[0].items;
    const x = (text: string) => first.find((item) => item.kind === "text" && item.text.startsWith(text))?.x;
    expect(x("S1FIRSTHEADER")).toBeCloseTo(96, 0);
    expect(x("Body")).toBeCloseTo(192, 0);
    for (const page of got.slice(1)) expect(page).toMatch(/^S2DEFAULTHEADER S2DEFAULTFOOTER Body \d+$/);
  });

  it("several empty continuous sections in a row leave page 1 to the first", () => {
    const result = build([
      { parts: ["first", "default"], titlePg: true },
      { parts: ["first", "default"], type: "continuous" },
      { parts: ["default"], type: "continuous" },
      { parts: ["default"], type: "continuous", paragraphs: BODY },
    ]);
    const got = pages(result);
    expect(got[0]).toBe("S1FIRSTHEADER S1FIRSTFOOTER Body 1");
    for (const page of got.slice(1)) expect(page).toMatch(/^S4DEFAULTHEADER S4DEFAULTFOOTER Body \d+$/);
  });

  it("titlePg in the continuous section 2 does not reach page 1 or page 2", () => {
    const result = build([
      { parts: ["default"] },
      { parts: ["first", "default"], type: "continuous", titlePg: true, paragraphs: BODY },
    ]);
    const got = pages(result);
    expect(got[0]).toBe("S1DEFAULTHEADER S1DEFAULTFOOTER Body 1");
    // Section 2 starts on page 1, so page 2 is not its first page.
    for (const page of got.slice(1)) expect(page).toMatch(/^S2DEFAULTHEADER S2DEFAULTFOOTER Body \d+$/);
  });

  it("with evenAndOddHeaders, page 1 is section 1's first page and page 2 is section 2's even page", () => {
    const result = build(
      [
        { parts: ["first", "default", "even"], titlePg: true },
        { parts: ["first", "default", "even"], type: "continuous", paragraphs: BODY },
      ],
      true,
    );
    const got = pages(result);
    expect(got.length).toBeGreaterThan(2);
    expect(got[0]).toBe("S1FIRSTHEADER S1FIRSTFOOTER Body 1");
    expect(got[1]).toMatch(/^S2EVENHEADER S2EVENFOOTER Body \d+$/);
    expect(got[2]).toMatch(/^S2DEFAULTHEADER S2DEFAULTFOOTER Body \d+$/);
  });
});
