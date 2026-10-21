import { DocxDocument as DocxOnlyDocument, type FromFilesOptions } from "./docx-document.js";
import { Package, holdXmlCounts, mainDocumentPart } from "./zip.js";
import { normalizeStrictPackage } from "./strict.js";
import { docPackageFiles } from "./doc/hooks.js";
import { isOleFile } from "./doc/streams.js";
import { isOldWordFile, oldWordPackageFiles } from "./oldword/hooks.js";

export type { FromFilesOptions, NotePart } from "./docx-document.js";

/** Word documents, including binary .doc and the oldest Word formats. */
export class DocxDocument extends DocxOnlyDocument {
  private constructor(pkg: Package, options: FromFilesOptions = {}) {
    super(pkg, options);
  }

  static override load(data: ArrayBuffer | Uint8Array): DocxDocument {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    const pkg = isOleFile(bytes) ? Package.fromFiles(docPackageFiles(bytes)) : isOldWordFile(bytes) ? Package.fromFiles(oldWordPackageFiles(bytes)) : Package.from(bytes);
    normalizeStrictPackage(pkg.raw());
    holdXmlCounts(pkg.raw(), mainDocumentPart(pkg.binary("_rels/.rels")));
    return new DocxDocument(pkg);
  }

  static override fromFiles(files: Record<string, Uint8Array>, options: FromFilesOptions = {}): DocxDocument {
    return new DocxDocument(Package.fromFiles(files), options);
  }

  override clone(): DocxDocument {
    return new DocxDocument(Package.fromFiles(this.saveFiles()));
  }
}
