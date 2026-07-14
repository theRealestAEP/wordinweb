# Fixture Privacy Audit

Audit of every DocxInWeb test fixture (`apps/demo/public/fixtures/*.docx`,
`fixtures-staging/`, `parity/`) plus the Word reference PDFs, for real / identifying /
sensitive information before the fixtures are hosted on the public internet.

- **Scanner:** `scripts/audit-fixtures.py` (reusable; re-run before any publish).
- **Full machine table:** `docs/FIXTURE-AUDIT-TABLE.md` (every finding, 1400+ rows).
- **This file:** executive summary, the BLOCKER/WARN detail, media review, remediation.
- **Scope scanned:** 129 `.docx` + reference `.pdf` across the three dirs.
- **Severity totals:** **BLOCKER 45**, WARN 1111, REVIEW (media, manual) 148, OK 133.

Nothing was modified. Fixtures/PDFs are byte-for-byte untouched (changing bytes would
invalidate the raster cache; re-exporting references is a separate Word-automation task).

---

## What the sanitizer got right

The `wild-*` / `wild2-*` body-text sanitizer (real text -> pseudowords like `Bavoqe`)
**also covered the non-body text channels** — headers, footers, footnotes, endnotes and
comment *bodies* are all pseudowords. External hyperlink URLs were rewritten to
`example.com` (562 rels, all one host). No macros (`vbaProject.bin`) and no embedded
fonts (`.odttf`) exist in any fixture. OLE blobs are MathType equation objects only
(`wild2-math-eq-as-images`, `wild2-sci-ieee-2col`), no author payload.

**Every miss is in a channel the text sanitizer never touches:** package metadata,
custom (Outlook) properties, relationship targets, Word owner-lock temp files, PDF
metadata, and one embedded logo image.

---

## BLOCKERS (real PII / identifying info) — must fix before publish

### B1. Word owner-lock temp files embed "Alex Pickett" (and should not ship at all)
`~$*.docx` files are Word's owner-lock sidecars; their first bytes are the last opener's
username in cleartext. Six exist; **four are committed to git** (already pushed to the
private repo — doubly bad):

| File | Tracked in git? |
|---|---|
| `apps/demo/public/fixtures/~$rity-revisions.docx` | no (gitignored dir) |
| `apps/demo/public/fixtures/~$ronology.docx` | no (gitignored dir) |
| `fixtures-staging/~$ld2-legal-nih-contract.docx` | **YES** |
| `fixtures-staging/~$ld2-lit-yiddish-rtl.docx` | **YES** |
| `fixtures-staging/~$ld2-math-eq-as-images.docx` | **YES** |
| `fixtures-staging/~$ld2-med-phase23-protocol.docx` | **YES** |

All contain `Alex Pickett`.

### B2. Outlook "Send for Review" trail leaks a real person + real email
`wild2-sci-ieee-2col.docx` (present in **both** `apps/demo/public/fixtures/` and the
**git-tracked** `fixtures-staging/`) — `docProps/custom.xml`:
- `_AuthorEmail = jhnelson@tva.gov`  (real address; TVA = Tennessee Valley Authority, a US federal agency)
- `_AuthorEmailDisplayName = Nelson, Jeffrey H`
- `_EmailSubject = New Conference paper template and revised transactions paper template`
- `_AdHocReviewCycleID = 1976721638`

The body sanitizer never looked at `custom.xml`. This is the single most sensitive find:
a real named individual and a working government email, **already in the private repo.**

### B3. Real attorney email in hyperlink relationships
`pleading.docx`, `pleading-anon.docx`, and `w.docx` — `word/_rels/document.xml.rels`
contains `mailto:kathleenmessinger@quinnemanuel.com` (Quinn Emanuel is a real law firm).
Note `pleading-anon.docx` had its core-metadata scrubbed to `Fixture` but **the rels
hyperlink survived** — the anonymization missed the relationship layer.

### B4. Real litigation title / description in package metadata
`pleading.docx` and `w.docx` — `docProps/core.xml`:
- `dc:title = Plaintiff's Notice of Deposition of Dan Westgarth`  (real person named)
- `dc:description = Draft Rule 30(b)(1) notice of deposition for attorney review.`

(`pleading-anon.docx` has these scrubbed to `Fixture`; only its hyperlink (B3) leaks.)

### B5. "Alex Pickett" in `docProps/core.xml` creator / lastModifiedBy
Real author identity in package metadata of these fixtures:
`alexpickett.docx`, `forsale.docx`, `pickett.docx`, `probe3-emoji.docx`,
`probe3-index-xrefs.docx`, `probe3-linked-textboxes.docx`, `probe3-shape-autofit.docx`
(fixtures dir); and in `fixtures-staging/`: `Menu.docx`, `YN.docx`, `probe-ehfix.docx`,
`probe-ehfix2.docx`, `probe-emptyheading15.docx`, `probe3-emoji.docx`,
`probe3-index-xrefs.docx`, `probe3-linked-textboxes.docx`, `probe3-shape-autofit.docx`;
and in `parity/`: `Alex Pickett.docx`, `Character formatting stress.docx`,
`Multilevel list.docx`, `Nested tables.docx`, `Nested tables_2.docx`, `Section 1.docx`.

### B6. Embedded real institutional logo (letterhead)
`wild-hamburg.docx` `word/media/image1.jpeg` **is the real Universität Hamburg logo**
("Universität Hamburg — DER FORSCHUNG | DER LEHRE | DER BILDUNG"). A trademarked mark
that identifies the source institution; the image sanitizer left it. Replace the image
(and re-export the reference PDF, since this changes rendering).

### B7. Reference-PDF metadata names "Alex Pickett"
`parity/pickett-word.pdf` and `parity/probe3-lo-provenance-word.pdf` carry
`Author: Alex Pickett` in PDF info. (`parity/` is gitignored, not pushed.)

---

## WARN (metadata worth scrubbing — not personal PII, but leaks provenance)

Full list in `docs/FIXTURE-AUDIT-TABLE.md`. Highlights:

- **Synthetic author names on comments/tracked-changes** — `Reviewer`, `Reviewer A`,
  `Editor B`, `Ada Reviewer`, `Bob Editor` (`dense-comments`, `parity-comments`,
  `parity-revisions`, `probe3-tracked-changes`, `real`, `sample`,
  `wild2-legal-ca-agreement`). Not real people, but scrub for cleanliness.
- **Placeholder creator/company** — `Fixture`, `Un-named`, `python-docx`, `Admin`,
  `Cobbery`, `Microsoft` in core.xml / app.xml Company/Manager across many fixtures,
  and mirrored in the reference PDFs' `Author`/`Subject`/`Keywords` (`Fixture`,
  `Un-named`). Cosmetic.
- **SharePoint / DMS provenance GUIDs** in `custom.xml`: `wild-athabasca` and
  `wild-wirfp` carry `ContentTypeId`, `_dlc_DocIdItemGuid`, `MediaServiceImageTags` —
  identify the originating SharePoint tenant. `real.docx` carries a
  `GrammarlyDocumentId` GUID. `forsale.docx` / `fixtures-staging/Menu.docx` carry
  template-marketplace `AssetID` (`TF1000...`).
- **Source-document titles** in `app.xml TitlesOfParts`: `Template Abschlussarbeiten`
  (Hamburg), `CONTRACT Language Template` (NIH), `Protocol Template` (NCCIH),
  `DOI: 10...` (chem paper) — reveal the real source template/paper.
- **562 external rels -> `example.com`** and **59 `instrText` HYPERLINK URLs** — already
  sanitized to example.com; harmless, listed for completeness.
- **156 drawing alt-text `name`/`descr`** — mostly generic; none contained emails/names
  on inspection, but worth a glance in the table.

## REVIEW (media, manually inspected)

148 media entries / 71 distinct images. Visually reviewed a sample of every distinct
non-equation raster (contact sheet built during audit):

- **BLOCKER:** `wild-hamburg` image1.jpeg = real Universität Hamburg logo (see B6).
- **OK / low-risk:** `wild2-sci-ieee-2col` image4.png is the IEEE-template sample bio
  photo of **Nikola Tesla** (public-domain historical figure, not living PII). The other
  IEEE images are the template's point-size table and a magnetization graph. `wild-hamburg`
  image2 = generic "Marketing" clip-art. `wild-athabasca` = a generic bar chart.
  `wild2-sci-chem-omml` TIFFs = scientific figures (XRD/SEM/spectra). `forsale` = stock
  kayak/lake photo. `alexpickett` = "Hot Face" emoji PNG + a Microsoft GLTF "Hot_Face"
  3D-emoji `.glb` (generator string only, no PII). `parity2-coverpage` = geometric cover
  art. `dense-imagestress`/`dense-skewtest` = synthetic stress images.
- **Could not render (flag for manual check if paranoid):** EMF/WMF equation images and
  `wild2-med-phase23-protocol` `image1.emf` (vector, PIL cannot rasterize; size/shape
  consistent with an equation/diagram, not a photo).

---

## Remediation plan

Two classes of fix, by whether the change touches rendering:

### Class A — metadata-only scrubs (do NOT affect rendering; PDFs do NOT need re-export)
These edit `docProps/*.xml`, `word/_rels/*.rels`, or delete sidecar files. The rasterized
page output is identical, so the raster cache stays valid and reference PDFs stand.

1. **Delete the six `~$*.docx` owner files** (B1). They are Word temp files with zero test
   value. For the four tracked ones, `git rm` them from `fixtures-staging/` and add
   `~$*` / `~$*.docx` to `.gitignore`. (History rewrite/force-push to purge the pushed
   blobs is the coordinator's call; at minimum stop shipping them going forward.)
2. **Strip `docProps/custom.xml`** from `wild2-sci-ieee-2col.docx` (B2) — remove the
   `_Author*` / `_EmailSubject` / `_AdHocReviewCycleID` properties (or drop the part and
   its `[Content_Types]`/rels entry). Do the same for the tracked staging copy.
3. **Remove the `mailto:` external relationship** from `pleading.docx`, `pleading-anon.docx`,
   `w.docx` `document.xml.rels` (B3) — repoint to `example.com` or delete the hyperlink rel.
4. **Rewrite `docProps/core.xml`** to neutral placeholders across all B4/B5 files: set
   `dc:creator`, `cp:lastModifiedBy` -> `Fixture`; set `pleading`/`w` `dc:title` and
   `dc:description` -> `Fixture`. A one-pass script over all fixtures normalizing
   creator/lastModifiedBy/title/subject/description/keywords is the clean approach.
5. **Scrub WARN provenance** opportunistically in the same pass: SharePoint GUIDs,
   `GrammarlyDocumentId`, `AssetID`, `app.xml Company/Manager/TitlesOfParts`.
6. **PDF metadata (B7):** `parity/pickett-word.pdf`, `parity/probe3-lo-provenance-word.pdf`
   `Author=Alex Pickett`. `parity/` is gitignored (not pushed) so lower urgency, but scrub
   with `exiftool`/`qpdf` (metadata-only, no re-render) before any public hosting of PDFs.

> All Class-A docx edits rewrite one XML part inside the zip. To keep the raster cache
> valid you still want Word/coordinator sign-off, but semantically the layout is unchanged.
> Practically: since the coordinator owns byte changes, deliver them the list above and let
> them run the scrub + confirm cache.

### Class B — content change (DOES affect rendering; reference PDF MUST be re-exported)
1. **`wild-hamburg.docx` Universität Hamburg logo (B6)** — replace `word/media/image1.jpeg`
   with a neutral placeholder logo of the same dimensions, then re-export
   `parity/wild-hamburg-word.pdf` via the Word-automation pipeline. This is the only
   finding that changes pixels.

### Re-verify
After remediation, re-run `python3 scripts/audit-fixtures.py --md docs/FIXTURE-AUDIT-TABLE.md`
and confirm the BLOCKER count is 0.

---

## How to run the scanner

```
python3 scripts/audit-fixtures.py \
  --root /Users/alexpickett/Desktop/Projects/DocxInWeb \
  --md docs/FIXTURE-AUDIT-TABLE.md --json /tmp/audit.json
```

Automated: package/custom/app metadata, author trails (comments, tracked changes,
people.xml), external hyperlink targets, field `instrText`, drawing alt-text, bookmark
names, owner-lock usernames, PDF metadata. Media is listed with sha256 + dimensions for
manual visual review (`REVIEW` severity). Severity: `BLOCKER` = real PII/identifying;
`WARN` = metadata worth scrubbing; `REVIEW` = eyeball the image; `OK` = benign default.

## Rulings & status (2026-07-14)

- **Class A remediation: EXECUTED.** Scanner re-run reports **BLOCKER = 0**
  (rendering verified unchanged: patched PDFs and re-zipped docx measure at
  their exact accepted parity values).
- **Hamburg logo: ACCEPTED, no action.** Owner's ruling: publicly-sourced
  content in wild-corpus fixtures is fine (the source document is public);
  only personal PII (owner's name, private-file traces) must go. Class B is
  therefore empty.
- **History purge: QUEUED.** ~70 pre-scrub objects in pushed history still
  carry owner-name metadata / the Outlook review trail. Plan: git-filter-repo
  dropping historical `fixtures-staging/*.docx|pdf`, `~$*`, and `.codex-*`
  blobs (current scrubbed files re-committed), then force-push. Scheduled
  after in-flight agent branches merge — a rewrite would strand them.
