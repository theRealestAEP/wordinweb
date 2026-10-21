These synthetic DOCX files reproduce four content-obscuring layout regressions:

- `page-break-table.docx`: the paragraph after a table on a new page overlapped the last table row.
- `shared-frame.docx`: three paragraphs with the same positioned frame occupied the same line.
- `long-footnote.docx`: an 80-paragraph footnote extended above the page and covered body text.
- `exact-height-frame.docx`: a picture extended below a fixed-height frame and covered body text.

`e2e/layout-regressions.spec.ts` opens them through the built `wordinweb` React package in Chromium, in view-only and editable modes. It checks painted bounds and footnote completeness. The editable footnote case also checks layout after typing in the body.

Run from the repository root:

```sh
npx playwright test e2e/layout-regressions.spec.ts --project chromium
```
