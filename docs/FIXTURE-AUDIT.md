# Fixture privacy audit

Scan every test fixture (`apps/demo/public/fixtures/`, `fixtures-staging/`,
`parity/`) and the Word reference PDFs for real or identifying information before
any fixtures are published.

Re-run the scanner:

```
python3 scripts/audit-fixtures.py \
  --root /Users/alexpickett/Desktop/Projects/DocxInWeb \
  --md /tmp/fixture-audit-table.md --json /tmp/audit.json
```

It covers package / custom / app metadata, comment and tracked-change author
trails, external hyperlink targets, field `instrText`, drawing alt-text, bookmark
names, owner-lock usernames, and PDF metadata. Media is listed with sha256 +
dimensions for manual review. Severity: `BLOCKER` = real PII, `WARN` = provenance
metadata, `REVIEW` = eyeball the image, `OK` = benign default.

## Rulings (2026-07-14)

- **Class A metadata scrubs: executed.** Scanner re-run reports BLOCKER = 0;
  rendering verified unchanged (patched PDFs and re-zipped docx measure at their
  exact accepted parity values).
- **Hamburg logo: accepted, no action.** Publicly-sourced content in wild-corpus
  fixtures is fine (the source document is public); only personal PII must go.
- **History purge: queued.** ~70 pre-scrub objects in pushed history still carry
  owner-name metadata and the Outlook review trail. Plan: `git-filter-repo` to
  drop historical `fixtures-staging/*.docx|pdf`, `~$*`, and `.codex-*` blobs, then
  force-push. Scheduled after in-flight agent branches merge.
- **Anonymized twins replace name-bearing fixtures (2026-07-14).** `pleading`
  (a FABRICATED deposition notice that referenced a real case, real attorneys
  and a real firm) and `pickett`/`alexpickett` (cover-letter template carrying
  the owner's real name in body text) are deleted; `pleading-anon` and
  `coverletter-anon` take their suite slots with fresh Word reference PDFs
  (pleading-anon 0.00% on all 7 pages, coverletter-anon 0.31%). Their old blobs
  join the history-purge list above: `apps/demo/public/fixtures/pleading.docx`,
  `pickett.docx`, `alexpickett.docx`, plus pre-scrub `coverletter-anon.docx` /
  `forsale.docx` (their `settings.xml.rels` embedded a `file:///Users/<owner>/…`
  attachedTemplate path — a leak class the scanner now needs to cover).
