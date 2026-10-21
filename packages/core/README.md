# @wordinweb/core

Internal DOCX parsing, layout, rendering, and editing workspace for WordInWeb.

This workspace is bundled into the public [`wordinweb`](https://www.npmjs.com/package/wordinweb) package and is not published separately.

The `@wordinweb/core/docx-only` entry exports `DocxDocument`, `layoutDocumentAsync`,
and `renderToDom` for ZIP-based DOCX callers. It excludes the binary Word converters.
The public package exposes the same entry as `wordinweb/docx-only`.

Documentation and source: https://github.com/theRealestAEP/wordinweb

Licensed under the MIT License. See `LICENSE`.
