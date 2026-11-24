# M0-R2 actual validation — 2026-11-24

Local Node 22.16.0, npm 10.9.2, TypeScript 5.8.3, Python 3.13.5, python-docx 1.2.0, openpyxl 3.1.5, python-pptx 1.0.2.

- tsc -p tsconfig.json: PASS.
- npm test: PASS. Original TypeScript generated basic DOCX and XLSX XML parts; Python test harness ZIP-packaged them and independent readers verified Unicode, XML escapes, bold, newline/tab, multiple worksheets, AA cell, numeric values, text starting '=', and explicit SUM formula preservation.
- Four rejected inputs: illegal XML character, duplicate case-insensitive worksheet name, out-of-bounds cell, non-finite number: PASS.
- Formula execution and cached values are NOT verified; formula string was serialized only.
- npm registry DNS unavailable in local environment; Electron/ZIP/XML candidate libraries not installed. No app build, native launch, Office GUI round-trip, import, PPTX implementation or release.
- Local tested blobs: packages/ooxml/src/export.ts a025174611413cade1f4fed668b5d875f8a7c528; tests/export-fixtures.mjs 453b24862e7234d86abe503653857df6e62c5d0e; tests/smoke_export.py 475c4bf0a64dd45f15fd2fb140b54b7b7a895c39; package.json ffdc8bd70a4776ead3994dcc8e712acbbb5ccdd4; tsconfig.json 75260b545e3b04cae9a0c8f64b4ddd162b98cb8f.

The previous build records apply only to the superseded strategy and are available in Git history. No upstream office editor is used.
