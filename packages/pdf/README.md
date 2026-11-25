# PDF boundary and local rendering

PDF.js 6.5.438 (Mozilla, Apache-2.0) supplies PDF parsing, page appearance rendering and glyph/text extraction. Office EX owns the read-only workspace, source-byte lifetime, bounded operations, navigation, selection, search, printing, output policy and Office document projection. No Office converter or other Office engine is used.

## Safety and resource policy

The reader receives bytes, never a PDF-controlled URL. Its dedicated module Worker is instantiated explicitly, so it cannot fall back to parsing in the application renderer or create a cross-origin blob wrapper. Built-in CMaps and fonts come only from the fixed local `pdf-assets/` directory. No scripting manager, QuickJS sandbox, attachment layer, link service, form layer, viewer plugin or XFA UI is instantiated. Annotation appearances may be painted on the canvas, but no action controls exist. `enableXfa` and WebAssembly are disabled; unsupported decoding must fail rather than load remote code. Password-protected and empty-password encrypted PDFs are rejected. The main process does not parse PDFs.

Limits: 100 MiB input, 500 pages, 40 million pixels for a rendered page, 100,000 text items / 2 million extracted characters per page, 8 million retained/searchable characters, 10,000 matches and 512 query characters. Search is literal and case-sensitive. A 30-second work watchdog survives caller cancellation until the underlying operation settles; text streams are cancelled and a stuck worker is terminated. Tab detachment cancels the view while preserving navigation state; closing a tab destroys its worker. Print appearances have a separate 100-million-pixel total budget.

Source PDF bytes remain immutable. PDF is not an Office model, so the derived selection layer uses PDF.js's glyph normalization and PDF `ActualText` replacements. It does not modify either the source PDF or an Office document string. Scanned pages with no text show a clear no-searchable-text message; OCR is not supplied.

## ActualText adapter

The upstream 6.5.438 worker omits marked-content `ActualText` values from text extraction. Chromium-generated PDFs use those values for ligatures, pre-base vowels, shared CJK glyphs and bidirectional text. Without them, the correct page appearance can copy different characters or NULs even though independent PDF readers recover the logical text.

`prepare-worker.mjs` applies one guarded change to the **copied distribution worker**, leaving `node_modules` and the upstream lock integrity untouched. It verifies the exact upstream SHA-256 `43710009ea26bfbcb615524946e46e3aaf8ad927514f5e465e8ff6d18044ffe9` and a unique source anchor. It adds a decoded `actualText` property to the existing marked-content result, resolving both inline and named property dictionaries. The original Apache-2.0 copyright header remains in the output. A changed upstream worker fails the build until this adapter is reviewed again.

Office EX's original `PdfActualTextBuilder` validates the resulting data, limits nesting to 128 and records to one million, restores logical `ActualText` replacements, and handles `ReversedChars` groups while retaining the glyph bounds for selection. The PDF canvas pipeline is unchanged. Malformed, unbalanced, unmeasurable or oversized groups produce a controlled error instead of invented text. Dedicated tests cover the Unicode defect and those limits.

## Office PDF output

An isolated sandboxed print window recovers a validated draft through the existing format worker. Writer uses the final native paragraph shaper and pagination engine; Slides uses `buildSlideScene` and `createSlideSceneElement`, the same revision-bound scene as editing and presentation. Native text remains text and shapes remain vectors when Chromium prints the document. CSS named pages preserve actual per-page dimensions; no editor controls are included. Images must decode successfully. Sheets export is blocked until print areas and pagination rules exist.

The font bundle and export allowlist are documented in `assets/fonts/README.md`. Arbitrary system fonts are not automatically approved for embedding. Existing PDFs are printed as safe page appearances; source files are never rewritten by printing or by Office PDF export.

## Sources

- [Mozilla release 6.5.438](https://github.com/mozilla/pdf.js/releases/tag/v6.5.438), 2026-11-01.
- [PDF.js API](https://mozilla.github.io/pdf.js/api/draft/module-pdfjsLib.html).
- [Mozilla security advisory GHSA-hq66-cqwq-w95j](https://github.com/mozilla/pdf.js/security/advisories/GHSA-hq66-cqwq-w95j), fixed in 6.2.108. The pinned version includes this fix; the application additionally never enables viewer scripting or XFA and retains its strict local CSP.
- [Electron webContents.printToPDF](https://www.electronjs.org/docs/latest/api/web-contents#contentsprinttopdfoptions).

Current verification is recorded separately in the M4 evidence. Reader-level tests are not a claim of universal PDF fidelity, OCR, editing, signatures or support for arbitrary encrypted documents.
