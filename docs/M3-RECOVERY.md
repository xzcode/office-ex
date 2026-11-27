# M3 source recovery — batch 1 (2026-11-25)

Base commit: 785d3c7afaf16376b17c479a1624e0e3c29c4ff9 (M2 Writer). The previous M3 work was not present on a GitHub branch when this handoff began. The browser artifacts are separately preserved in recovery/m3-20261125, commit b4d84c9d7d04750d412ad6a5af38954c98d4a451. No claim of recovering the original TypeScript byte for byte.

## Implemented in this batch

- Restored and typed the SHT-03 lexer/parser/AST, column and cell reference syntax, bounded ranges, canonical formula serializer and support detection from the archived browser bundle.
- Reconstructed an independent, bounded Intl formatter cache and I18N-03 regional formula parsing/localization (German and French aliases, canonical fallback).
- Restored the public sheets API; malformed/incomplete token streams now produce a controlled OfficeError rather than a possible TypeError.
- Ten focused source-level tests cover precedence, references, ranges, literal kinds, locale roundtrip, unknown calls, malformed input, limits and cancellation.

## Verification

On the authorized Seoul Linux host, using the exact M2 GitHub archive and existing locked dependency installation (Node 24.21.0): npm run typecheck passed, npm run check:architecture passed (13 workspaces, 58 source files), and node --import tsx --test tests/sheets-formula-recovery.test.ts passed 10/10. This is source recovery verification only; it is not an independent reference Excel evaluation, a complete M3 verification, a release, or a native desktop result.

## Remaining

The parser identifies candidate supported function IDs, but the evaluator/dependency graph, spreadsheet commands and UI, XLSX/BIFF8 adapters, desktop IPC and full M3 tests are not yet restored in this branch. First recover and type the evaluator from the archived recalculation worker (without importing another office engine), then command/UI, format adapters and missing tests. Verify the independently reconstructed source and reproduce artifact-correlated evidence. The 196 tests and format/performance figures reported in the prior conversation are historical claims, not re-run here; do not mark SHT-03 or M3 verified based on parser tests. Keep M2/M2-L, M3/M3-L, all six Office formats and the 22-locale gate intact.

## Batch 2 — formula evaluation and incremental recalculation (2026-11-25)

Recovered a typed, runtime-independent workbook calculator from the archived recalculation worker. It maintains sparse cell/range indexes, static dependency and reverse-dependency indexes, parsed formula cache and derived calculation results. Supports arithmetic, comparison, concatenation, reference/range coercion and the twelve P0 function IDs. Uses topological evaluation for acyclic dependencies and lazy IF branches; cycle members are explicitly marked even when a function would otherwise swallow an error. Unsupported/stale values are not treated as valid operands; matching source caches may be retained for display only. Results are isolated by immutable workbook snapshot/document identity; step, depth, cancellation and deadline limits abort without committing partial state. Also fixed string tokens being interpreted as operators or references.

Evidence on authorized Seoul Linux host: typecheck passed; architecture passed (13 workspaces, 59 source files); focused formula parser and calculation suites passed 21/21 (10 previous + 11 new), including 1,499-formula chain, full-row/column sparse ranges, dependent-only recomputation, undo/redo, stale caches, cycle and abort behavior. These are targeted model-level tests, not a full suite or reference Excel validation. No native builds, XLS/XLSX interoperability or broad performance claim. The source recovery remains on m3-sheets-recovery-20261125.

Next: restore typed spreadsheet commands, formatting and selection, virtualized grid/workspace and clipboard; integrate calculator/worker through existing revision guard. Continue with file adapters and desktop acceptance afterward. Keep remaining SHT-03/04/05, I18N-03 and M3/M3-L gates open, as well as earlier M2/M2-L and all release gates.

## Batches 3–5 — editor, file profiles and desktop integration (2026-11-25)

Batch 3 restores editable sparse grid/workspace, command transactions, formula reference rewrites, merge, formatting, dimensions, range operations, clipboard, freeze/RTL, regional input and number/date formatting. Restored 22 candidate locale catalogs (288 keys each), including system-language matching and RTL placeholder isolation; translation review remains required. All recovered TypeScript is versioned source, not a production dependency on the archived bundle.

Batch 4 adds independent XLSX and BIFF8 XLS read/write profiles. XLSX supports the closed Transitional workbook/worksheet/shared-string/basic-style profile. BIFF8 supports a deliberately closed, reconstructible workbook with typed cells, string tables (record-boundary continuation only), supported RPN formula tokens/internal sheet references, numeric/error/boolean/string formula caches, basic fonts, opaque palette colors, uniform borders/fills, number formats, merges, row/column sizes, frozen panes, RTL and date system. XLS dimensions, palette, record/token/argument limits and unsupported whole-axis conversion are checked. BIFF8 formula source may be normalized and dimensions quantized to BIFF8 units. The semantic gate rejects unknown records/parts, macros, unsupported styles, external dependencies and source profiles that cannot be reconstructed safely. Ordinary third-party Excel files with additional metadata or features frequently fall outside the editable profile; they remain read-only structural inspection/preserved copy. No claim of arbitrary XLS/XLSX compatibility, PivotTable/chart/macro support or real Excel calculation equivalence.

Batch 5 connects the workspace to real desktop tabs and the revision-guarded calculation worker. New XLS/XLSX documents, editable-profile opening, semantic save/save-as, cross-format conversion to an independent tab and the existing main-owned sandbox verifier/atomic write/backup pipeline are wired. Source token and original-byte constraints remain in force. Unsupported features or unrepresentable target data cause a controlled failure rather than overwriting the original. Native smoke script updated for the virtual grid and native format save/conversion/reopen; real native three-platform runs still pending.

### Reproduced evidence (Seoul Linux, Node 24.21.0)

- Typecheck; architecture verification (13 workspaces, 67 source files); full node suite 198/198; desktop build all passed.
- Locked dependency audit passed (66 npm packages, 23 installed notices). No new production third-party Office engine; xlrd pinned as test-only.
- Independent openpyxl 3.1.5 and xlrd 2.0.2 successfully decoded fixture values/formulas (XLSX), formula caches, Unicode, both worksheet names, cross-sheet result, 1904 dates, formatting and merges/panes. This is not real Microsoft Excel verification of BIFF8 formula recalc.
- Fresh source-built headless Chromium 159.0.8154.6: 9/9 grid checks passed with sandbox on. See docs/verification/m3/browser-result.json (artifact hashes and samples) and sheets-grid.png. 100,000 populated cells initial grid median 703.9 ms; continuous scroll frame p95 50.0 ms (target <=33 ms NOT met); scroll-to-DOM p95 14.8 ms. Keep this performance gate open.
- No signed release, macOS/Windows native execution or formal M3/M3-L acceptance. All five recovery implementation batches have candidate code, but full-format, native/security, performance and reference-Office acceptance remain mandatory.
