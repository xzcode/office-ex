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
