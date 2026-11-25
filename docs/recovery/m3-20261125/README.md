# M3 recovery checkpoint — 2026-11-25

This branch is an evidence/recovery snapshot, **not** the M3 source branch or an accepted release.
Base: `785d3c7afaf16376b17c479a1624e0e3c29c4ff9` (M2 Writer, not merged into master).
The prior Office-EX-01 conversation reported 196 tests and additional format/performance checks; those full test sources and raw format/performance results were not recovered in this handoff. Do not report them as re-run here.

## Preserved artifacts

The archive was copied from `/app/agent/office-ex-m3-20261125` on the authorized Seoul Linux host, without changing the original files. It contains the browser bundle (entry.js), recalculation worker, stylesheet, headless Playwright smoke script, raw result JSON and screenshot. The files are test/build outputs, not the original TypeScript source. Neither JS file contains an embedded source map. The archive excludes the large Node executable, node_modules, and credentials. SHA-256 of archive:
`01db3aa3f7e438bcc4b7bfb2e9458b342565baebf8639e96c2e081ce49aea73f`.

The raw result records nine browser checks, including formula and undo/redo, TSV, merge, formatting, cross-sheet formula, freeze and RTL, regional editing, rejecting invalid mutations, cached BIFF preview, and 100k sparse-cell viewport. Browser-only result: first grid median 628.2ms, max 744.2ms; continuous rAF intervals p95 50ms, max 83.4ms (33ms target **not met**); scroll-to-DOM p95 23.2ms. Hashes of entry.js, worker and stylesheet matched the raw result before archiving. This is not native Electron, reference Office, or file-format acceptance.

## Resume sequence

1. Recover the original M3 working tree, if accessible, and compare with M2 base. Otherwise restore typed modules using the bundle's source-path markers and original M2 sources, and reconstruct missing format adapters, host IPC and regression tests separately. Do not pretend that de-bundled JS is the original TS or that the UI bundle contains XLS/XLSX adapters.
2. Keep all M2/M2-L and M3/M3-L acceptance gates. Restore source and targeted tests on a dedicated development branch; reproduce source-correlated build, verify format round-trips, conversion bounds, unsupported content and three-platform native behavior.
3. Optimize the actual continuous scroll frame metric before declaring the target met. Preserve this baseline.
4. Continue M4 Slides/PPT/PPTX only after the M3 code is durably stored and remaining acceptance limitations are documented.

M2 branch `m2-writer-20261124` and `master` remain unchanged. This archive is not a production dependency. Do not merge the binary archive into production solely to claim M3 was implemented.
