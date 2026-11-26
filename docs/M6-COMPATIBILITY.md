# Office EX M6 — compatibility and editing experience

Development baseline: public dev.2 source `e0a827fa61a9643aa2ba4fc0b62edbd2aedacd42`, documentation head `761e75573062617325c5f526dbc7842ea7eab529`. Continue on `m5-distribution-20261125`. Published tags and master are unchanged; all formal M5 release gates remain in force.

## M6.1 batch A — ordinary DOCX table styles and native width semantics

The independent python-docx default Table Grid fixture previously fell back to read-only because the adapter required fixed, explicit grid/padding and rejected all default table styles. The adapter now recognizes the actual verified definitions, not only the names, of the default TableNormal and its TableGrid child. It checks the complete supported style payload: zero indentation, explicit four-sided default cell margins, uniform automatic-color grid border, Grid paragraph spacing and exact ancestry. Unknown conditional style rules, changed/custom definitions and unsupported default styles remain protected.

- Admit bounded rectangular ordinary tables with native auto width, autofit, and validated table-look flags. Retain these semantics and explicit native style references on save; no raw XML is accepted from renderer data. The editor projects the stored grid, not the Microsoft Word autofit algorithm.
- Project inherited cell padding, uniform border and Table Grid paragraph spacing; named paragraph and direct formatting retain precedence. Use a weak per-immutable-model index for Grid paragraph scope; numbering traversal carries the scope explicitly.
- Preserve explicit border and margin overrides even when they equal inherited values, and represent removal of an inherited Grid border with uniform explicit nil edges. Direct style/width constraints are validated on output. A no-edit save is byte exact; edited saves retain all ancillary package parts byte-for-byte and verify by reopening.
- Existing row/cell editing, row insertion, border change and undo remain available. Arbitrary table styles, conditional formatting, merged/nested tables, varying preferred cell widths, complex table alignment, automatic content-driven column sizing, row splitting and full reference Word typography are not established. Unsupported structures remain on the protected path. New source-dependent table metadata requires its verified style definitions; conversion that cannot retain those semantics must fail rather than silently flatten them.

### Local evidence

- Final-source `npm run check`: TypeScript, architecture (14 workspaces, 95 source files), 263/263 tests, zero failures/skips. Focused Writer table/style subset 12/12 on the initial candidate. Final desktop build and dependency audit (79 locked npm packages, 29 installed notice files) passed.
- `tests/office_external_compat.py` generates normal and TableGrid DOCX with python-docx plus merged-table negative fixture, runs source-backed edit, insertion, border override, byte-exact no-edit, all ancillary bytes, modified-style rejection and independent Python readback. Existing independent Excel/PowerPoint compatibility samples also pass.
- Final native sandboxed Linux focused smoke 3/3: real cell editing and row insertion, original native style/auto width/auto layout and all non-main parts retained, reopen editable with no renderer exceptions. Raw [report](verification/m6/native-word-table.json). Reproduce with `python tests/office_external_compat.py .work/m6-external`, build, then `node --import tsx scripts/smoke-m6-word-table.mjs .work/m6-external`. A local runtime setup issue (Electron distribution absent after npm ci, libgbm/libasound missing from system loader path) was resolved using the existing pinned Electron distribution and existing local test libraries; sandbox remains enabled.
- Exact new-source CI still required. No version bump or release in this batch. No production dependency additions.

## Next priorities

Continue M6.1 with merged cell model/geometry and more source styles, then M6.2 section-owned headers/footers and page numbering, M6.3 layout/pagination, M6.4–6 spreadsheet editing/formulas/charts, M6.7–8 presentation tables/groups/layout, M6.9 editor interaction consistency, M6.10 compatibility and a separately versioned preview. Do not call these later items complete or lower reference Office acceptance.


## M6.1 batch B — merged Word table cells

Baseline `85ee48862146267fdd256289d93f36a5aa261ac0` passed all four verification/native jobs in [CI 45774403536](https://github.com/xzcode/office-ex/actions/runs/45774403536) (Windows x64, macOS ARM64, Linux x64). This is evidence for batch A's exact source, not the new batch B source.

### Implementation

- The Writer model retains a rectangular grid of stable cell IDs. A merge anchor owns the original paragraphs and has bounded row/column spans; covered slots point to the anchor and contain no hidden content. Validation rejects out-of-bounds, overlapping, dangling or content-bearing covered slots.
- DOCX imports/exports horizontal `gridSpan` and vertical `vMerge` restart/continuation chains, including mixed rectangular merges. Explicit cell widths must match the summed grid widths; continuation paragraphs must be empty and unformatted. Unknown properties, mismatched chains and content-bearing continuations remain protected. No-edit saves remain byte exact; edited saves reopen and preserve all non-main package parts. The existing TableNormal/TableGrid native style and autofit semantics remain intact.
- Table layout produces one owning rectangle per merged cell, computes combined column width and satisfies multi-row height constraints. Connected vertical row groups paginate together, and hit testing resolves to owning paragraphs. Oversized merged groups fail explicitly; arbitrary row splitting and exact Word autofit remain unimplemented.
- Real Writer text-view tables use rowspan/colspan and omit covered inputs. Page view and PDF projection use the shared cell geometry. The Insert ribbon exposes explicit row and column dimensions from the current anchor, merge and split commands, with candidate labels in all 22 locales. Merge keeps every paragraph and its stable ID in row-major order; split keeps existing content in the top-left cell and creates empty paragraphs in other cells. Undo/redo restore the transaction. Column-width/image checks use the actual spanned width. Insert/delete row or column is temporarily disabled while the table contains merges; splitting merged cells first restores those operations. Nested tables, arbitrary conditional styles, content-bearing continuation cells and unsupported geometry remain protected.

### Evidence

- Final local `npm run check`: TypeScript, architecture (14 workspaces / 95 source files), 267/267 tests, no failures or skips. Targeted Writer table/style regressions 16/16; new merge tests cover domain validation, content preservation and undo, pagination/hit testing, oversized groups and native DOCX chains.
- Independent python-docx corpus now contains horizontal, vertical and rectangular merged tables and a negative continuation-content fixture. Source-backed editing, exact no-edit preservation, original span/style and ancillary bytes, and independent Python readback passed; existing Excel/PowerPoint samples passed. Desktop build and dependency audit passed (79 locked npm packages, 29 installed notices).
- Linux native sandboxed desktop 4/4: external merged-table rendering, cell editing and exact DOCX span preservation, ribbon merge retaining every paragraph, split/undo/redo/save and no renderer exceptions. Raw report: [native merged table](verification/m6/native-word-merged-table.json). Reproduce with `python tests/office_external_compat.py .work/m6-merged-corpus`, build, then `node --import tsx scripts/smoke-m6-merged-table.mjs .work/m6-merged-corpus`. No production dependency additions.
- New-source three-platform CI is still required. No new preview release, version bump, tag or master merge.

Next: extend M6.2 with bounded source-owned section header/footer stories, page numbering and independent DOCX samples; continue complex table styles and structural operations as separate M6.1 follow-ups. Do not describe the current subset as full Word compatibility. All formal M5 gates remain open.
