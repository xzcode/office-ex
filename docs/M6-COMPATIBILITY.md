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
