# Compatibility baseline

P0 (first editable release): DOCX/XLSX/PPTX creation, editing, Save As and reopening; preserve original file and compare output. No silent loss or corrupted OOXML. PDF reading, search, export and print; PDF editing comes later. Legacy DOC/XLS/PPT, ODF, RTF, CSV are secondary interoperability targets, not claimed supported yet. Password-protected files, VBA and external links must be treated safely and explicitly documented if not supported.

## Executable starter fixtures

tests/compat.py generates original synthetic DOCX, XLSX and PPTX. It verifies basic semantic invariants through independent Python parsers and ZIP integrity checks. These are smoke cases, not a complete fidelity suite.

1. Generate .work/fixtures and check it.
2. Create a separate .work/roundtrip directory. In the tested editor, open each fixture, save a copy under the same file name in that directory, close and reopen. Run python tests/compat.py check .work/roundtrip.
3. Record editor name/build, OS, font environment, file SHA256, exact save steps and failures. Test intentional edits with separate expected values. Never overwrite originals.
4. Manually compare pagination, layout, chart rendering and slides with Microsoft Office on a supported platform; LibreOffice may be used as an independent cross-reader, not a substitute for Microsoft Office verification.

## Expanded test matrix (not yet implemented)

| Format | P0 expanded coverage | P1 |
| --- | --- | --- |
| DOCX | mixed Chinese/English, tables, page breaks, images, headings, styles, headers/footers | tracked changes, footnotes, fields, section breaks, embedded objects |
| XLSX | data types, formulas and references, multiple sheets, merged cells, chart, formats | pivot tables, conditional formats, named ranges, validation, external refs |
| PPTX | text, shapes, images, slide size and ordering, theme | masters, animations, transitions, media |
| PDF | open, search, print, zoom, export | annotations, forms, editing, signatures |

Each critical case must pass a no-edit save cycle and intended-edit save cycle. Compare structural semantics, calculated values, visual appearance and reopen results; unknown unsupported elements must not disappear silently. Security cases (ZIP expansion limits, untrusted macros, external resource loading, malformed input) and crash recovery must be separately covered. Test data must be synthetic or redistributable. No user documents in CI.

Pass requires recorded results, not merely a successfully generated ZIP file.
