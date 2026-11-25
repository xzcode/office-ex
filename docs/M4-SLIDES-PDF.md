# M4 Slides and PDF — recovered candidate

Date: 2026-11-25. Source recovery base `8273c89d3e16d04cbcc9ab613b085040016e9460` (`m3-sheets-recovery-20261125`). Preserved original M4 archive on the Seoul host at `/app/agent/office-ex-m4-20261125/m4-source.tar.gz`, SHA-256 `1e379e4448820864032c9119c1767a9af965a0d1722bb1313e25d65a24af9237`. The new branch keeps master and M2/M3 history intact. Source snapshot excludes runtime caches, downloads, output and secrets.

## Recovered implementation

- SLD-01/02: ordered editable slides, text and basic vector/image shapes; shared EMU scene, geometry, rotation, fill/stroke, layout, text, thumbnails and presentation; session transactions and undo/redo. Review `packages/slides`, `packages/render/src/slides-scene.ts` and `packages/ui/src/slides-workspace.ts`.
- SLD-03/LEG-04: bounded supported PPTX and PowerPoint 97–2003 PPT semantic profiles, source preservation or refusal for unsupported structures, CFB streams/persist maps and independent conversion tabs. Original-format save uses the isolated verifier and existing atomic file services. See `packages/ooxml/src/slides-pptx.ts`, `packages/format-service/src/legacy-ppt.ts`, `apps/desktop/src/docx-verifier.ts`.
- PDF-01/02: local PDF.js worker, bounded read-only navigation/zoom/text selection/search, disabled active content, safe appearance printing; Writer/Slides PDF export from final layout/shared scene in a sandboxed print window. Sheets export explicitly blocked pending print area/pagination. See `packages/pdf/README.md`.
- All 22 candidate UI locales have matching message keys; formal terminology, native RTL/CJK shaping and real IME review remain required.

## Recovery fixes

The transferred source lacked its eight approved Noto font resources and their license manifest. Restored the assets from pinned official upstream commits, regenerated actual Unicode cmap coverage, checked font OS/2 embedding flags and license text, committed a reproducible SHA-256/size-verified build setup command (`npm run setup:fonts`) and notices. Font binaries are fetched at build setup, not from an external service at runtime. Recovered the matching Electron distribution and two missing Linux runtime libraries locally; no sandbox bypass or system-wide package install. Added independent PPTX/PPT fixture inspection with python-pptx and olefile; new M4-only CI gate does not rewrite previous milestone evidence.

## Observed evidence

Seoul OpenCloudOS 9.6 x64, Node 24.21.0 / Electron 45.4.1. TypeScript and architecture checks passed. Full core regression 236/236 (including 37 M4 Slides/PDF targeted tests); independent 5-file PPT/PPTX fixture inspection passed. Build and dependency audit passed (79 locked packages, 29 installed notice files). All 22 candidate locale catalogs validated. Independently inspected native Slides PDF with Poppler: 3 pages, 960×540 pt, Unicode text including Chinese, Arabic, Hebrew, Hindi and Thai. Real Electron boot retained sandbox/context isolation; forged capabilities rejected; editing text/shapes/image/pages and presentation passed; PPTX native save, isolated PDF export, and conversion to an independently parsed PPT passed. See the evidence record for subsequent native smoke result. These checks do not constitute reference PowerPoint rendering or physical printer/IME acceptance.

## Remaining gates

Keep M4 and M4-L open: complete end-to-end native PPT reopen/edit/save and reverse conversion, PDF reader/print native path, browser timing, broader unsupported-source/real-world fixture coverage, independent reference PowerPoint playback and all three OS native checks. Confirm final Writer/Slides PDF output dimensions, text and fonts; PDF OCR/encrypted input not supported. The existing M3 performance/compatibility and M2-L release gates also remain open. No merge to master and no Release until acceptance is met.
