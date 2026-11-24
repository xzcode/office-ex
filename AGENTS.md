# Office EX development rules

- Read docs/HANDOFF.md, docs/LICENSING.md and docs/specs/README.md with relevant normative specs before coding. The v1.0 spec supersedes older roadmap details.
- User requires DOC/XLS/PPT (Office 97–2003) in addition to DOCX/XLSX/PPTX. All six are P0 for the first formal release; use independent CFB/legacy adapters and shared editor models. Never silently force conversion or rename an extension. M2-L/M3-L/M4-L are mandatory release gates.
- Use stable requirement IDs, preserve unknown source content safely, record actual evidence in features.json. Do not lower acceptance to make tests pass.
- The user chose independent development on 2026-11-24. Never fork, embed or wrap another Office suite/editor engine. Do not introduce ONLYOFFICE, LibreOffice or a ready-made office editing suite as the production engine. Independent readers may be used for tests only.
- Original code owns models, commands, undo/redo, pagination, layout, formula evaluation, slide scene and format adapters. Generic desktop/runtime, XML, ZIP, font and rendering libraries are permitted; prefer MIT/BSD/Apache-2.0, audit transitive dependencies and bundled assets.
- Desktop first; offline by default. No telemetry, external content, macros or cloud by default. Renderer must be sandboxed, context-isolated, with narrowly scoped IPC.
- Never silently discard unsupported imported parts or overwrite the original. Use capability reporting, safe Save As, atomic writes and format preservation.
- Report actual tests and limitations. No mock editor marketed as compatibility, no release without real platform builds and interoperability evidence. Do not run unnecessary tests.
