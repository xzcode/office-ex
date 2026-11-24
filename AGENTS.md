# Office EX development rules

- Read docs/HANDOFF.md, docs/M0.md, docs/LICENSING.md.
- The user chose independent development on 2026-11-24. Never fork, embed or wrap another Office suite/editor engine. Do not introduce ONLYOFFICE, LibreOffice or a ready-made office editing suite as the production engine. Independent readers may be used for tests only.
- Original code owns models, commands, undo/redo, pagination, layout, formula evaluation, slide scene and format adapters. Generic desktop/runtime, XML, ZIP, font and rendering libraries are permitted; prefer MIT/BSD/Apache-2.0, audit transitive dependencies and bundled assets.
- Desktop first; offline by default. No telemetry, external content, macros or cloud by default. Renderer must be sandboxed, context-isolated, with narrowly scoped IPC.
- Never silently discard unsupported imported parts or overwrite the original. Use capability reporting, safe Save As, atomic writes and format preservation.
- Report actual tests and limitations. No mock editor marketed as compatibility, no release without real platform builds and interoperability evidence. Do not run unnecessary tests.
