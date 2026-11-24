# Office EX development rules

- Read docs/HANDOFF.md, docs/M0.md and docs/LICENSING.md before implementing.
- Scope: offline-first desktop suite; M0 is research, reproducible source preparation and compatibility baseline; do not claim a build passed without artifacts and logs.
- Do not rewrite the document engine or substitute a web-only mock for the editor.
- Maintain upstream commit IDs; never silently mix versions or pull moving upstream branches for a release.
- Keep independent MIT tooling separate from AGPL-covered derived code. Keep upstream notices; audit fonts, icons, dictionaries, CEF/Qt and other dependencies. Do not ship unreviewed assets or commercial extensions.
- Never enable cloud/AI/telemetry/automatic upstream updates by default. Do not send real user documents to external services. Treat document macros and embedded/external content as untrusted.
- Test only meaningful changes: round-trip semantic integrity, visual fidelity and native build/launch. Preserve original documents and test on copies.
- No Release until the M0 source build and licensing gates are satisfied, then follow M1-M5 roadmap.
