# Handoff — M0-R2, 2026-11-24

Repository xzcode/office-ex, master. User chose plan 2: independent office suite with generic helper libraries only. The prior ONLYOFFICE/AGPL plan was abandoned; do not restore its upstream lock or scripts. Git history preserves the old plan. No upstream office editor code was ever imported.

Read docs/M0.md, docs/LICENSING.md, docs/M0_VERIFICATION.md and AGENTS.md. Original TypeScript DOCX/XLSX XML writer spike passed independent read tests; existing Python fixture generator retained. No production ZIP/XML adapter or UI yet.

Resume: verify network/dependency versions and licenses, lock exact runtime/dev versions, integrate bounded ZIP/XML adapters and build the secure Electron shell on macOS ARM64 and other targets. M1 then develops tabs, file operations, themes and shared document commands. Follow M2-M5; build genuine editors, not wrappers or screenshots. No release yet.
