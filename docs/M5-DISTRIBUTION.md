# M5 — Distribution and release readiness

Scope: QA-03–07, six native Office formats, M2-L/M3-L/M4-L, installation, compatibility, safety, performance and language release gates. The v1.1 normative specs remain authoritative.

## Batch 1: native development-preview packaging

M5 development branch starts from M4 commit `3b49ace8560839354700f3c97aba01de439f0eef`, which completed all four jobs in [Actions 45595130602](https://github.com/xzcode/office-ex/actions/runs/45595130602). This is the M4 bounded implementation evidence, not formal compatibility certification.

- `scripts/package-preview.mjs`: packages the actual pinned Electron distribution plus prebuilt self-contained application, offline PDF resources, hash-verified OFL fonts and third-party notices. Linux x64 tar.gz, Windows x64 ZIP, macOS ARM64 ZIP; native installer option generates DEB, Inno Setup EXE or DMG respectively on the matching native host. No new runtime dependency, network updater, release credentials or publishing action.
- The preview version is retained, macOS signing is ad-hoc only, Windows unsigned, Linux portable archive has a non-setuid sandbox helper and requires a supported Chromium sandbox configuration. The installer workflow uses the native OS build tools; no installation or global file associations are claimed on the developer host.
- `scripts/verify-preview.mjs`: verifies SHA256SUMS, exact per-file content manifest and internal symlinks, extracts the trusted locally produced archive, starts the **packaged executable** with a temporary profile, confirms sandbox/context isolation/no Node/no DevTools, exercises creation/native save/independent semantic parse/reopen for DOC/DOCX/XLS/XLSX/PPT/PPTX, blocks external fetch, and checks orderly quit. Native dialog choices are automated; this is not independent Microsoft Office GUI testing.
- `.github/workflows/m5.yml`: one locked core/interop/license check and three native packaging jobs, archive launch verification and development-preview artifact uploads only on success. Native installers require the platform tools; the new cross-platform workflow remains to be verified. No automatic GitHub Release.
- `scripts/release-readiness.mjs` plus `docs/verification/m5/release-gates.json`: explicit non-bypassable-by-preview-version check of tracked mandatory feature statuses and nine acceptance gates. The report is generated locally; `--require-formal` exits nonzero when not ready.
- First packaged test exposed a genuine blank DOC round-trip bug: empty text run paragraph-mark styling was dropped by the canonical binary writer. Preserve the effective character style of an empty run for an empty paragraph, keeping the original binary structural equality guard. Added a targeted blank/styled-empty DOC regression; Writer focused tests 6/6, TypeScript and build passed.

## Open formal release gates

The preview archive and installer workflows do not authorize formal release. Reference Office GUI and legacy acceptance, three-platform installer installation/uninstallation and trust, full QA-05 budgets, all 22 locales' human review and native IME/printer verification remain open or blocked. See the gate ledger and normative `docs/specs/08-quality.md`, `09-i18n.md`. Next M5 batch prioritizes six-format compatibility corpus, Sheet performance and reproducible installation evidence. No master merge or release.

## Local packaged executable evidence

Seoul Linux x64, sandboxed headless Electron 45.4.1 / Chromium 156.0.8078.100. Final locally built archive: 142,396,197 bytes; SHA256 `104ee956360b16d8a66f5e86be4f03f28a723b6c05f83a5131a7150774193947`; 327 exact manifest entries. All 10 packaged checks passed after the empty DOC fix, including empty DOC native save and reopening, other five native formats, offline renderer boundary and normal shutdown. Measured one instrumented launch to renderer ready: 1312.5 ms, not a five-run cold-start or RSS certification. The local artifact was built from the M4 base plus this uncommitted M5 working tree; `sourceClean:false` is retained, and the forthcoming source-correlated CI must rebuild it. The full package stays in the authorized host's out/m5-preview (not committed or released). See `docs/verification/m5/linux-packaged-preview.json`. Windows and macOS native installer builds, Linux DEB generation and actual installation remain pending; local Linux test is archive extraction, not DEB installation.

The first smoke had an Electron web preference metadata mismatch and premature cleanup; the corrected check verifies that DevTools cannot actually be opened and waits for process exit before removing test files. The first empty DOC save then exposed the canonical round-trip failure, which was fixed and verified from the rebuilt packaged executable. The formal readiness command rejects this development version and all still-open gates.

## Batch 2 — native installer acceptance and bounded Sheets viewport

Previous M5 source `7db8704fe416c87f608b83f81702616137fb9011` [Actions run 45597866933](https://github.com/xzcode/office-ex/actions/runs/45597866933): core/check/independent interop/licenses passed; Windows x64 and Ubuntu x64 completed the installer build and all extracted-app checks successfully. Their artifacts were uploaded by the workflow. macOS ARM64 built and verified its DMG checksum and exact archive manifest, then failed the extracted-app gate because Electron reported `app.isPackaged=false`. The original app executable still used Electron's default name. Renamed the main macOS executable to Office EX, updated CFBundleExecutable, and retained the strict packaged check and ad-hoc signing. This change requires native macOS CI confirmation; the failure is not accepted as a pass.

Added `scripts/verify-preview-install.mjs`, an explicit opt-in test limited to disposable Actions runners and a clean source-correlated manifest. It checks installer hashes, installs a fresh Linux DEB / Windows per-user EXE / macOS DMG copy, verifies packaged file hashes, launches actual installed executable with the sandbox, opens a native worksheet, exits normally and uninstalls/removes the test installation. Linux DEB installs the root-owned 4755 Chromium sandbox helper; portable archives retain non-setuid mode. Mac DMG includes Applications shortcut; Windows install avoids closing or restarting other applications. No ordinary workstation install is attempted by the verification script. Guard and JS syntax checks passed locally; the actual native installer tests require the new CI run.

Sheets reuses outgoing cell and header nodes, prioritizes reuse within columns, retains text nodes, caches immutable style CSS and skips unchanged accessibility/status/input writes. The scroll plane still provides full logical coordinates and dimensions; the actual drawing layer is clipped and contained to the viewport. Merged anchors, frozen cells, address/formula input, styles, opaque formulas and regional formatting remain model-driven. New browser regression verifies DOM identity reuse, no stale merged/style/accessibility attributes and viewport-height bound. All 12 browser checks passed on the final implementation.

Same Seoul host, same headless Chromium 159.0.8154.6, same 100k-cell workload (milliseconds):

| Candidate | Constructed-model first grid median | Continuous scroll frame p95 | Scroll-to-DOM p95 |
|---|---:|---:|---:|
| Before | 672.7 | 50.0 | 14.7 |
| Initial pooling | 626.4 | 33.4 | 18.2 |
| Column reuse intermediate | 614.6 | 50.1 | 12.2 |
| Final bounded viewport | 690.4 | 50.0 | 11.5 |

The final scroll-to-DOM latency improved, but continuous scroll-frame p95 remains 50 ms (target <=33 ms). Intermediate results vary; do not select the best run as a certified pass. Constructed-model first-grid timings are not disk-open measurements or complete QA-05 performance evidence. Keep QA-05 and M3/M3-L release gates open. Local final TypeScript/architecture/build/dependency audit passed; 250/250 core and 31/31 real sandboxed Linux desktop checks passed. Browser/native raw evidence archived in docs/verification/m5. New-source CI, actual installation/uninstallation, publisher trust, full reference Office and linguistic acceptance remain open. No release/master merge.
