# Dependency review — M1

Review date: **2026-11-24**. The application remains independently developed under its original MIT license. This inventory covers the exact npm lock and the installed package notice files; it distinguishes the application's JavaScript libraries, native Electron distribution, and development tools.

## Direct dependencies

| Component | Pinned version | Declared license | Purpose | Distribution |
|---|---|---|---|---|
| [Electron](https://github.com/electron/electron) | 45.4.1 | MIT for Electron; bundled components have their own notices | Desktop runtime, OS dialogs, sandboxed Chromium/Node processes | Native application host |
| [fflate](https://github.com/101arrowz/fflate) | 0.8.3 | MIT | General DEFLATE codec; Office EX implements ZIP limits, CRC/path validation and OPC semantics | Bundled in format worker |
| [@xmldom/xmldom](https://github.com/xmldom/xmldom) | 0.9.12 | MIT | General XML DOM; Office EX enforces XML resource limits and rejects unsafe constructs | Bundled in format worker |
| [TypeScript](https://github.com/microsoft/TypeScript) | 7.0.2 | Apache-2.0 | Strict type checking, including a separate platform-free domain check | Development only |
| [tsx](https://github.com/privatenumber/tsx) | 4.23.15 | MIT | TypeScript loader for Node's test runner | Development only |
| [esbuild](https://github.com/evanw/esbuild) | 0.28.2 | MIT | Bundle desktop, preload, renderer and worker outputs | Development only |
| [playwright-core](https://github.com/microsoft/playwright) | 1.65.1 | Apache-2.0 | Native Electron/browser smoke automation | Development only |
| [@types/node](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/node) | 24.19.2 | MIT | Main-process and development type declarations | Development only |

The exact licenses and copyright statements are copied verbatim from the installed packages into [`notices/`](../notices/). Apache NOTICE files are retained where the package supplies them. Source links above identify the upstream projects; the immutable npm tarball URLs and SHA-512 integrity values in [`DEPENDENCIES.lock.json`](DEPENDENCIES.lock.json) identify the installed versions.

## Transitive dependencies

`DEPENDENCIES.lock.json` lists every third-party package in `package-lock.json`, including optional binaries for platforms other than the current test machine. Entries identify version, license, npm origin, integrity, dependency edges, platform restrictions and use. An optional package listed in the lock is not evidence that its platform was executed or tested.

The two application JavaScript libraries, fflate and xmldom, have no additional npm runtime dependencies in this lock. Electron's npm downloader and its helper packages are installation/development tooling; the native Electron binary is tracked separately. The closed production dependency allowlist prevents an Office engine or converter from being added unnoticed.

Most locked npm packages declare MIT, Apache-2.0 or ISC. One reviewed tooling exception is **graceful-fs 4.2.12**, an indirect dependency of Electron's downloader, under **BlueOak-1.0.0**. Its exact license permits use and modification with notice preservation; the installed license text is retained. The audit allows this particular version only in the development/installation graph. It does not broadly approve new non-default licenses. Sources: [license steward](https://blueoakcouncil.org/license/1.0.0), [SPDX identifier and text](https://spdx.org/licenses/BlueOak-1.0.0.html).

The script checks installed package names, versions and declared licenses against the lock; verifies reviewed notice hashes; rejects unreviewed license identifiers, unpinned direct dependencies, unexpected production dependencies and Office engine/converter packages. This is a reproducible inventory and policy check. It does not infer that every file inside a native binary has the package's root license.

## Native notices and build behavior

Run `node scripts/dependency-inventory.mjs` to verify the checked-in review. After an intentional dependency change and review, run `node scripts/dependency-inventory.mjs --write` to regenerate the inventory and exact notice copies. Run with `--stage`, or call the exported `auditDependencies({stageDir: 'dist/licenses'})` during the build, to stage redistributable notices.

Staging requires a native Electron distribution matching the pinned npm version and copies its **LICENSE** and **LICENSES.chromium.html** without modification. The staged manifest records version, platform, architecture, file sizes and SHA-256 hashes. The large Chromium notice file stays in build output; it is not committed here. Every platform package must retain both the staged notices and the native distribution's own notices. The source release retains the root MIT license, this inventory and exact package notices.

Bundled Chromium/Node components retain their individual terms. Their notice coverage comes from the matching upstream distribution, not from the npm root-license allowlist. Any source changes, alternative native distribution, added media codec or packaged third-party asset require a new review. Electron's official [repository](https://github.com/electron/electron) and [third-party attribution build rules](https://github.com/electron/electron/blob/v45.4.1/BUILD.gn) are the primary origins.

## Security and asset status

The M1 dependency installation was checked with npm's advisory audit on **2026-11-24**; the run reported **0 known vulnerabilities** across runtime and development dependencies. This is a dated advisory result, not a guarantee of safety or a substitute for Office EX's parser, IPC and file-safety tests. Refresh the audit when dependencies change and before release, and retain the machine-readable output with the verification evidence.

M1 bundles **no external font, icon collection, image, template or user document**. Canvas measurements use available system fonts; missing-font diagnostics are part of the rendering interface. No installed-font availability, typography fidelity or font redistribution permission is claimed by this dependency review.

Python packages in [`tests/requirements.txt`](../tests/requirements.txt) are independent fixture/interop tools used only in tests. They are excluded from application bundles. CFB, shared models, command history, format recognition, save planning, file authority, geometry and document semantics remain original Office EX code. Native Office compatibility, historical Office application coverage and typography equivalence have separate later-stage acceptance gates.
