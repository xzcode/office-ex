# Office EX

**English** | [简体中文](README.zh-CN.md)

An independent, open-source, cross-platform, offline-first office suite for word processing, spreadsheets, and presentations.

Office EX is written in TypeScript and runs as a desktop application. Its original document models, editing commands, undo/redo history, layout, formulas, and Office format adapters are developed within this project. Electron and general-purpose open-source libraries provide the runtime, compression, XML, and PDF infrastructure. The project does not embed Microsoft Office, LibreOffice, or ONLYOFFICE as its editing engine.

## Download

The latest published development preview is **[v0.1.0-dev.3](https://github.com/xzcode/office-ex/releases/tag/v0.1.0-dev.3)**.

| Platform | Installer |
| --- | --- |
| Windows x64 | [Download EXE](https://github.com/xzcode/office-ex/releases/download/v0.1.0-dev.3/Office-EX-0.1.0-dev.3-win32-x64-setup.exe) |
| macOS Apple Silicon | [Download DMG](https://github.com/xzcode/office-ex/releases/download/v0.1.0-dev.3/Office-EX-0.1.0-dev.3-darwin-arm64.dmg) |
| Linux x64 | [Download DEB](https://github.com/xzcode/office-ex/releases/download/v0.1.0-dev.3/Office-EX-0.1.0-dev.3-linux-x64.deb) |

See [all releases](https://github.com/xzcode/office-ex/releases), [SHA-256 checksums](https://github.com/xzcode/office-ex/releases/download/v0.1.0-dev.3/SHA256SUMS), and the [release acceptance record](https://github.com/xzcode/office-ex/actions/runs/45803738626).

This is a development preview, not a production release. The Windows installer is unsigned; the macOS application uses ad-hoc signing and is not notarized. Keep system security protections enabled and back up important documents.

## Features

The published dev.3 preview implements a bounded, tested subset of the following features.

| Module | Capabilities |
| --- | --- |
| Writer | DOC / DOCX editing, text styles, tables and merged cells, section layout, headers and footers, automatic page numbers, paragraph pagination |
| Sheets | XLS / XLSX editing, virtualized cell grid, direct formatting, 48 formula functions, conditional formatting, basic column, line and pie charts |
| Slides | PPT / PPTX editing, text and shapes, images, groups, ordinary native tables, alignment and distribution, presentation mode |
| PDF | PDF reading and supported document export |
| Desktop | Multiple documents, undo/redo, native save and save as, previous-version backups, crash recovery, themes, 22 UI locales |

Office EX preserves original native package parts where supported. Documents outside its editable profile are protected rather than silently flattened or stripped of unsupported content. Old binary formats, advanced layout, formulas, charts, tables, printing, accessibility and reference Office fidelity are not yet fully compatible. The internal .oex draft format is not a Microsoft Office format.

## Development status

- **M0–M6:** development-preview delivery completed. The dev.3 candidate passed 293 core tests and Windows, macOS and Linux native package/installer verification.
- **M7:** ongoing deeper compatibility, performance, stability and product refinement. Post-dev.3 features belong to development builds, not the currently published installer.

The default `master` branch currently retains the earlier implementation baseline. For current source and documentation, use the [active development branch](https://github.com/xzcode/office-ex/tree/m5-distribution-20261125). For reproducible released source, use the [dev.3 tag](https://github.com/xzcode/office-ex/tree/v0.1.0-dev.3). The README describes the published preview and identifies ongoing work separately.

## Build from source

Use Node.js 24.x (at least 24.13.0). The initial installation requires npm and Electron downloads; normal document editing runs locally.

```sh
git clone -b m5-distribution-20261125 https://github.com/xzcode/office-ex.git
cd office-ex
npm ci --ignore-scripts
npm run setup:desktop
npm run start
```

Linux requires a graphical session, a working Chromium sandbox, and suitable fonts. Do not disable the sandbox.

## Development and verification

Python is used for independent compatibility fixtures and readback tests, not the production editing engine.

```sh
python -m pip install -r tests/requirements.txt
npm run check
npm run test:interop
npm run audit:licenses
npm run build
npm run smoke:desktop
```

See the active branch for the current [development handoff](https://github.com/xzcode/office-ex/blob/m5-distribution-20261125/docs/HANDOFF.md), [M6 compatibility profiles](https://github.com/xzcode/office-ex/blob/m5-distribution-20261125/docs/M6-COMPATIBILITY.md), and [M7 stability work](https://github.com/xzcode/office-ex/blob/m5-distribution-20261125/docs/M7-STABILITY.md).

## License

Original project code is licensed under the [MIT License](LICENSE). Third-party components retain their own licenses; see [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). Office EX is an independent project and is not affiliated with Microsoft.
