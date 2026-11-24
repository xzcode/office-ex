# Third-party notices

Office EX original source is licensed under the root [MIT license](LICENSE). Third-party software retains its own copyright and license terms.

The pinned package inventory, origins, integrity hashes and license identifiers are recorded in [docs/DEPENDENCIES.lock.json](docs/DEPENDENCIES.lock.json). Exact, unmodified license and notice files copied from the installed npm packages are in [notices/](notices/). See [docs/DEPENDENCIES.md](docs/DEPENDENCIES.md) for their purpose and the scope of the review.

Production format infrastructure currently uses **fflate 0.8.3** and **@xmldom/xmldom 0.9.12**. **Electron 45.4.1** provides the desktop runtime. These are general compression, XML and desktop components. No third-party Office editor or document conversion engine is a production dependency.

## Native Electron distribution

Electron's root MIT license does not replace the licenses for Chromium, Node.js or their bundled components. Every native build must retain the exact `LICENSE` and `LICENSES.chromium.html` files from the matching Electron distribution. The build copies them to `dist/licenses/electron-distribution/`, alongside a version/platform/hash manifest. A release must carry this directory and the native distribution's own notices. Missing files block notice staging; the 5.6 MB Chromium notice file is not duplicated in this source repository.

The JavaScript bundler emits external legal-comment files. Those files must remain next to their associated output and accompany the complete notices directory. Build/test tools in `notices/` are documented for reproducibility; the application does not bundle the development toolchain.

## Assets and tests

No third-party font, icon set, image library, Office template or user document is bundled by M1. System fonts are requested from the host operating system; their presence and redistribution rights are not implied. Adding or embedding font or media files requires a separate asset review.

Python compatibility readers and fixture generators are test-only. They are installed through `tests/requirements.txt`, are not imported by the application, and are not shipped as a production Office engine.
