# Reviewed PDF font assets

Office EX uses eight unmodified Noto faces from the exact upstream commits in `manifest.json`. Run `npm run setup:fonts` once after checkout; this fetches only pinned HTTPS sources and verifies length and SHA-256 before installation. Subsequent builds and all document operations use local bytes, with no font downloads or network access at runtime. The complete upstream OFL-1.1 copyright and license notices are retained as `NotoCore-COPYRIGHT.txt` and `NotoCJK-COPYRIGHT.txt`.

The manifest records the actual font names, embedded license strings and OS/2 fsType=0. The Unicode coverage table in `packages/pdf/src/font-coverage.ts` reflects the union of these fonts’ Unicode cmap format 4/12/13 mappings, excluding missing-glyph entries. Variation-selector coverage and all script shaping are not guaranteed by this table alone. The approved font fallback includes a single SC CJK face; no assertion of full locale-specific CJK typography is made. Arbitrary system fonts are not approved for embedding.

No font is relicensed under MIT. This setup process is required for a source build and CI, not for using a built application. Do not weaken the manifest, embedding, or missing-glyph checks.
