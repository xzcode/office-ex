# Licensing / third-party policy

This is a preliminary engineering review, not legal advice.

The existing MIT LICENSE (Copyright 2026 Zai) applies to the independent M0 scripts and documentation. The upstream code is not yet vendored or redistributed. When a combined derivative is distributed, the covered work must be licensed under AGPLv3 plus applicable terms; MIT does not replace those obligations. Preserve the MIT notice for applicable original files and establish accurate root licensing scope before importing the engine.

Verified LICENSE blob 53bc62cde3c705aa2d03b895b966bd6250f389ae for the pinned DesktopEditors, core, desktop-apps, desktop-sdk, sdkjs, web-apps, and build_tools revisions. Its additional terms require retention of attribution and notices, identification/date of modifications, visible accessible legal notices in the UI, and no automatic grant of upstream trademarks. Certain non-code assets use CC BY-SA 4.0. The dictionaries repository has no root LICENSE at the pinned revision; audit each included dictionary separately.

Distribution gates:
- Include the applicable full unmodified licenses and copyright notices, document modifications and dates, and make complete corresponding source/build scripts available for each binary.
- Display prominent legal notices identifying the original developer Ascensio System SIA, modified status and license details; product branding remains independent and avoids confusing endorsement.
- Audit all third-party libraries and assets, including fonts, spellcheck dictionaries, Qt/CEF, plugins, translations, SDK assets and any binary dependency; record version, SPDX identifier, source URL and obligations.
- Exclude proprietary components, commercial addons and assets without redistribution permission. Do not simply strip upstream legal notices, or publish a binary without corresponding source.
- Do not use Microsoft or upstream logos as our app logo. Office EX is a working name; trademark clearance is a separate release gate.
- Test tooling dependencies are not shipped in the office application; review their own licenses if redistributed.

Sources:
- https://github.com/ONLYOFFICE/DesktopEditors/blob/daf8c6948148827201b2907ac1134d222fbfa7b4/LICENSE
- https://www.onlyoffice.com/trademark-policy
- https://www.gnu.org/licenses/agpl-3.0.html
- https://creativecommons.org/licenses/by-sa/4.0/
