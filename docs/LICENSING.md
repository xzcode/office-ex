# Licensing policy (M0-R2)

The repository retains the existing MIT license (Copyright 2026 Zai). The independently implemented application is intended to be MIT licensed. No third-party office suite source is present; the prior AGPL derivative plan is superseded.

Prefer MIT, BSD-2/3-Clause, Apache-2.0 and compatible permissive components for generic infrastructure. Candidate Electron, fflate and @xmldom/xmldom are MIT at project level; verify the exact dependency versions, transitive licenses, advisories, Chromium/Node notices and bundled assets prior to installation/release. Do not assume a single root MIT notice covers all bundled software. Do not import any existing office suite/editor engine, GPL/AGPL code or commercial extension as part of this implementation without explicit scope reconsideration.

Keep a dependency inventory with version, license, origin, notices, security review and reason. Review fonts, icons and images separately. Use original branding and verify trademark availability before release. Python office-format readers are test-only and not shipped.

Reference:
- https://github.com/electron/electron
- https://www.electronjs.org/docs/latest/tutorial/security
- https://github.com/101arrowz/fflate
- https://github.com/xmldom/xmldom
- https://ecma-international.org/publications-and-standards/standards/ecma-376/
