# M0 execution record — 2026-11-24

Tested revision: 0a64062b037998bf9ef959476170c9d0398bfc25. GitHub master ref was re-read after push. Local script and manifest Git blob IDs verified against GitHub:

- tests/compat.py: bfd82b8817185aad67c54cc305a1d27c47ecad37
- scripts/preflight.py: 238e34fe90bc304fc9d625407317bbbd29a8e8ba
- upstream/lock.json: 1925fa7e0916e7f54327c7659e044ec0d98fa265

Environment: Debian 13 x86_64; Python 3.13.5, python-docx 1.2.0, openpyxl 3.1.5, python-pptx 1.0.2.

| Check | Actual result |
| --- | --- |
| Generate 3 synthetic fixtures | PASS |
| Check ZIP integrity and basic semantic invariants | PASS |
| Negative case: replace B4 formula with constant 42 in copy | PASS — checker rejected XLSX values/formula |
| Read-only preflight (5 CPU, 29.8 GiB disk, 5.8 GiB RAM, 0 swap, Debian) | FAIL — build prerequisites not met |
| Git network tag probe | FAIL — github.com DNS lookup failed |
| prepare_sources.py syntax and --help | PASS; local blob a16e20d3b19be77e453984d699b230be367f7704 |
| Actual upstream clone/build/launch | NOT RUN |
| ONLYOFFICE/MS Office editor round-trip | NOT RUN |
| macOS ARM64 and Windows native builds | NOT RUN |

The fixture check tests our tooling, not the selected editor. The source snapshot preparation script does not install dependencies or produce a binary; actual network use remains unverified. No upstream build was attempted on an unsuitable host. M0 exit remains blocked on complete dependency/license audit, real source build and runtime tests. No release was published.
