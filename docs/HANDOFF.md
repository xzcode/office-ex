# Handoff — M0, 2026-11-24

Repository: xzcode/office-ex, master. Initially public, MIT license only.

Completed: upstream desktop v10.0.0/tag and six pinned gitlinks identified; build_tools master snapshot and available Linux automation verified; core license text reviewed; source policy, architecture, compatibility scope, preflight and starter fixture tooling committed.

Build environment observation: Linux x86_64 Debian 13, about 30 GiB available disk, 5.8 GiB RAM, no swap, no Docker CLI; git ls-remote to github.com failed due to DNS. Upstream documented Linux build baseline is Ubuntu 24.04 amd64 with 4 cores, 8 GB RAM, 4 GB swap, 100 GB storage. No actual upstream source compilation or Office EX runtime has been completed. No source tarball or binary release. Do not claim fixture tests prove engine interoperability.

Resume in order:
1. Run scripts/preflight.py --network on a dedicated Ubuntu build machine; verify storage, RAM, swap, internet and available tools.
2. Retrieve exact source revisions from upstream/lock.json and audit build_tools additional dependencies, toolchain downloads, versions, checksums and licenses. Disable moving updates for the locked build. Never modify unrelated server projects.
3. Complete Linux source compilation and native launch, record commands/logs/SHA256, then macOS ARM64 build and launch. Run fixture round-trip and visual checks; record outstanding risks.
4. Finish M0 gates in docs/M0.md, then begin M1. Do not create release or run unnecessary builds before that.
