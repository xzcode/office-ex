#!/usr/bin/env python3
"""Read-only Linux source build preflight; no installation or remote mutation."""
import argparse
import json
import os
import platform
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GIB = 1024 ** 3

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--network", action="store_true", help="verify pinned desktop tag via git")
    parser.add_argument("--workspace", type=Path, default=ROOT, help="existing build volume path")
    args = parser.parse_args()
    lock = json.loads((ROOT / "upstream/lock.json").read_text())
    for entry in [lock["desktop"], lock["build_tools"], *lock["components"]]:
        if not re.fullmatch(r"[0-9a-f]{40}", entry["commit"]):
            raise ValueError("Invalid pinned commit: " + entry["repo"])
    checks = []
    def check(name, ok, detail):
        checks.append(ok)
        print(f"{'PASS' if ok else 'FAIL'} {name}: {detail}")
    check("platform", platform.system() == "Linux" and platform.machine() in ("x86_64", "AMD64"), platform.platform())
    release = platform.freedesktop_os_release() if platform.system() == "Linux" else {}
    check("OS", release.get("ID") == "ubuntu" and release.get("VERSION_ID") == "24.04", release.get("PRETTY_NAME", "unknown"))
    cpus = len(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else (os.cpu_count() or 0)
    check("CPU", cpus >= 4, str(cpus))
    disk = shutil.disk_usage(args.workspace).free
    check("free disk", disk >= 100 * 10**9, f"{disk/GIB:.1f} GiB (require 100 GB)")
    mem = {}
    if Path("/proc/meminfo").is_file():
        for line in Path("/proc/meminfo").read_text().splitlines():
            key, value = line.split(":", 1)
            mem[key] = int(value.split()[0]) * 1024
    for key, minimum in [("MemTotal", 8 * 10**9), ("SwapTotal", 4 * 10**9)]:
        value = mem.get(key, 0)
        check(key, value >= minimum, f"{value/GIB:.1f} GiB")
    for tool in ("git", "python3", "sudo"):
        check(tool, shutil.which(tool) is not None, shutil.which(tool) or "not installed")
    if args.network:
        tag = lock["desktop"]["tag"]
        url = "https://github.com/" + lock["desktop"]["repo"] + ".git"
        try:
            result = subprocess.run(["git", "ls-remote", url, "refs/tags/" + tag, "refs/tags/" + tag + "^{}"], capture_output=True, text=True, timeout=20)
            refs = {line.split()[1]: line.split()[0] for line in result.stdout.splitlines() if len(line.split()) == 2}
            sha = refs.get("refs/tags/" + tag + "^{}", refs.get("refs/tags/" + tag))
            check("upstream tag", result.returncode == 0 and sha == lock["desktop"]["commit"], sha or result.stderr.strip() or "missing tag")
        except (OSError, subprocess.TimeoutExpired) as exc:
            check("upstream tag", False, str(exc))
    else:
        print("SKIP network (use --network)")
    print("Preflight only; not a compilation or dependency lock verification.")
    return 0 if all(checks) else 1

if __name__ == "__main__":
    sys.exit(main())
