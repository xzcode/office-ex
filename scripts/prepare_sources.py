#!/usr/bin/env python3
"""Fetch/verify pinned top-level source snapshots, not a complete build workspace."""
import argparse
import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def git(path, *args):
    return subprocess.check_output(["git", "-C", str(path), *args], text=True).strip()


def require(ok, message):
    if not ok:
        raise RuntimeError(message)


def checkout(path, entry):
    path.mkdir()
    git(path, "init")
    git(path, "remote", "add", "origin", "https://github.com/" + entry["repo"] + ".git")
    git(path, "fetch", "--no-tags", "--depth=1", "origin", entry["commit"])
    git(path, "checkout", "--detach", entry["commit"])


def verify_repo(path, entry):
    require(path.is_dir(), f"Missing directory: {path}")
    require(Path(git(path, "rev-parse", "--show-toplevel")).resolve() == path.resolve(), f"Not a checkout: {path}")
    require(git(path, "rev-parse", "HEAD") == entry["commit"], f"Wrong revision: {path}")
    require(not git(path, "status", "--porcelain", "--untracked-files=all"), f"Dirty checkout: {path}")
    if "license_blob" in entry:
        require(git(path, "rev-parse", "HEAD:LICENSE") == entry["license_blob"], f"License mismatch: {path}")


def verify_links(desktop, entries):
    for entry in entries:
        link = git(desktop, "ls-tree", "HEAD", "--", entry["path"])
        expected = "160000 commit " + entry["commit"] + "\t" + entry["path"]
        require(link == expected, f"Gitlink mismatch: {entry['path']}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("prepare", "verify"))
    parser.add_argument("workspace", type=Path, help="new path for prepare; existing snapshot for verify")
    args = parser.parse_args()
    workspace = args.workspace.resolve()
    lock = json.loads((ROOT / "upstream/lock.json").read_text())
    desktop = workspace / "DesktopEditors"
    build = workspace / "build_tools"
    if args.action == "prepare":
        # Fail rather than touch any existing workspace; a failed fetch is left for inspection.
        workspace.mkdir(parents=True, exist_ok=False)
        checkout(desktop, lock["desktop"])
        verify_links(desktop, lock["components"])
        git(desktop, "submodule", "update", "--init", "--depth=1")
        checkout(build, lock["build_tools"])
    verify_links(desktop, lock["components"])
    verify_repo(desktop, lock["desktop"])
    for entry in lock["components"]:
        verify_repo(desktop / entry["path"], entry)
    verify_repo(build, lock["build_tools"])
    print("PASS: pinned source snapshots; recursive dependencies and compilation not verified.")


if __name__ == "__main__":
    main()
