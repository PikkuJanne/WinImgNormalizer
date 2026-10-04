#!/usr/bin/env python3
"""Optional, standard-library-only Codex handoff helper (Python 3.9+).

verify: validate an extracted immutable bundle.
install: preview an additive import; write only with --apply --reviewed-head SHA.
status: read Git state; contact the remote only with --check-remote.
plan: validate task/case records, not the image-normalizer implementation.

Never stages, commits, pushes, fetches, changes branches or edits Git configuration.
Integrity hashes are NOT publisher signatures. Run in a trusted, single-writer
workspace. This helper does not defend against a malicious concurrent local actor.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import subprocess
import sys
from typing import Any, Optional

EXPECTED_REPO = "PikkuJanne/WinImgNormalizer"
MANIFEST_NAME = "BUNDLE_MANIFEST.json"
DOC_REL = "docs/codex-winimg"
STATES = {"pending", "in_progress", "blocked", "awaiting_owner", "done"}


class SafetyError(RuntimeError):
    """A failed safety condition. No destructive fallback is permitted."""


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def relative_path(value: str) -> str:
    """Accept a portable, non-traversing forward-slash relative path only."""
    if not isinstance(value, str) or not value or "\\" in value or "\x00" in value:
        raise SafetyError("Invalid relative path in manifest/plan.")
    parts = value.split("/")
    if PurePosixPath(value).is_absolute() or any(p in ("", ".", "..") for p in parts):
        raise SafetyError("Absolute or traversing path rejected.")
    for part in parts:
        if ":" in part or part.endswith((".", " ")) or part.casefold() == ".git":
            raise SafetyError("Unsafe Windows or Git-control path rejected.")
        if re.fullmatch(r"(?i)(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?", part):
            raise SafetyError("Reserved Windows filename rejected.")
    return "/".join(parts)


def no_links(root: Path, rel: str) -> Path:
    """Reject existing symlink components, including the final target."""
    current = root
    for part in relative_path(rel).split("/"):
        current = current / part
        if current.is_symlink():
            raise SafetyError("Symbolic-link target/component rejected.")
        # On Windows junctions are reparse points even when is_symlink() is false.
        if current.exists():
            attrs = getattr(current.lstat(), "st_file_attributes", 0)
            if attrs & 0x400:
                raise SafetyError("Reparse-point target/component rejected.")
    return current


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise SafetyError("Unable to read valid JSON: " + path.name) from exc


def files_without_links(root: Path) -> set[str]:
    found = set()
    for here, dirs, files in os.walk(root, followlinks=False):
        base = Path(here)
        for name in dirs + files:
            rel = (base / name).relative_to(root).as_posix()
            no_links(root, rel)
        for name in files:
            found.add((base / name).relative_to(root).as_posix())
    return found


def allowed_install_path(rel: str) -> bool:
    return (rel == "AGENTS.md" or rel.startswith("docs/codex-winimg/")
            or rel.startswith("tools/codex-winimg/"))


def verify_bundle(bundle: Path) -> dict[str, Any]:
    bundle = bundle.resolve(strict=True)
    manifest_path = no_links(bundle, MANIFEST_NAME)
    manifest = load_json(manifest_path)
    if not isinstance(manifest, dict) or manifest.get("schema_version") != 1:
        raise SafetyError("Unsupported bundle manifest.")
    if manifest.get("repository") != EXPECTED_REPO:
        raise SafetyError("Bundle is for a different repository.")
    entries = manifest.get("files")
    if not isinstance(entries, list) or not entries:
        raise SafetyError("Empty or invalid bundle manifest.")
    wanted = {MANIFEST_NAME}
    folded = {MANIFEST_NAME.casefold()}
    install_folded = set()
    for entry in entries:
        if not isinstance(entry, dict):
            raise SafetyError("Invalid manifest entry.")
        rel = relative_path(entry.get("path"))
        if rel.casefold() in folded:
            raise SafetyError("Duplicate/case-colliding manifest path.")
        folded.add(rel.casefold())
        wanted.add(rel)
        file = no_links(bundle, rel)
        if not file.is_file():
            raise SafetyError("Missing bundle file: " + rel)
        expected_hash = entry.get("sha256")
        if not isinstance(expected_hash, str) or not re.fullmatch(r"[0-9a-f]{64}", expected_hash):
            raise SafetyError("Invalid SHA-256 in manifest.")
        if not isinstance(entry.get("bytes"), int) or entry["bytes"] < 0:
            raise SafetyError("Invalid byte length in manifest.")
        if file.stat().st_size != entry["bytes"] or digest(file) != expected_hash:
            raise SafetyError("Bundle file hash/size mismatch: " + rel)
        install = entry.get("install_path")
        if rel.startswith("repo-overlay/"):
            if install != rel[len("repo-overlay/"):]:
                raise SafetyError("Overlay/install mapping mismatch.")
            install = relative_path(install)
            if not allowed_install_path(install):
                raise SafetyError("Importer may not replace application or Git-control files.")
            if install.casefold() in install_folded:
                raise SafetyError("Case-colliding install paths.")
            install_folded.add(install.casefold())
        elif install is not None:
            raise SafetyError("Only repo-overlay files may be installed.")
    actual = files_without_links(bundle)
    if actual != wanted:
        missing = len(wanted - actual)
        extra = len(actual - wanted)
        raise SafetyError(f"Bundle file set differs from manifest (missing={missing}, extra={extra}).")
    return manifest


def git(repo: Path, *args: str, allow_failure: bool = False) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env["GIT_TERMINAL_PROMPT"] = "0"
    env["GIT_OPTIONAL_LOCKS"] = "0"
    try:
        result = subprocess.run(["git", "-C", str(repo), *args],
                                stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE, text=True, encoding="utf-8",
                                errors="replace", timeout=30, env=env, check=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise SafetyError("Git unavailable or timed out; no synchronization claim is possible.") from exc
    if result.returncode and not allow_failure:
        # Do not echo stderr: Git can include credential-bearing URLs in it.
        raise SafetyError(f"git {args[0] if args else ''} failed (exit {result.returncode}); inspect locally without sharing credentials.")
    return result


def github_identity(url: str) -> Optional[str]:
    patterns = (
        r"https://github\.com/([^/\s]+)/([^/\s]+?)(?:\.git)?/?",
        r"git@github\.com:([^/\s]+)/([^/\s]+?)(?:\.git)?/?",
        r"ssh://git@github\.com/([^/\s]+)/([^/\s]+?)(?:\.git)?/?",
    )
    for pattern in patterns:
        match = re.fullmatch(pattern, url.strip(), flags=re.IGNORECASE)
        if match:
            owner, name = match.groups()
            if any(c in owner + name for c in "?#@:"):
                return None
            return (owner + "/" + name).casefold()
    return None


def inspect_repo(repo: Path, check_remote: bool = False, remote: str = "origin",
                 expected_identity: Optional[str] = EXPECTED_REPO) -> dict[str, Any]:
    """Read-only inspection. expected_identity=None is used ONLY by local bare-repo tests."""
    repo = repo.resolve(strict=True)
    top = Path(git(repo, "rev-parse", "--show-toplevel").stdout.strip()).resolve(strict=True)
    if top != repo:
        raise SafetyError("Pass the repository root, not a nested folder.")
    head = git(repo, "rev-parse", "HEAD").stdout.strip()
    if not re.fullmatch(r"[0-9a-f]{40,64}", head):
        raise SafetyError("No valid local commit is available.")
    branch_result = git(repo, "symbolic-ref", "--quiet", "--short", "HEAD", allow_failure=True)
    branch = branch_result.stdout.strip() if branch_result.returncode == 0 else None
    dirty = bool(git(repo, "status", "--porcelain=v1", "--untracked-files=all", "-z").stdout)
    fetch_urls = git(repo, "remote", "get-url", "--all", remote).stdout.splitlines()
    push_urls = git(repo, "remote", "get-url", "--push", "--all", remote).stdout.splitlines()
    if len(fetch_urls) != 1 or len(push_urls) != 1:
        raise SafetyError("Exactly one reviewed fetch URL and one push URL are required.")
    if expected_identity:
        expected = expected_identity.casefold()
        if github_identity(fetch_urls[0]) != expected or github_identity(push_urls[0]) != expected:
            raise SafetyError("Fetch/push identity is not the expected GitHub repository. URLs are withheld to avoid credential leakage.")
    result = {
        "repository": expected_identity or "local-test-repository",
        "branch": branch, "head": head, "worktree_clean": not dirty,
        "remote_name": remote, "remote_head": None, "sync_state": "not_checked",
        "ci_state": "not_observed", "network_check_performed": check_remote,
    }
    if check_remote:
        if not branch:
            result["sync_state"] = "detached_head"
            return result
        ref = "refs/heads/" + branch
        try:
            response = git(repo, "ls-remote", "--exit-code", "--refs", remote, ref, allow_failure=True)
        except SafetyError:
            result["sync_state"] = "not_verified"
            return result
        if response.returncode == 2:
            result["sync_state"] = "remote_branch_missing"
            return result
        if response.returncode != 0:
            result["sync_state"] = "not_verified"
            return result
        matches = [line.split() for line in response.stdout.splitlines()
                   if len(line.split()) == 2 and line.split()[1] == ref]
        if len(matches) != 1 or not re.fullmatch(r"[0-9a-f]{40,64}", matches[0][0]):
            result["sync_state"] = "not_verified"
            return result
        remote_head = matches[0][0]
        result["remote_head"] = remote_head
        if remote_head != head:
            result["sync_state"] = "different_commits"
        elif dirty:
            result["sync_state"] = "remote_match_but_worktree_dirty"
        else:
            result["sync_state"] = "synchronized"
    return result


def install_bundle(bundle: Path, repo: Path, apply: bool = False,
                   reviewed_head: Optional[str] = None) -> dict[str, Any]:
    bundle, repo = bundle.resolve(strict=True), repo.resolve(strict=True)
    # Keep the immutable reference outside the checkout; it must not become accidental staged data.
    try:
        bundle.relative_to(repo)
        raise SafetyError("Extract the bundle outside the target repository.")
    except ValueError:
        pass
    manifest = verify_bundle(bundle)
    state = inspect_repo(repo)
    blockers = []
    if not state["worktree_clean"]:
        blockers.append("worktree_not_clean")
    if not state["branch"] or not state["branch"].startswith("codex/winimg-"):
        blockers.append("dedicated_codex_winimg_feature_branch_required")
    if apply and (not reviewed_head or reviewed_head != state["head"]):
        blockers.append("reviewed_head_missing_or_changed")
    plan = []
    for entry in manifest["files"]:
        rel = entry.get("install_path")
        if rel is None:
            continue
        try:
            target = no_links(repo, rel)
            if target.exists():
                same = target.is_file() and digest(target) == entry["sha256"]
                action = "unchanged" if same else "conflict"
            else:
                # A non-directory existing parent would make later mkdir fail; catch it in preview.
                parent = target.parent
                while parent != repo:
                    if parent.exists() and not parent.is_dir():
                        raise SafetyError("Existing parent is not a directory.")
                    parent = parent.parent
                action = "add"
        except SafetyError:
            action = "unsafe_path"
        plan.append({"path": rel, "action": action})
        if action in ("conflict", "unsafe_path"):
            blockers.append(action + ":" + rel)
    report = {"mode": "apply" if apply else "preview", "head": state["head"],
              "branch": state["branch"], "blockers": blockers, "files": plan,
              "files_added": 0, "git_writes_performed": False}
    if not apply:
        return report
    if blockers:
        raise SafetyError("Import blocked before writes: " + "; ".join(blockers))
    # Recheck the clean HEAD immediately before writing. Assumes no concurrent writer.
    fresh = inspect_repo(repo)
    if fresh["head"] != reviewed_head or fresh["branch"] != state["branch"] or not fresh["worktree_clean"]:
        raise SafetyError("Repository changed after preview; import stopped.")
    by_path = {e["install_path"]: e for e in manifest["files"] if e.get("install_path")}
    try:
        for item in plan:
            if item["action"] != "add":
                continue
            rel = item["path"]
            entry = by_path[rel]
            target = no_links(repo, rel)
            target.parent.mkdir(parents=True, exist_ok=True)
            target = no_links(repo, rel)
            data = no_links(bundle, entry["path"]).read_bytes()
            if hashlib.sha256(data).hexdigest() != entry["sha256"]:
                raise SafetyError("Bundle changed during import.")
            # Exclusive create: never silently overwrite a file that appeared after planning.
            with target.open("xb") as output:
                output.write(data)
            report["files_added"] += 1
    except (OSError, SafetyError) as exc:
        raise SafetyError("Import stopped during writing. Some new handoff files may exist; inspect them. No rollback, overwrite or Git operation was attempted.") from exc
    return report


def validate_plan(repo: Path) -> dict[str, Any]:
    root = repo.resolve(strict=True)
    doc = root / DOC_REL
    registry = load_json(doc / "TASKS.json")
    corpus = load_json(doc / "TEST_CASES.json")
    tasks = registry.get("tasks", [])
    cases = corpus.get("cases", [])
    if not tasks or not cases:
        raise SafetyError("Plan has no tasks or test cases.")
    task_map = {t["id"]: t for t in tasks}
    case_map = {c["id"]: c for c in cases}
    if len(task_map) != len(tasks) or len(case_map) != len(cases):
        raise SafetyError("Duplicate task or case IDs.")
    if registry.get("next_task") is not None and registry["next_task"] not in task_map:
        raise SafetyError("Unknown next_task.")
    visited, active = set(), set()
    def visit(id: str) -> None:
        if id in active:
            raise SafetyError("Task dependency cycle.")
        if id in visited:
            return
        active.add(id)
        for dep in task_map[id].get("depends_on", []):
            if dep not in task_map:
                raise SafetyError("Unknown task dependency.")
            visit(dep)
        active.remove(id)
        visited.add(id)
    for id in task_map:
        visit(id)
    coverage = set()
    for id, task in task_map.items():
        if not re.fullmatch(r"M\d+-T\d+", id) or task.get("state") not in STATES:
            raise SafetyError("Invalid task ID or state.")
        if not no_links(doc, task["card"]).is_file():
            raise SafetyError("Task card missing.")
        for ref in task.get("references", []):
            if not no_links(doc, ref).is_file():
                raise SafetyError("Task reference missing: " + ref)
        coverage.update(task.get("requirements", []))
        for cid in task.get("case_ids", []):
            if cid not in case_map:
                raise SafetyError("Unknown test case reference.")
        if task["state"] in {"in_progress", "awaiting_owner", "done"}:
            if any(task_map[d]["state"] != "done" for d in task.get("depends_on", [])):
                raise SafetyError("Active/done task has an unaccepted dependency.")
        if task["state"] == "done":
            evidence_paths = task.get("evidence", [])
            if not evidence_paths:
                raise SafetyError("Done task has no evidence record.")
            evidence = []
            for path in evidence_paths:
                record = load_json(no_links(doc, path))
                if record.get("task_id") != id:
                    raise SafetyError("Evidence task ID mismatch.")
                evidence.append(record)
            outcomes = {}
            for record in evidence:
                for item in record.get("case_results", []):
                    outcomes[item["case_id"]] = item
            for cid in task.get("case_ids", []):
                outcome = outcomes.get(cid, {})
                if outcome.get("state") == "passed":
                    continue
                if (outcome.get("state") == "skipped" and outcome.get("reason") and
                        case_map[cid]["kind"] in {"optional-codec", "windows-capability"}):
                    continue
                raise SafetyError("Done task lacks an accepted required case outcome: " + cid)
            if task.get("owner_gate"):
                approvals = [record.get("owner_approval", {}) for record in evidence]
                if not any(a.get("state") == "approved" and a.get("actor") and
                           a.get("recorded_utc") and a.get("approved_commit") for a in approvals):
                    raise SafetyError("Owner-gate task lacks actual approval evidence fields.")
    if not set(range(1, 19)).issubset(coverage):
        raise SafetyError("An accepted review item has no task coverage.")
    for cid, case in case_map.items():
        owner = case.get("owner_task")
        if owner not in task_map or cid not in task_map[owner].get("case_ids", []):
            raise SafetyError("Test case owner/task mapping is inconsistent.")
    next_task = registry.get("next_task")
    if next_task and task_map[next_task]["state"] == "done":
        raise SafetyError("next_task points at an already accepted task.")
    return {"tasks": len(tasks), "cases": len(cases), "review_items_covered": 18,
            "tasks_done": sum(t["state"] == "done" for t in tasks),
            "next_task": next_task,
            "scope": "Structural plan validation only; no application tests or approval authentication."}


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    verify = sub.add_parser("verify", help="Verify the extracted immutable bundle.")
    verify.add_argument("--bundle", type=Path, required=True)
    install = sub.add_parser("install", help="Preview by default; explicit --apply writes new handoff files.")
    install.add_argument("--bundle", type=Path, required=True)
    install.add_argument("--repo", type=Path, required=True)
    install.add_argument("--apply", action="store_true")
    install.add_argument("--reviewed-head")
    status = sub.add_parser("status", help="Read-only Git state; optional live remote SHA check.")
    status.add_argument("--repo", type=Path, required=True)
    status.add_argument("--check-remote", action="store_true")
    status.add_argument("--remote-name", default="origin")
    plan = sub.add_parser("plan", help="Validate tracked task/case records, not the application.")
    plan.add_argument("--repo", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        code = 0
        if args.command == "verify":
            manifest = verify_bundle(args.bundle)
            result = {"verified": True, "bundle_id": manifest.get("bundle_id"),
                      "files_verified": len(manifest["files"]),
                      "note": "Integrity only, not a publisher signature or application validation."}
        elif args.command == "install":
            result = install_bundle(args.bundle, args.repo, args.apply, args.reviewed_head)
            code = 2 if result["blockers"] else 0
        elif args.command == "status":
            result = inspect_repo(args.repo, args.check_remote, args.remote_name)
            code = 0 if not args.check_remote or result["sync_state"] == "synchronized" else 2
        else:
            result = validate_plan(args.repo)
        print(json.dumps(result, indent=2, ensure_ascii=False))
        return code
    except (SafetyError, OSError, KeyError, TypeError, ValueError) as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, indent=2), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
