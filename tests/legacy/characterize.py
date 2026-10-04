"""Bounded observations of a narrowly instrumented legacy snapshot on Windows.

Requires an existing, verified ImageMagick executable. Never downloads tools,
dot-sources the app, or runs the original entry point against Known Folders.
Raw artifacts go exclusively to a new ignored, marked scratch directory.
"""
import argparse
import hashlib
import json
import ntpath
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys
from datetime import datetime, timezone

import PIL
from PIL import Image


HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def save_json(path, value):
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def run_owned(argv, *, env=None, timeout=60, artifact_dir=None):
    # Terminate only this worker's process tree and retain timeout failures.
    child_env = dict(os.environ if env is None else env)
    if Path(argv[0]).name.lower() in ("powershell.exe", "pwsh.exe"):
        # Let each edition build its own default module path instead of inheriting
        # an incompatible PS7 module path through the Python controller.
        for key in list(child_env):
            if key.casefold() == "psmodulepath":
                del child_env[key]
    process = subprocess.Popen(argv, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               env=child_env, encoding="utf-8", errors="replace")
    try:
        out, err = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        killed = subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"],
                                capture_output=True, timeout=15)
        if killed.returncode and process.poll() is None:
            process.kill()
        out, err = process.communicate(timeout=15)
        report = {"command": argv, "owned_pid": process.pid, "timeout_seconds": timeout,
                  "stdout": out, "stderr": err, "taskkill_exit_code": killed.returncode,
                  "taskkill_stdout": killed.stdout.decode("utf-8", errors="replace"),
                  "taskkill_stderr": killed.stderr.decode("utf-8", errors="replace"),
                  "tree_termination_confirmed": killed.returncode == 0}
        if artifact_dir:
            save_json(artifact_dir / "timeout.json", report)
        raise RuntimeError(f"Owned worker timed out; tree termination confirmed={killed.returncode == 0}: {argv[0]}")
    return {"exit_code": process.returncode, "stdout": out, "stderr": err}


def shell_command(shell, script, *args):
    return [str(shell), "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
            "-File", str(script), *map(str, args)]


def source_state(source):
    result = {}
    for path in sorted(source.rglob("*")):
        if path.is_file():
            stat = path.stat()
            result[path.relative_to(source).as_posix()] = {
                "sha256": sha256(path), "length": stat.st_size,
                "creation_ns": stat.st_birthtime_ns, "modified_ns": stat.st_mtime_ns,
            }
    return result


def inspect_outputs(pictures):
    outputs = []
    for path in sorted(pictures.rglob("*")):
        if not path.is_file():
            continue
        record = {"path": path.relative_to(pictures).as_posix(), "bytes": path.stat().st_size,
                  "sha256": sha256(path), "creation_ns": path.stat().st_birthtime_ns,
                  "modified_ns": path.stat().st_mtime_ns}
        if path.suffix.lower() in (".jpeg", ".jpg"):
            try:
                with Image.open(path) as image:
                    image.load()  # Complete independent decode, not only header detection.
                    record["decode"] = {"valid": True, "format": image.format,
                                        "size": list(image.size), "frames": getattr(image, "n_frames", 1),
                                        "corner_rgb": list(image.convert("RGB").getpixel((0, 0)))}
            except Exception as exc:
                record["decode"] = {"valid": False, "error_type": type(exc).__name__}
        outputs.append(record)
    return outputs


def nested_planning_observation():
    # Fictional Windows paths only; do not create them or enumerate recursively.
    source = r"Q:\synthetic\Pictures"
    destination = ntpath.join(source, "Pictures_WinImgNormalized_MOCK")
    included_relative = ntpath.relpath(destination, source)
    mirrored = ntpath.join(destination, included_relative)
    return {"kind": "mocked_path_planning", "filesystem_execution": False,
            "source": source, "destination": destination,
            "destination_is_inside_source": ntpath.commonpath([source, destination]) == source,
            "first_mirrored_descendant": mirrored,
            "finding": "Destination created before enumeration is eligible as a source descendant; actual recursion behavior remains not_run."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--magick", type=Path, required=True)
    parser.add_argument("--ps51", type=Path, default=Path(shutil.which("powershell") or "powershell.exe"))
    parser.add_argument("--ps7", type=Path, default=Path(shutil.which("pwsh") or "pwsh.exe"))
    args = parser.parse_args()
    if os.name != "nt":
        parser.error("Actual Windows required; no substitute compatibility claim")
    root = args.root.resolve()
    scratch = (REPO / ".scratch").resolve()
    if root == scratch or scratch not in root.parents or root.exists():
        parser.error("--root must be a new child of the repository's ignored .scratch directory")
    if any(parent.is_symlink() or parent.is_junction() for parent in [root, *root.parents] if parent.exists()):
        parser.error("Scratch ancestors must not be reparse points")
    ignored = subprocess.run(["git", "check-ignore", "--quiet", str(root)], cwd=REPO)
    if ignored.returncode:
        parser.error("Scratch root must be ignored before execution")
    magick = args.magick.resolve(strict=True)
    if magick.name.lower() != "magick.exe":
        parser.error("Real --magick executable must be named magick.exe")
    root.parent.mkdir(parents=True, exist_ok=True)
    build_command = [sys.executable, "-B", str(HERE / "build_fixtures.py"),
                     "--root", str(root), "--magick", str(magick)]
    build = run_owned(build_command, timeout=120)
    if root.is_dir():
        save_json(root / "build-transcript.json", {"command": build_command, **build})
    if build["exit_code"]:
        raise RuntimeError(f"Fixture generation failed, not skipped: {build}")
    environment = {"schema_version": 1, "recorded_utc": datetime.now(timezone.utc).isoformat(),
                   "python": platform.python_version(), "python_architecture": platform.machine(),
                   "pillow": PIL.__version__, "shell_environment_adapter": "Remove inherited PSModulePath in worker environment; each edition resolves its own defaults",
                   "imagemagick_executable_sha256": sha256(magick),
                   "imagemagick": run_owned([str(magick), "-version"]),
                   "imagemagick_formats": run_owned([str(magick), "-list", "format"]),
                   "imagemagick_delegates": run_owned([str(magick), "-list", "delegate"]),
                   "imagemagick_policy": run_owned([str(magick), "-list", "policy"]), "shells": {}}
    executions = []
    checks = []
    for label, shell in [("ps51", args.ps51.resolve()), ("ps7", args.ps7.resolve())]:
        inventory_command = shell_command(shell, HERE / "Inventory-Environment.ps1")
        inventory = run_owned(inventory_command)
        if inventory["exit_code"]:
            raise RuntimeError(f"Inventory failure: {inventory}")
        environment["shells"][label] = json.loads(inventory["stdout"].lstrip("\ufeff"))
        expected_major = 5 if label == "ps51" else 7
        if int(environment["shells"][label]["powershell_version"].split(".")[0]) != expected_major:
            raise RuntimeError(f"Wrong shell for {label}")
        environment["shells"][label]["inventory_command"] = inventory_command
        for case, folder, fake_mode in [
            ("ordinary", "ordinary", None), ("collision", "collision", None),
            ("frames", "frames", None), ("failed_first_duplicate", "duplicate", "first_failure"),
            ("weak_zero", "weak", "weak_zero"), ("weak_invalid", "weak", "weak_invalid"),
        ]:
            work = root / "runs" / f"{label}-{case}"
            work.mkdir(parents=True)
            pictures = work / "pictures"
            source = root / "corpus" / folder
            env = os.environ.copy()
            executable = magick
            if fake_mode:
                fake = work / "fake-native"
                fake.mkdir()
                executable = fake / "magick.cmd"
                # Native cmd forwards only fixed, synthetic M0 case paths.
                executable.write_text(f'@echo off\n"{sys.executable}" -B "{HERE / "fake_magick.py"}" %*\nexit /b %errorlevel%\n', encoding="utf-8")
                env["WINIMG_FAKE_TRACE"] = str(work / "native-calls.jsonl")
            command = shell_command(shell, HERE / "Invoke-LegacySnapshot.ps1",
                                    "-Repository", REPO, "-ScratchRoot", root, "-Source", source,
                                    "-PicturesRoot", pictures, "-MagickPath", executable)
            if fake_mode:
                command.extend(["-FakeMode", fake_mode])
            before = source_state(source)
            result = run_owned(command, env=env, timeout=60, artifact_dir=work)
            after = source_state(source)
            preserved = before == after
            if not preserved:
                raise RuntimeError(f"Source changed in {label}/{case}; stop and preserve evidence")
            outputs = inspect_outputs(pictures)
            logs = [p for p in pictures.rglob("*.log") if p.is_file()]
            log_text = "\n".join(p.read_text(encoding="utf-8-sig", errors="replace") for p in logs)
            match = re.search(r"SUMMARY ConvertedImages=(\d+) CopiedVideos=(\d+) Duplicates=(\d+) Unsupported=(\d+) Errors=(\d+)", log_text)
            stats = dict(zip(["converted", "copied_video", "duplicates", "unsupported", "errors"], map(int, match.groups()))) if match else None
            snapshots = [json.loads(p.read_text(encoding="utf-8-sig")) for p in root.rglob("containment.json")]
            # Keep all source state, streams and native traces for raw audit.
            execution = {"shell": label, "case": case, "native": "controlled_fake" if fake_mode else "real_imagemagick",
                         "command": command, **result, "source_preserved": preserved,
                         "source_before": before, "source_after": after, "outputs": outputs,
                         "output_directories": [p.relative_to(pictures).as_posix() for p in sorted(pictures.rglob("*")) if p.is_dir()],
                         "application_log": log_text, "summary": stats,
                         "containment_records_seen": len(snapshots)}
            if fake_mode:
                execution["fake_trace"] = [json.loads(line) for line in (work / "native-calls.jsonl").read_text(encoding="utf-8").splitlines()]
            if result["exit_code"] != 0 or stats is None:
                raise RuntimeError(f"Legacy observation infrastructure incomplete: {label}/{case}: {result}")
            if case == "ordinary":
                video_suffixes = (".mp4", ".mov", ".mkv")
                videos = {name: v["sha256"] for name, v in before.items() if name.endswith(video_suffixes)}
                copied = {o["path"].split("/", 1)[1]: o["sha256"] for o in outputs if o["path"].endswith(video_suffixes)}
                execution["video_bytes_match"] = bool(videos) and videos == copied
                images = [o for o in outputs if o["path"].endswith(".jpeg")]
                execution["all_jpeg_outputs_decode"] = bool(images) and all(o["decode"]["valid"] for o in images)
                mapped_sources = {str(Path(name).with_suffix(".jpeg")).replace("\\", "/")
                                  if name.endswith((".jpg", ".png", ".bmp", ".tiff")) else name: value
                                  for name, value in before.items()}
                execution["output_times_match_source"] = all(
                    o["path"].split("/", 1)[1] in mapped_sources and
                    o["creation_ns"] == mapped_sources[o["path"].split("/", 1)[1]]["creation_ns"] and
                    o["modified_ns"] == mapped_sources[o["path"].split("/", 1)[1]]["modified_ns"]
                    for o in outputs if not o["path"].endswith(".log"))
                if not (execution["video_bytes_match"] and execution["all_jpeg_outputs_decode"] and execution["output_times_match_source"]):
                    raise RuntimeError(f"Ordinary observation's integrity checks failed: {label}")
            save_json(work / "execution.json", execution)
            executions.append(execution)
            checks.append({"shell": label, "check": f"{case}: complete capture and source preservation", "passed": True})

        # Rejection controls run BEFORE the app snapshot can execute.
        for control in ["missing_marker", "outside_source", "overlap", "wrong_hash"]:
            rejected = root / "controls" / f"{label}-{control}"
            rejected.mkdir(parents=True)
            controls_source = root / "corpus" / "ordinary"
            controls_scratch = root
            controls_pictures = rejected / "pictures"
            extra = []
            if control == "missing_marker":
                controls_scratch = rejected
                controls_source = rejected / "source"
                controls_source.mkdir()
            elif control == "outside_source":
                controls_source = REPO  # Read-only path validation must reject before touching it.
            elif control == "overlap":
                controls_pictures = controls_source / "forbidden-test-output"
            elif control == "wrong_hash":
                extra = ["-ExpectedLegacySha256", "0" * 64]
            command = shell_command(shell, HERE / "Invoke-LegacySnapshot.ps1",
                                    "-Repository", REPO, "-ScratchRoot", controls_scratch,
                                    "-Source", controls_source, "-PicturesRoot", controls_pictures,
                                    "-MagickPath", magick, *extra)
            result = run_owned(command, timeout=15)
            if result["exit_code"] == 0 or controls_pictures.exists():
                raise RuntimeError(f"Containment control failed: {label}/{control}")
            checks.append({"shell": label, "check": control, "passed": True, "command": command, **result})
        # Confirm process-only Bypass did not modify persistent policy.
        after_inventory = run_owned(inventory_command)
        if after_inventory["exit_code"]:
            raise RuntimeError("Post-run policy inventory failed")
        after_policy = json.loads(after_inventory["stdout"].lstrip("\ufeff"))["execution_policy"]
        if after_policy != environment["shells"][label]["execution_policy"]:
            raise RuntimeError("Execution policy changed; stop")
        checks.append({"shell": label, "check": "persistent execution policy unchanged", "passed": True})
    save_json(root / "environment.json", environment)
    result = {"schema_version": 1, "recorded_utc": datetime.now(timezone.utc).isoformat(),
              "legacy_sha256": sha256(REPO / "WinImgNormalizer.ps1"),
              "test_files_sha256": {p.relative_to(REPO).as_posix(): sha256(p) for p in sorted(HERE.iterdir()) if p.is_file()},
              "fixture_manifest_sha256": sha256(root / "manifest.json"),
              "environment": environment, "executions": executions, "harness_checks": checks,
              "nested_planning": nested_planning_observation(),
              "containment": [json.loads(p.read_text(encoding="utf-8-sig")) for p in root.rglob("containment.json")]}
    save_json(root / "results.json", result)
    print(json.dumps({"artifact": str(root / "results.json"), "legacy_observations": len(executions),
                      "harness_checks_passed": len(checks), "real_runs": 6, "controlled_fake_runs": 6,
                      "regression_suite": False}))


if __name__ == "__main__":
    main()
