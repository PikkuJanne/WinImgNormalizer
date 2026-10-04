"""Exercise containment and controlled legacy outcomes without ImageMagick.

This is harness validation and fake-process characterization, not real codec
coverage or a passing regression suite for the unfixed application.
"""
import argparse
import json
import os
from pathlib import Path
import struct
import sys
import zlib

from PIL import Image

from build_fixtures import fixed_times
from characterize import HERE, REPO, inspect_outputs, run_owned, save_json, sha256, shell_command, source_state


def png(path, rgb):
    # Standard-library PNG recipe; Pillow is the independent full decoder.
    def chunk(kind, body):
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body))
    row = b"\0" + bytes(rgb) * 24
    path.write_bytes(b"\x89PNG\r\n\x1a\n" +
                     chunk(b"IHDR", struct.pack(">IIBBBBB", 24, 16, 8, 2, 0, 0, 0)) +
                     chunk(b"IDAT", zlib.compress(row * 16, 9)) + chunk(b"IEND", b""))
    with Image.open(path) as image:
        image.load()
        assert image.format == "PNG" and image.size == (24, 16)
        assert set(image.get_flattened_data()) == {tuple(rgb)}
    fixed_times(path)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--ps51", type=Path, required=True)
    parser.add_argument("--ps7", type=Path, required=True)
    args = parser.parse_args()
    root = args.root.absolute()
    scratch = (REPO / ".scratch").resolve()
    if os.name != "nt" or root.exists() or scratch not in root.parents:
        parser.error("Use a new child of the ignored .scratch directory on Windows")
    # Use the same path/ignore guards as the full controller before any writes.
    import subprocess
    if any(p.is_symlink() or p.is_junction() for p in [root, *root.parents] if p.exists()):
        parser.error("No reparse ancestors permitted")
    if subprocess.run(["git", "check-ignore", "--quiet", str(root)], cwd=REPO).returncode:
        parser.error("Scratch must already be ignored")
    root.mkdir(parents=True)
    (root / ".winimg-fixture-root").write_text("synthetic harness validation\n", encoding="utf-8")
    fixtures = []
    for folder, rgb in [("duplicate/a", (224, 32, 32)), ("duplicate/b", (32, 224, 32)), ("weak", (32, 32, 224))]:
        path = root / "corpus" / folder / "source.png"
        path.parent.mkdir(parents=True)
        png(path, rgb)
        fixtures.append({"path": path.relative_to(root).as_posix(), "recipe": "standard-library PNG chunks; seeded-free uniform RGB",
                         "privacy_status": "synthetic", "rgb": list(rgb), "format": "PNG", "dimensions": [24, 16],
                         "independent_decoder": "Pillow complete decode and all-pixel comparison", "sha256": sha256(path)})
    save_json(root / "manifest.json", {"schema_version": 1, "fixtures": fixtures, "real_imagemagick": False})
    records = []
    controls = []
    environment = {}
    for label, shell in [("ps51", args.ps51), ("ps7", args.ps7)]:
        inventory_command = shell_command(shell, HERE / "Inventory-Environment.ps1")
        inventory = run_owned(inventory_command)
        save_json(root / f"{label}-inventory-transcript.json", {"command": inventory_command, **inventory})
        if inventory["exit_code"]:
            raise RuntimeError("Inventory failed")
        environment[label] = json.loads(inventory["stdout"].lstrip("\ufeff"))
        if int(environment[label]["powershell_version"].split(".")[0]) != (5 if label == "ps51" else 7):
            raise RuntimeError("Wrong target shell")
        for mode in ["first_failure", "weak_zero", "weak_invalid"]:
            work = root / "runs" / f"{label}-{mode}"
            fake = work / "fake-native"
            fake.mkdir(parents=True)
            executable = fake / "magick.cmd"
            executable.write_text(f'@echo off\n"{sys.executable}" -B "{HERE / "fake_magick.py"}" %*\nexit /b %errorlevel%\n', encoding="utf-8")
            env = os.environ.copy()
            env["WINIMG_FAKE_TRACE"] = str(work / "native-calls.jsonl")
            source = root / "corpus" / ("duplicate" if mode == "first_failure" else "weak")
            pictures = work / "pictures-ää-😀"
            command = shell_command(shell, HERE / "Invoke-LegacySnapshot.ps1",
                                    "-Repository", REPO, "-ScratchRoot", root, "-Source", source,
                                    "-PicturesRoot", pictures, "-MagickPath", executable, "-FakeMode", mode)
            before = source_state(source)
            result = run_owned(command, env=env, artifact_dir=work)
            after = source_state(source)
            outputs = inspect_outputs(pictures)
            trace = [json.loads(line) for line in (work / "native-calls.jsonl").read_text(encoding="utf-8").splitlines()]
            logs = "\n".join(p.read_text(encoding="utf-8-sig") for p in pictures.rglob("*.log"))
            record = {"shell": label, "case": mode, "native": "controlled_fake", "command": command, **result,
                      "source_before": before, "source_after": after, "source_preserved": before == after,
                      "unicode_destination_created": pictures.is_dir(), "outputs": outputs,
                      "fake_trace": trace, "application_log": logs}
            save_json(work / "execution.json", record)
            if result["exit_code"] or before != after or "SUMMARY " not in logs:
                raise RuntimeError(f"Incomplete capture or changed source: {label}/{mode}")
            records.append(record)
        for control in ["missing_marker", "outside_source", "overlap", "wrong_hash", "drive_root"]:
            work = root / "controls" / f"{label}-{control}"
            work.mkdir(parents=True)
            source = root / "corpus" / "weak"
            selected_scratch = root
            pictures = work / "pictures"
            extra = []
            if control == "missing_marker":
                selected_scratch = work
                source = work / "source"
                source.mkdir()
            elif control == "outside_source":
                source = REPO
            elif control == "overlap":
                pictures = source / "forbidden-output"
            elif control == "wrong_hash":
                extra = ["-ExpectedLegacySha256", "0" * 64]
            elif control == "drive_root":
                source = Path(root.anchor)
            command = shell_command(shell, HERE / "Invoke-LegacySnapshot.ps1",
                                    "-Repository", REPO, "-ScratchRoot", selected_scratch,
                                    "-Source", source, "-PicturesRoot", pictures, *extra)
            result = run_owned(command, timeout=15, artifact_dir=work)
            record = {"shell": label, "control": control, "command": command, **result, "output_absent": not pictures.exists()}
            save_json(work / "execution.json", record)
            if result["exit_code"] == 0 or pictures.exists():
                raise RuntimeError("Unsafe containment control")
            if control == "missing_marker" and "ownership marker" not in result["stderr"]:
                raise RuntimeError("Missing-marker control exercised a different guard")
            controls.append(record)
        after_inventory = run_owned(inventory_command)
        after_policy = json.loads(after_inventory["stdout"].lstrip("\ufeff"))["execution_policy"]
        if after_inventory["exit_code"] or after_policy != environment[label]["execution_policy"]:
            raise RuntimeError("Persistent policy check failed")
        environment[label]["persistent_policy_unchanged"] = True
    result = {"schema_version": 1, "scope": "harness validation and instrumented fake-process observations only",
              "legacy_sha256": sha256(REPO / "WinImgNormalizer.ps1"),
              "test_files_sha256": {p.relative_to(REPO).as_posix(): sha256(p) for p in sorted(HERE.iterdir()) if p.is_file()},
              "environment": environment, "executions": records, "containment_controls": controls,
              "containment": [json.loads(p.read_text(encoding="utf-8-sig")) for p in root.rglob("containment.json")]}
    save_json(root / "results.json", result)
    print(json.dumps({"fake_legacy_observations": len(records), "rejection_controls_passed": len(controls),
                      "unicode_destinations_tested": len(records), "real_codec_runs": 0}))


if __name__ == "__main__":
    main()
