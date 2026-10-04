# Instrumented legacy observations (M0-T02)

This development harness records the current application before behavioral fixes.
It does not make its known defects passing regression expectations. It does not
load the application into the test host or run the repository entry point against
the user's Pictures folder.

Use Windows, an existing explicitly selected ImageMagick 7 `magick.exe`, and
Python with Pillow. Dependencies are development-only. No script downloads or
installs tools. The recorded run pins exact executable/content hashes and versions;
repeat JPEG bytes are only meaningful within that toolchain. Future tool versions
must be recorded and independently inspected again.

```powershell
# Use your existing Python-with-Pillow and verified portable ImageMagick paths.
& $fixturePython -B tests/legacy/characterize.py `
  --root .scratch/legacy-new-run `
  --magick $verifiedMagick `
  --ps51 (Get-Command powershell).Source `
  --ps7 (Get-Command pwsh).Source
```

The root must be a **new child** of the ignored `.scratch` directory. Generated
fixtures, snapshots, raw paths and tool binaries remain there. `build_fixtures.py`
records recipes, tool versions, intended properties, independent Pillow decodes
and SHA-256 in `manifest.json`. All inputs are synthetic. Opaque video bytes test
copying, not video decoding. External colour profiles and absent optional codecs
must be reported without a support claim.

`check_harness.py --root .scratch/new-harness-run --ps51 <powershell.exe>
--ps7 <pwsh.exe>` runs the containment and controlled-process subset without
ImageMagick. Its three PNG inputs come from standard-library binary recipes and
are independently fully decoded by Pillow. The same case is run under both shells,
including Unicode output destinations and rejection of a real drive-root input.
It cannot close ordinary-image, collision or frame characterization. The proposed
official portable asset and digest are in `toolchain.json`; obtaining it requires
the owner's authorization. The owner approved it on 2026-10-04; the recorded build
was downloaded, hash-verified and extracted into ignored scratch without an installer
or permanent PATH change. An existing verified executable is also usable.

`Invoke-LegacySnapshot.ps1` requires a marked workspace, contained nonoverlapping
paths, no reparse points and the exact reviewed legacy hash. It replaces only the
legacy Pictures-resolution block in a disposable snapshot, records both blocks
and hashes, then runs that snapshot in another `-File` child process. Destination
names, discovery, conversion, duplicate bookkeeping and reporting remain legacy
code. These are **instrumented legacy** results; they do not prove the original
Known Folder resolution, `.bat` drag-and-drop or launcher exit propagation.

The controller bounds each worker to 60 seconds and terminates its owned process
tree on timeout. Temporary native files stay inside the owned snapshot directory.
The harness never recursively deletes artifacts. It verifies source hashes and
creation/modified times before and after each run, copied video hashes, complete
JPEG decode, and persistent PowerShell policy before/after.

Real ImageMagick observations cover ordinary images/video, competing output names
and GIF/TIFF frames. `fake_magick.py` is separately labeled: it models a failed
first duplicate and failed native commands that leave zero-byte or invalid output.
Its trace captures native arguments even though legacy code discards diagnostics.
Containment rejection controls must fail before any output directory is created.
Nested-output behavior uses fictional Windows path planning only; actual unsafe
recursion remains unexecuted.

Raw transcripts are `results.json`, `environment.json`, per-run `execution.json`,
logs and fake traces. Sanitized copies under `docs/codex-winimg/evidence` bind them
to reviewed source and test hashes. No private paths or generated media belong in
the tracked evidence. M0-T03 introduces application test seams and CI separately.
