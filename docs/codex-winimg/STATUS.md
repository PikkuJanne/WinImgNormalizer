# Project status

Updated: 2026-10-04 — M0-T02 characterization completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Reviewed continuation HEAD: 38f1a33f753600fdc4e70af3dc68a323a65c19ae
Next task: **M0-T03 — Introduce minimal test seams and early Windows checks**
Task progress: **2 / 28 accepted**
Characterization cases: **T001-T004 completed (4 / 75 cases)**
Corrected-behavior regression suite: **not implemented or claimed by M0-T02**
Real instrumented legacy runs: **6**; controlled fake-process runs: **6**
GitHub sync for this containing checkpoint: **pending_verification**
CI: **unconfigured; no passing CI claimed**
Owner quality acceptance: **not requested / not granted**

## Completed characterization

The owner approved the pinned portable ImageMagick download and continuation.
Official ImageMagick 7.1.2-32 Q16 x64 archive matched the recorded SHA-256 and size before
extraction. It ran only from ignored scratch, without an installer, permanent PATH
change or machine-wide policy change. Tool provenance is in tests/legacy/toolchain.json.

Built 45 synthetic fixtures with independent complete Pillow inspection, explicit
recipes/provenance, source SHA-256 and actual Windows creation/modified timestamps.
The independent audit passed 45/45. JPEG/PNG/BMP/TIFF, offset GIF, animated WebP,
EXIF 1-8, alpha, legal filename, benign-invalid and opaque-video controls are present.
Advanced fixtures are inspected inputs; they were not all passed through the app.
HEIC/HEIF, trusted external profiles and remaining environment-failure corpus remain
explicitly deferred to their owning tasks, without a broad format-support claim.

Windows PowerShell 5.1.26100.9444 and PowerShell 7.6.5 each ran ordinary, collision,
frame and three fake-process scenarios: twelve observations total, all captured.
The controller passed 22 harness checks (12 complete captures/source preservation,
8 containment rejections, 2 persistent-policy checks). Earlier standalone harness
checks also cover Unicode destinations and drive-root rejection.

Ordinary mixed trees produced 4 fully decodable JPEGs and copied 3 videos with
exact hashes and mapped timestamps; empty directories mirrored. Logs and all
source bytes/creation/modified times are recorded. Known defects remain separate:

- Collision: JPG/TIFF overwrite one same-stem JPEG; an existing .jpeg directory is
  logged OK [1 bytes]; another PNG is duplicate-skipped after the directory case.
  Actual summary Converted=4 / Duplicates=1 / Errors=0 retains only 2 JPEG files.
- GIF/TIFF/WebP: 6 numbered JPEGs remain at the last 50% retry, while intended base
  filenames are absent; summary Converted=0 / Errors=3, host exit 0 in each shell.
- Fake failed-first: 12 attempts against a, no conversion of b, duplicate skip.
- Fake weak outputs: zero/23-byte undecodable files count as converted although
  native failure controls return 9; host exits 0. No corrected success is claimed.

All twelve snapshots differ from the original only by the Pictures substitution.
The tracked application, launcher and assets are unchanged. No dot-sourcing, real
Pictures write, source/profile mutation or unsafe nested recursion ran. Fictional
nested planning is recorded. No remaining M0-T02 acceptance blocker.

## Evidence and next action

Evidence: M0-T02.json, M0-T02-real-observations.json and M0-T02-fixtures.json.
The previous partial transcript remains M0-T02-observations.json; its pending-tool
state describes the earlier checkpoint only. Raw artifacts remain ignored.
The initial GIF generator failure was a recipe-scoping error, fixed without
weakening assertions. WebP detection now recognizes the current native listing.

The one-commit checkpoint binds final tested source and evidence by precommit
hashes; actual final push/remote comparison is reported externally after commit.
At continuation start, clean feature HEAD independently equaled remote 38f1a33;
main remained 2a51c0b with the prior merge and no application delta. Draft PR #2
continues this series. Start M0-T03 only in a subsequent thread. Publication gates
and subjective owner acceptance remain separate; no CI or release readiness implied.
