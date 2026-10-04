# Next session handoff

Next task: **M1-T04 — Validate temporary outputs and finalize without overwriting**.
M1-T03 is complete. Begin only that next task in a fresh thread.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; the configured
WinImgNormalizer-main directory is the preserved non-Git snapshot. Read AGENTS.md,
STATUS.md, TASKS.json, GIT_WORKFLOW.md and tasks/M1-T04.md with its referenced cases.
Recheck clean state, canonical fetch/push URLs and the exact advertised feature SHA.
Tested runtime/tests implementation is da91c0d648b387b270bcdd7269a48ee1c63d11eb.
The final evidence checkpoint SHA/live sync observation is reported externally in
the preceding response and draft PR #6; verify it again at session start.

Get-WinImgOutputPlan returns the complete eligible image/video map in ordinal
relative-source order. Each row retains Source (FileInfo), SourceRelativePath,
OutputRelativePath, Kind and NamingReason. OrdinalIgnoreCase reservations include
mirrored directory paths and generated namespace/work/reports; unique legacy image
and video targets are reserved before collision suffixes. Colliding stems use
__lowercase_sourceext and checked __N starting at 2. Processing uses these exact rows;
the full PLAN log precedes native conversion and includes later skipped/failed rows.

Image candidates are absent image.jpeg paths under exclusively allocated GUID
directories in generated work. Existing conversion retry/size-only acceptance still
uses that candidate across attempts. Get-WinImgNativeOutputPath adds an internal
extended drive/UNC spelling only for long generated scratch passed to ImageMagick.
Managed paths/map names remain ordinary; public device paths remain unsupported.
Move-WinImgPlannedImage rechecks parent and occupancy, then uses File.Move(src,dest)
without replacement. Copy-WinImgPlannedVideo similarly uses File.Copy(src,dest,false).
Naming/allocation/finalization/cleanup failures return 2. Cleanup removes only the
known candidate and its empty exclusively owned directory, leaving unknown numbered
outputs for diagnosis. No complete transaction/content-validation claim is made.

Implement only M1-T04: validate each fresh attempt's successful native outcome,
nonzero JPEG bytes, positive dimensions, one frame and full pixel decode; prevent
stale candidate acceptance. Stage videos in owned partials, verify completion and
source stability, then finalize without replacement. Retain source preservation,
collision map, literal-safe filesystem operations and PS 5.1 / PS 7 compatibility.
Current duplicate registration/length heuristic and frame/colour/general lifecycle
defects are later tasks. Never normalize real Pictures, drive roots or private media.

Actual desktop Windows PS 5.1 / PS 7 passed 119/119 with zero skips; controls 120 total,
119 passed, one deliberate false assertion and exit 1. Real colliding JPG/PNG/BMP
outputs full-decode; source/video bytes and creation/modified timestamps pass.
Case-only source inventory and HEIC routing use mandatory controlled tests; no real
HEIC decoding or NTFS case-sensitivity changes. Native long scratch write/move is
tested, not universal long source/final-name support. Existing junction tests are
real; file-symlink creation was denied Win32 1314 with controlled coverage recorded.
CI push/PR Windows Server matrices both passed at the implementation SHA; exact
versions/counts/job links are in M1-T03-ci.json. No live UNC, real Pictures, manual
launcher, owner quality acceptance, merge, release or deployment is claimed.

Commands, separately in each fresh Windows shell:

```powershell
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
.\tests\Invoke-Tests.ps1 -DeliberateFailure # expected exit 1, one false assertion
```

Pinned bootstrap reuses verified cached archives or explicitly downloads development
bytes into ignored scratch, without installer, persistent PATH/policy changes or
runtime bootstrap. Children remove inherited PSModulePath; never redefine USERPROFILE.
Default runner includes four suites, including Naming.Tests.ps1. Keep raw results,
media and tools ignored. Retain initial development failures honestly in evidence.

Keep the feature branch and draft PR #6; if the owner merged it, inspect exact remote
and tree identity and create one successor draft without rewriting history. No
merge, tag/release, settings change, deployment or owner quality approval is authorized
by the handoff. Stop after M1-T03; start M1-T04 only in the next requested thread.
