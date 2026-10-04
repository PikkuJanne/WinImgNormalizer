# Next session handoff

Next task: **M1-T05 — Make heuristic duplicate handling deterministic and success-based**.
M1-T04 is complete. Begin only that next task in a fresh requested thread.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; the configured
WinImgNormalizer-main directory is the preserved non-Git snapshot. Read AGENTS.md,
STATUS.md, TASKS.json, GIT_WORKFLOW.md and tasks/M1-T05.md with its referenced cases.
Recheck clean state, canonical fetch/push URLs and the exact advertised feature SHA.
Tested runtime/tests implementation is ed7fb7691eca9dd55082338230657804c1865681.
The final evidence checkpoint SHA/live sync observation is reported externally in
the preceding response and draft PR #7; verify it again at session start.

Get-WinImgOutputPlan already returns the complete eligible image/video map in total
ordinal relative-source order. Processing follows those rows and PLAN log lines
precede conversion. OrdinalIgnoreCase reservations and familiar unique names plus
extension/numeric collision suffixes remain unchanged.

Convert-ImageMagick now reserves each attempt with FileMode.CreateNew, accepts only
an unambiguous successful native outcome, and calls Test-WinImgImageCandidate for
nonempty regular bytes plus full single-frame JPEG decoding and positive dimensions.
Validation checks measured length/modification time before and after decoding. Every
fallback/scale receives a new exact candidate; superseded attempts are cleaned before
the next. Only the actual successful candidate and settings can be finalized. A valid
last-scale above-cap result retains the existing best-effort note.

Copy-WinImgPlannedVideo reserves video.partial in owned work, streams with source
FileShare.Read and destination FileShare.None, flushes/closes, then verifies partial
length plus source length/modification time. Move-WinImgPlannedImage checks same-volume
roots, ancestors and final availability before no-overwrite File.Move(src,dest).
Remove-WinImgOwnedCandidate deletes only explicitly owned exact files and empty owned
directories; unknown siblings and reservation conflicts survive. Item errors return 2.

The duplicate block still calls seen.Add on filename plus modification-time ticks
before validation or finalization and records only first-seen source. Implement only
M1-T05: deterministic success-based retained keys, an equal-length guard before skip,
and retained successful source/output mapping for each skip. T025-T027 require a good
second candidate after failed-first, different-length same-key processing and stable
retained links. Keep the documented heuristic limitation for same-key/same-length
different bytes. No mandatory runtime content hashing or original removal.

Desktop Windows PS 5.1 / PS 7 passed 151/151 with zero skips; controls each
152 total, 151 passed, one deliberate false assertion and exit 1. Actual full
JPEG decode/truncated-header and stable video hash tests pass. Source-change/partial/
interruption cases use controlled seams; no real Ctrl+C/hard-kill result is claimed.
Case-only source inventory and HEIC routing remain controlled tests. Real junction
checks remain distinct from denied file-symlink creation (Win32 1314). Native long
generated scratch is tested, not universal long source/final-name support. CI push/PR
Server matrices passed at the implementation SHA; versions/counts/job links are in
M1-T04-ci.json. No live UNC, real Pictures, private media, manual launcher, owner
quality acceptance, merge, release or deployment is claimed.

Commands, separately in each fresh Windows shell:

```powershell
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
.\tests\Invoke-Tests.ps1 -DeliberateFailure # expected exit 1, one false assertion
```

Pinned bootstrap reuses verified cached archives or explicitly downloads development
bytes into ignored scratch without installer, persistent PATH/policy changes or
runtime bootstrap. Children remove inherited PSModulePath; never redefine USERPROFILE.
The default runner includes five suites, including Transactions.Tests.ps1. Keep raw
results, media and tools ignored; retain development failures honestly in evidence.

Keep the feature branch and draft PR #7; if the owner merged it, inspect exact
remote and tree identity and create one successor draft without rewriting history.
No merge, tag/release, settings change, deployment or owner quality approval is
authorized by the handoff. Stop after M1-T04; start M1-T05 only when next requested.
