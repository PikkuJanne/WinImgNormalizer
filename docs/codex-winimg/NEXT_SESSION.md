# Next session handoff

Next task: **M1-T02 — Protect traversal and allocate unique run directories**.
M1-T01 is complete. Begin only that next task in a fresh thread.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening. The configured
WinImgNormalizer-main folder remains the preserved non-Git snapshot. Read AGENTS.md,
STATUS.md, TASKS.json, GIT_WORKFLOW.md and tasks/M1-T02.md with referenced specs.
Recheck clean state, canonical fetch/push URLs and exact advertised feature SHA.
The final checkpoint/sync is reported in the preceding response and draft PR #4;
tested runtime/tests revision is f57b8adbb6859e8812471a4cc123b73530a2b301.

M1-T01 adds friendly argument/source/positive Int64 checks, root-aware paths,
application-only ImageMagick resolution and structured bounded preflight queries.
The public positional interface and batch launcher are retained. Minimum reviewed
ImageMagick is 7.1.2-32, with upstream sources in SOURCES I5; recheck notices during
future dependency work. MagickPath/PreflightRunner are internal test seams, alongside
OutputParent and conversion-only ProcessRunner. Compiled format modes include
optional module columns; HEIF/HEIC are distinct read coders. Missing decoder files
are skipped with explicit errors, supported media continues and code 2 is returned.
Video-only/empty batches require ImageMagick but no JPEG writer unless readable
images exist. Actual broad image-error/batch exit behavior remains later scope.

A DeleteOnClose exclusive probe checks destination write/delete rights. The space
estimate uses copied videos, min(cap, max(1 MiB, 4 times input)) per readable image,
twice largest readable input and 64 MiB reserve; missing decoders use no conversion
space. Unknown capacity warns, measured shortage fails setup. Directory-create
failure returns friendly code 1. Large divisible cap quotients are exact Int64;
legacy extent units/default conversion sequence are retained for later M2-T04.

Next implement only M1-T02: segment-aware canonical containment, reject destination
inside/equal source before any writes (including probes), protect traversal and
allocate exclusive distinct run roots/generated namespaces. Existing timestamp-only
run creation still reuses directories, and the mirror still has unsafe nesting;
do not execute those cases against real Pictures, user data or drive roots.
Use contained disposable/mocked tests for T013-T016. Keep the smallest patch.

Actual desktop PS5.1/PS7 passed 66/66 with zero skips and expected deliberate failure
controls (67 total, 66 passed, one false assertion, exit1). Real -File ordinary forms
and setup failures, executable shadow controls, JPEG full decode and video/source
hash/timestamp preservation pass. Both push/PR Server CI matrices pass on the same
implementation; actual CI PS versions differ and are in M1-T01-ci.json. Live UNC,
real Pictures/manual launcher/ACL exhaustion and real HEIC/HEIF decode are not claimed.

Commands, separately in each fresh Windows shell:

```powershell
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
.\tests\Invoke-Tests.ps1 -DeliberateFailure # expected exit 1 with one false assertion
```

Bootstrap reuses verified cached archives or explicitly downloads pinned development
bytes into ignored scratch, with fresh extraction and hashes. No installer/persistent
PATH/policy change; runtime never bootstraps. Child shells remove inherited PSModulePath;
never redefine USERPROFILE or run normalization against real Pictures.

Evidence: M1-T01.json and M1-T01-ci.json, bound to committed code/hashes. Raw artifacts
remain ignored. Final containing checkpoint sync needs live verification next session.
Keep the feature branch and draft PR #4; if the owner has merged it, inspect exact
remote/tree identity and create one successor draft without rewriting local history.
No merges, tags/releases, settings changes, website deployment or subjective owner
quality approval are authorized by this handoff. Stop after M1-T01.
