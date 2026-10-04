# Next session handoff

Next task: **M2-T01 — Implement deliberate first-frame and first-page handling**.
M1-T06 is complete and the covered output-safety milestone is closed. Begin only
M2-T01 when next requested.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening. The configured
WinImgNormalizer-main directory is the preserved non-Git snapshot. Read AGENTS.md,
STATUS.md, TASKS.json, GIT_WORKFLOW.md, tasks/M2-T01.md and only its referenced
specifications/cases T029-T031. Recheck clean state, canonical fetch/push remote
identity and exact advertised feature SHA before editing.
Tested runtime/tests implementation: 4d7f9f2e7681902971b584560211139d555bf011.
The final evidence-checkpoint SHA/live sync is reported externally in the preceding
response and draft PR #9; independently verify it again at session start.

Covered M1 preflight/traversal/collision/finalization/duplicate behavior remains the
baseline: complete total ordinal output plan; OrdinalIgnoreCase name reservations;
fresh exclusive run and image/video attempt paths; successful native results and full
single-frame JPEG decode before saving; stable staged video copying; same-volume
no-overwrite File.Move; exact owned cleanup; input-length guarded success-based
duplicate keys with retained source/output/status logs. Original sources are preserved.
Length/time and duplicate checks remain heuristics; matching bytes are not guaranteed
and no runtime hashing requirement exists. D20-D23 document the design.

T028 verifies mixed synthetic source bytes and creation/modified timestamps plus
separate deterministic output runs, preserved first-run outputs/logs and existing
user entries. All mandatory combined suites passed under actual desktop PS 5.1 and
PS 7 at the exact clean implementation: 173/173, zero skips. Controls each
174 total/173 passed/one T007 false assertion/exit 1. Push/PR Windows Server
matrices passed; raw/Git binding, environments, artifacts and the complete milestone
diff review are in M1-T06.json and M1-T06-ci.json.
M1-T06 changes test coverage and test documentation; runtime and batch bytes are
unchanged from the starting checkpoint.

M2-T01 selects one deliberate first displayed animation frame and one first document
page. Respect canvas placement and orientation; never stack unrelated pages or depend
on implicit numbered JPEG output. Report original frame/page counts and omitted
content honestly. T029-T031 require exactly one intended JPEG and visible dropped
frames/pages, with codec-specific availability explicit. Preserve the established
safe output plan, full validation, no-overwrite finalization and source bytes/times.
Do not start colour/metadata, native lifecycle, general cancellation or publication
work in this task.

Commands, separately in each fresh Windows shell:

```powershell
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
.\tests\Invoke-Tests.ps1 -DeliberateFailure # expected exit 1, one false assertion
```

Pinned development bootstrap reuses verified cached archives or explicitly downloads
only to ignored scratch; no installer, runtime bootstrap or persistent PATH/policy
change. Children clear inherited PSModulePath; never redefine USERPROFILE. Keep raw
artifacts, tools and media ignored. Retain actual development failures honestly.

Source access times are excluded. Inherited controlled HEIC/case-only inventory and
source-change/cancellation paths remain distinct from actual separate-directory
casing and real junction checks. Denied file-symlink capability, lexical drive/UNC
roots, no real Pictures/live UNC, tested long generated scratch limits, no real
Ctrl+C/hard-kill, manual launcher, owner quality acceptance or full corpus/analyzer
closure remain explicit. Do not claim batch atomicity, crash durability or a hostile
filesystem sandbox.

Keep the feature branch and draft PR #9. If the owner merged it, inspect exact
remote/tree identity and create one successor draft without rewriting history.
Normal scoped feature commits/pushes remain authorized; merge, default-branch changes,
tags/releases, settings, deployment and owner approval are separate gates.
Stop after M1-T06; begin M2-T01 only when next requested.
