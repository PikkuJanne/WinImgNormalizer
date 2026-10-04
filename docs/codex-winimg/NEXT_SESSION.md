# Next session handoff

Next task: **M1-T03 — Plan collision-free output names before conversion**.
M1-T02 is complete. Begin only that next task in a fresh thread.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening. The configured
WinImgNormalizer-main folder remains the preserved non-Git snapshot. Read AGENTS.md,
STATUS.md, TASKS.json, GIT_WORKFLOW.md and tasks/M1-T03.md with referenced specs/cases.
Recheck clean state, canonical fetch/push URLs and exact advertised feature SHA.
Final checkpoint/sync is reported in the preceding response and draft PR #5;
tested runtime/tests revision is 8c9b157338419958a14c662ca379ccb8c50e8825.

M1-T02 rejects canonical destinations inside/equal to source before enumeration or
writes, including probes. OrdinalIgnoreCase whole-segment comparisons are deliberately
conservative on Windows. Existing-directory handles canonicalize paths; missing
output tails are appended after ancestor inspection. Linked source/output roots or
ancestors, dangling links, device paths and components ending in dots/spaces fail
setup. Never normalize real Pictures or a drive root.

Get-WinImgSourceTree is one nonrecursive inventory: directories, files, top-level
names and warnings. Reparse entries are skipped before descent/read; unreadable
subtrees are incomplete scans. Mirrors use this inventory. Source/output paths are
rechecked before processing. Scan/safety omissions return code 2 even without
eligible files. General conversion/copy exit classification remains later work;
this is not a hostile-filesystem sandbox.

New-WinImgRunDirectory uses exclusive native mkdir, timestamp/GUID suffix and eight
collision retries; existing objects are never adopted. Native calls add internal
extended drive/UNC prefixes so generated descendants can exceed MAX_PATH without
shortening names. The native 272-character check passes; it does not certify every
long source/native ImageMagick path across hosts.

Generated work/report/log paths use the first free `.WinImgNormalizer`, `__2`, etc.,
reserved against every observed top-level source name, including ignored files and
empty directories. Similar source directories mirror normally. The log is opened
CreateNew in `reports/WinImgNormalizer_<stamp>.log`. An arrival is preserved; failed
log setup stops with the owned run retained, without the old shared TEMP fallback.
`work` is reserved for later transactions.

Next implement only M1-T03: deterministic complete per-directory namespace planning,
stable extension/numeric suffixes, explicit source/output map, file/directory and
secondary/case collisions, and protection against final-target external arrivals.
Keep legacy names when unique. Current direct final image/video writes and duplicate
heuristic retain known later-task defects. Use owned synthetic tests T017-T019,
never real user media. Include the M1-T02 generated namespace in reservations.

Actual desktop PS5.1/PS7 passed 95/95 with zero skips; controls produced 96 total,
95 passed, one false assertion and exit 1. Real ordinary positional forms, JPEG full
decode, source/video hashes/timestamps, loop/outside/dangling junctions and simultaneous
same-stamp child runs pass. Desktop file symlink creation was denied Win32 1314;
mandatory controlled coverage is distinct. Drive/UNC controls are lexical; denied
scan failures use mocks. Actual CI results/versions are in M1-T02-ci.json. No live
UNC/real Pictures/manual launcher/owner visual-quality acceptance is claimed.

Initial c6157fb desktop passed 94/94, but push/PR PS7 CI failed six setup tests because
expanded generated paths hit the native mkdir limit. Owned reproduction and internal
extended-prefix correction produced 8c9b157. Initial infra/fixture/CI failures remain
in evidence rather than being relabeled.

Commands, separately in each fresh Windows shell:

```powershell
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
.\tests\Invoke-Tests.ps1 -DeliberateFailure # expected exit 1 with one false assertion
```

Pinned bootstrap reuses verified cached archives or explicitly downloads development
bytes into ignored scratch. No installer, persistent PATH/policy change or runtime
bootstrap. Child shells remove inherited PSModulePath; never redefine USERPROFILE.
Keep raw results, media and tools ignored. Evidence binds exact code/hashes. Verify
the containing checkpoint's live sync at the next session.

Keep the feature branch and draft PR #5; if the owner merged it, inspect exact
remote/tree identity and create one successor draft without rewriting history.
No merges, tags/releases, settings changes, deployment or subjective owner quality
approval are authorized by this handoff. Stop after M1-T02.
