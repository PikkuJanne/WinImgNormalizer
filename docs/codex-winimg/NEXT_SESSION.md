# Next session handoff

Next task: **M1-T06 — Close the output-safety milestone**.
M1-T05 is complete. Begin only that next task in a fresh requested thread.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; the configured
WinImgNormalizer-main directory is the preserved non-Git snapshot. Read AGENTS.md,
STATUS.md, TASKS.json, GIT_WORKFLOW.md and tasks/M1-T06.md with its referenced cases.
Recheck clean state, canonical fetch/push URLs and the exact advertised feature SHA.
Tested runtime/tests implementation is c9988e204d25bbb13616b6df601d0f5c400a73c1.
The final evidence checkpoint SHA/live sync observation is reported externally in
the preceding response and draft PR #8; verify it again at session start.

Get-WinImgOutputPlan already returns the complete eligible image/video map in total
ordinal relative-source order, with OrdinalIgnoreCase reservations. Processing follows
those rows; PLAN log lines precede conversion. Complete path containment/inventory,
exclusive run allocation, naming and no-overwrite finalization remain established.

M1-T04 reserved fresh neutral image attempts, requires successful native outcomes and
full single-frame JPEG decoding, and prevents stale earlier results from being saved
after later failure. Video copies stream into owned partials and verify length plus
source length/modification time after close. Same-volume File.Move refuses final
replacement. Cleanup removes only explicitly owned exact files and empty owned
directories; unknown siblings and reservation conflicts survive. Item errors return 2.

M1-T05 removes early duplicate-key registration. Get-WinImgDuplicateKey combines
invariant lowercase filename, invariant LastWriteTimeUtc ticks and input length in an
Ordinal dictionary, keeping a separate first success for every length. Retained
mappings are added only after verified finalization and hold SourceRelativePath,
OutputRelativePath and Status; every skip logs all three. A failed-first image/video cannot
poison the retained key. Different lengths are processed separately. Valid above-cap
finalized images may register ConvertedWithWarning. Live regular-file metadata is
read before lookup; image registration rechecks length/time after finalization and
keeps changed-source output with a partial warning without registering it. Videos
register the stable copy helper's measured length/time snapshot. Same-key/same-length different
content can still match falsely; this remains a documented lightweight heuristic,
not content identity. No runtime hashing or original deletion was introduced.

Implement only M1-T06: run combined mandatory M1 cases, inspect every collision and
finalization path, verify synthetic source bytes plus creation/modified times and
two-run isolation, then review the full milestone diff for scope or dependency
changes. T028 is the source-preservation and safety-regression closure. Actual Windows
PS 5.1 and PS 7 results remain mandatory; missing target results cannot be called full
milestone acceptance. Record remaining native/frame/colour/lifecycle limits honestly.

Desktop Windows PS 5.1 / PS 7 passed 172/172 with zero skips; controls each
173 total, 172 passed, one deliberate false assertion and exit 1. Exact raw
checkout and committed blob binding is in M1-T05.json, including Git line endings.
CI push/PR Server matrices passed at the implementation SHA; versions/counts/job
links are in M1-T05-ci.json. Source hashes/timestamps and video identity remain
synthetic regression evidence. Inherited case-only single-directory inventory/HEIC
routing and source-change/cancellation outcomes remain controlled where documented.
New duplicate case variants use actual separate directories. Real junction checks are distinct from
denied file-symlink creation (Win32 1314); native long generated scratch does not
prove universal long source/final-name support. No live UNC, real Pictures, private
media, real Ctrl+C/hard-kill, manual launcher, owner quality acceptance, merge,
release or deployment is claimed.

Commands, separately in each fresh Windows shell:

```powershell
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
.\tests\Invoke-Tests.ps1 -DeliberateFailure # expected exit 1, one false assertion
```

Pinned bootstrap reuses verified cached archives or explicitly downloads development
bytes into ignored scratch without installer, persistent PATH/policy changes or
runtime bootstrap. Children remove inherited PSModulePath; never redefine USERPROFILE.
The default runner includes six suites, including Duplicates.Tests.ps1. Keep raw
results, media and tools ignored; retain development failures honestly in evidence.

Keep the feature branch and draft PR #8; if the owner merged it, inspect exact
remote and tree identity and create one successor draft without rewriting history.
No merge, tag/release, settings change, deployment or owner quality approval is
authorized by the handoff. Stop after M1-T05; start M1-T06 only when next requested.
