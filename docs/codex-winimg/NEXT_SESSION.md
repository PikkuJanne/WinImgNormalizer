# Next session handoff

Next task: **M3-T02 — Bound processing and own native process lifetime**. M3-T01 is complete.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; WinImgNormalizer-main remains the preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json, GIT_WORKFLOW.md, tasks/M3-T02.md and its referenced specs/cases T051-T053. Recheck clean state, canonical fetch/push identity and the exact advertised feature SHA before editing. Tested implementation is 66f609ad32637d82ce1e9918ea1600400378422e; verify the final evidence-checkpoint SHA and draft PR #16 independently.

A denied folder or uninspectable source entry leaves the scan incomplete while readable
siblings continue. The report separates inaccessible-directory, uninspectable-entry and
skipped-link counts; it does not invent a file count for unreadable contents.
Directory/file reparse points are skipped and reported without following loops or
outside targets. These omissions return application code 2, even when no eligible media
was readable; linked roots and ancestors remain setup rejections.

Log creation or append failure disables the failed disk sink once and marks reporting
degraded while valid media work continues. A visible console/stderr fallback retains at
most 8,192 UTF-16 characters, limits each stored line to 1,024 characters, escapes
control characters and reports dropped/truncated lines. Final reporting exposes
LogWarnings and DiskLogIncomplete, including failure on the last required log write, and
returns application code 2.

Creation and modified timestamps are restored independently from captured source
metadata after JPEG/video finalization. A failed field is named, valid finalized bytes
stay retained, and TimestampWarnings counts the affected output once within completed
image/video totals. Processing continues, application code 2 exposes the warning, and a
later heuristic duplicate links to the retained warning status. Public arguments, the
default byte cap, scale/JPEG/colour/frame policies and byte-identical video copying are
unchanged.

On actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1, each normal suite passed 287/287 with native exit 0 and zero skipped or
incomplete tests. Each control had 288 total, 287 passed, exactly one T007 deliberate
failure and native exit 1. Summary/XML, parent native records and console hashes bind 28
tested source paths, including the recovery suite and all eight colour/reference assets,
to observed raw bytes and separate Git blobs. Only verified text CRLF-to-LF
normalization is allowed; ICC bytes remain exact. Persistent execution policies were
unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37277567021); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37277571157).
Evidence: evidence/M3-T01.json and evidence/M3-T01-ci.json.

M3-T02 bounds individual native processing, applies process-local resource limits that respect stricter installed policies, and owns the lifetime of only the native process tree it starts. Preserve source/no-overwrite safety, subtree/link reporting, visible ancillary warnings, public positional forms, the default 1,048,576-byte cap, scale/JPEG settings and byte-identical video copying. Check current upstream security notices as required by that card.

ACL and link observations apply to the actual owned Windows fixtures and available local
capabilities. Controlled log-sink and timestamp-setter failures must be identified
separately from real permission or media failures; neither kind of probe establishes
every filesystem or remote provider. Root/ancestor link rejection remains conservative,
and concurrent hostile filesystem changes remain outside the isolation guarantee.

The fallback is bounded and may drop lines; it cannot reconstruct a complete disk log.
Direct stderr is best effort when the console is absent/closed or the host pipeline has
stopped. Explicit PipelineStoppedException propagation and helper-state observations do
not establish actual Ctrl+C, a cancellation exit contract, an I/O deadline, or native
descendant termination. Per-file bounds/process lifetime, cooperative cancellation and
complete report/CSV accounting remain later cards.

Original bytes and captured creation/modified metadata remain the preservation targets;
OS-maintained access time is not guaranteed. Existing pinned codec, colour and
lossy-JPEG qualifications remain in prior evidence. This card does not tune conversion
defaults or repeat the size/quality benchmark, and automated regression acceptance does
not grant owner approval, release or deployment permission.

Fresh children may use process-only Bypass and clear inherited PSModulePath; never change persistent PATH/policy or redefine USERPROFILE. Keep raw paths, media, argv and tools ignored. Preserve actual failures. Use one successor draft if the owner merged this PR; normal scoped feature commits/pushes remain authorized. Merges, tags/releases, settings/deployment and owner acceptance remain separate gates. This session stops at M3-T01. Once requested, complete only M3-T02 and stop.
