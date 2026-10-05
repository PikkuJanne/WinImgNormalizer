# Project status

Updated: 2026-10-05 — M3-T01 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 66f609ad32637d82ce1e9918ea1600400378422e
Next task: **M3-T02 — Bound processing and own native process lifetime**
Task progress: **16 / 28 accepted**
Specified cases exercised: **T001-T050 (50 / 75); T001-T004 are characterization**
Pester suite: **287/287 passed in each desktop shell, zero skipped**
Failure controls: **288 total, 287 passed, one intended failure; native exit 1**
CI: **implementation push/PR Windows matrices passed, including controls**
Containing evidence checkpoint synchronization: **pending_verification**

## Current behavior

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

T047 — passed: Both actual fresh Windows hosts establish UnauthorizedAccess directory
enumeration using an explicit owned ListDirectory deny ACE. Accessible native JPEG/video
siblings complete; denied-only batch returns warning outcome; scan reports one
incomplete directory and unknown files remain uncounted. Applied source ACL is not
changed by the application, original DACL restored exactly in finally, and source
hashes/creation/modified/directory state match. The denied-only fixture explicitly
establishes a no-auto-inheritance original descriptor before the source baseline.
Original full requested mask7 owner/group/DACL descriptor bytes, DACL bytes, SDDL and
control flags are restored exactly using a fixture-only selected native DACL write,
while natural auto-inherited originals retain Set-Acl. Applied descriptor/control/DACL
hashes remain unchanged by the application before finally restoration.

T048 — passed: Actual owned junction loop and outside junction are skipped with
accessible native siblings, two skipped links and warning outcome; outside/source state
retained. Actual source-root junction alias fails setup before output creation. Optional
real file-symlink attempt is unavailable with Win32 1314 and is explicitly not actual
file-symlink coverage; existing controlled file-reparse tests remain separate.

T049 — passed: Controlled initial log creation and late SUMMARY/Processing-ended append
failures retain actual validated JPEG/video media and persist one degraded warning.
Exclusively held real log makes actual Add-Content append fail during delegated native
conversion. Bounded fallback retains quiet diagnostics, counts dropped/truncated lines,
disables failed sink after one attempt, emits once and catches unavailable emergency
stderr. Fresh-child PipelineStopped helper state proves propagation is not ordinary log
I/O degradation; no Ctrl+C lifecycle claim.

T050 — passed: Four actual-media runs independently inject LastWriteTimeUtc or
CreationTimeUtc setter failure on finalized image/video. Each attempts the other field,
retains valid data and exact video bytes, completes later siblings, reports one
timestamp warning/zero file errors, preserves warning duplicate status, retains sources
and empties owned work.

## Verification and handoff

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

Actual development results and their source/artifact qualifications are retained in
evidence/M3-T01.json and SESSION_LOG.md. Clean implementation desktop and hosted gates
are separate acceptance observations; earlier task histories retain their original
revisions.

Started clean at 5cbaa2246923f5bd94066feea34f71e0faafce00, equal to the live feature
branch; owner-merged PR #15 main 446683923f7d4b2aed7a5c9a00e1bf11d2924b68 had the same
tree. Continued without changing checkout/history. Successor draft PR #16 contains this
task. Implementation synchronization is a past exact-SHA observation. This record-only
checkpoint's SHA/live synchronization remains pending_verification until separately
observed after normal commit/push.

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

This session stops at M3-T01. Begin M3-T02 only when requested.
