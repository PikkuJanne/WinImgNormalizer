# Project status

Updated: 2026-10-05 — M3-T03 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 1501eb2e52ee46a14f438dd240e35ed6da99aa7f
Next task: **M3-T04 — Reconcile counts and add useful safe reports**
Task progress: **18 / 28 accepted**
Specified cases exercised: **T001-T056 (56 / 75); T001-T004 are characterization**
Pester suite: **316/316 passed in each desktop shell, zero skipped**
Failure controls: **317 total, 316 passed, one intended failure; native exit 1**
CI: **implementation push/PR Windows matrices passed, including controls**
Containing evidence checkpoint synchronization: **pending_verification**

## Current behavior

A scoped per-run C# state receives actual CTRL_C_EVENT without executing PowerShell,
reporting or cleanup on the Windows control thread. Command execution enables console
capture; callable tests explicitly opt in or supply controlled request state. Session
disposal restores the prior ambient state and removes or safely deactivates its rooted
handler.

Native launch keeps the existing six-argument API, private atomic Windows job and
bounded readers. Cancellation is checked before launch/resume and at native waits of at
most 50 milliseconds without restarting the deadline; the owned job and its descendants
are terminated through the existing finite cleanup grace. Cancellation cannot become a
timeout retry or authorize an image candidate.

Image snapshots and video partials copy through 256 KiB chunks with request checks
around reads/writes and flush. The application stops taking new items and interrupts
retry backoff. Synchronous I/O itself is not preempted.

T054 — passed: Controlled cancellation at
inspection/conversion/validation/precommit/after-finalization boundaries and actual
isolated console image observations on PS5.1/PS7/BAT retain completed media, stop
unstarted items, reject unfinished finals and produce application130 with bounded
interrupted reporting. Image event uses genuine conversion followed by owned native
pacing, not a claim about the codec compute interval.

T055 — passed: Actual video source copy reaches a real 256 KiB chunk then actual
private-console Ctrl+C on PS5.1/PS7/BAT; partial is not finalized, owned work is
cleaned, completed JPEG/source/hash/time/directory/sentinel state preserved. Controlled
foreign-neighbor arrival additionally verifies exact ownership cleanup.

T056 — passed: Actual owned-host PID/start-time force termination leaves recognizable
nonfinal scratch, retains prior completed media and source state, releases its private
native tree and leaves unrelated process alive. Distinct fresh normal run proves prior
run unchanged in core Pester; no force-cleanup or130 promise.

## Verification and handoff

On actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1, each normal suite passed 316/316 with native exit 0 and zero skipped or
incomplete tests. Each control had 317 total, 316 passed, exactly one T007 deliberate
failure and native exit 1. Summary/XML, parent native records and console hashes bind 30
tested source paths, including the cancellation suite and all eight colour/reference
assets, to observed raw bytes and separate Git blobs. Only verified text CRLF-to-LF
normalization is allowed; ICC bytes remain exact. Persistent execution policies were
unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37300472643); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37300498704).
Evidence: evidence/M3-T03.json and evidence/M3-T03-ci.json.

Actual PS5.1/PS7 console, unchanged BAT and forced-exit outcomes are retained with hashed artifacts. Public command observations and event-controlled pacing/routing are distinguished from controlled cancellation seams. Host/BAT force interruption is not assigned an unconditional cooperative exit code or cleanup guarantee.

Actual development results and their source/artifact qualifications are retained in
evidence/M3-T03.json and SESSION_LOG.md. Clean implementation desktop and hosted gates
are separate acceptance observations; earlier task histories retain their original
revisions.

Started clean at ac0fd08f8abc55ec195efc7c942338d625a04379, equal to the live feature
branch; owner-merged PR #17 main 27937f2a9ed84001aaedfe48523873fa02410c58 had the same
tree. Continued without changing checkout/history. Successor draft PR #18 contains this
task. Implementation synchronization is a past exact-SHA observation. This record-only
checkpoint's SHA/live synchronization remains pending_verification until separately
observed after normal commit/push.

Exit 130 and the interrupted record belong to cancellation handled by the active
application session. A request before handler registration, absent console, closed sink,
stopped host pipeline, terminal closure or force termination may bypass
cleanup/reporting or produce a different native host exit. No universal
Ctrl+C/force-exit code or finally guarantee is claimed.

Native polling and chunk boundaries do not bound a blocked filesystem or kernel I/O
call. Owned job termination and stream completion use the inherited finite grace; an
unconfirmed tree retains its affected cache and staging instead of deleting data a
remaining process may use.

Only known owned incomplete paths are cleaned. Unknown neighboring files and output
arrivals remain. Private allocation and cache naming provide practical ownership;
hostile concurrent filesystem replacement or indistinguishable same-pattern arrivals are
not a security sandbox.

Actual console delivery uses disposable private Windows consoles with verified
membership and group-zero CTRL_C_EVENT. Controlled phase routing and video pacing are
qualified separately. Safe -File drivers invoke the unchanged production Command with
explicit ignored output; the exact BAT copy uses that paired driver. This does not
certify the unmodified public script's automatic real Pictures path or physical
keyboard/Explorer interaction.

The BAT's existing Done message, pause and exit propagation are unchanged. Its actual
native outcome is observed separately from the application's interrupted summary; M3-T05
remains responsible for launcher hardening.

Forced worker evidence establishes the observed exact owned PID/start identity exit,
retained completed media and recognizable nonfinal scratch, with a fresh run that does
not adopt it. It does not promise universal cleanup or a successful summary after host
death.

This completes only M3-T03. Structured per-file reporting remains M3-T04, and owner
quality/default acceptance, merge, release and deployment are separate authorized gates.

This session stops at M3-T03. Begin M3-T04 only when requested.
