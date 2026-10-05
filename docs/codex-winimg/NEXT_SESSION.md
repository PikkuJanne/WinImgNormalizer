# Next session handoff

Next task: **M3-T04 — Reconcile counts and add useful safe reports**. M3-T03 is complete.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; WinImgNormalizer-main remains the preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json, GIT_WORKFLOW.md, tasks/M3-T04.md and its referenced specs/cases T057-T060. Recheck clean state, canonical fetch/push identity and the exact advertised feature SHA before editing. Tested implementation is 1501eb2e52ee46a14f438dd240e35ed6da99aa7f; verify the final evidence-checkpoint SHA and draft PR #18 independently.

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

M3-T04 reconciles discovered file outcomes and adds useful safe local reports. Count ignored files at discovery, keep inaccessible directories separate, and partition converted/copied/duplicate/error/cancelled outcomes without double counting warning flags. Add duration, image-only byte accounting, size warnings and retained mappings; neutralize formula-leading CSV fields and report output/log failures. Preserve public positional forms, source/no-overwrite safety, finalized outputs, byte-identical videos and accepted resource/cancellation behavior.

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

Fresh children may use process-only Bypass and clear inherited PSModulePath; never change persistent PATH/policy or redefine USERPROFILE. Keep raw paths, media, argv and tools ignored. Preserve actual failures. Use one successor draft if the owner merged this PR; normal scoped feature commits/pushes remain authorized. Merges, tags/releases, settings/deployment and owner acceptance remain separate gates. This session stops at M3-T03. Once requested, complete only M3-T04 and stop.
