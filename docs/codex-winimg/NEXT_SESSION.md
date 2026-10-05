# Next session handoff

Next task: **M3-T05 — Propagate exit codes and harden the batch launcher**. M3-T04 is complete.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; WinImgNormalizer-main remains the preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json, GIT_WORKFLOW.md, tasks/M3-T05.md and its referenced specs/cases T061-T063. Recheck clean state, canonical fetch/push identity and the exact advertised feature SHA before editing. Tested implementation is 158315f1e48e6199c5c18d3e323b5945390789b0; verify the final evidence-checkpoint SHA and draft PR #19 independently.

Discovery creates one terminal outcome per successfully inspected regular file,
including Ignored unsupported files; inaccessible directories, uninspectable entries and
skipped links remain separate scan issues with unknown file totals.

Validated no-overwrite image and stable video final moves commit their retained outcomes
before ancillary operations; timestamp, cleanup, logging and observer warnings do not
double-count files as errors.

Exact Decimal aggregates pair only finalized image input/output lengths, allow signed
negative savings and keep video bytes separate; elapsed duration and finalized-file
throughput are observed values.

On actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1, each normal suite passed 348/348 with native exit 0 and zero skipped or
incomplete tests. Each control had 349 total, 348 passed, exactly one T007 deliberate
failure and native exit 1. Summary/XML, parent native records and console hashes bind 31
tested source paths, including the reporting suite and all eight colour/reference
assets, to observed raw bytes and separate Git blobs. Only verified text CRLF-to-LF
normalization is allowed; ICC bytes remain exact. Persistent execution policies were
unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37311855781); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37311860891).
Evidence: evidence/M3-T04.json and evidence/M3-T04-ci.json.

M3-T05 propagates the documented script exit codes through the batch launcher, selects honest success/warning/error/cancelled messages, preserves the familiar pause, and rejects extra dropped folders without losing literal arguments. Verify actual Windows launch behavior and unchanged persistent security settings. Preserve the accepted reporting/accounting, source/no-overwrite safety, completed media, byte-identical videos and native resource/cancellation policy.

Actual Import-Csv and independent reversible-decoder checks establish the documented
machine import contract, not actual Excel/LibreOffice execution or consumer
editing/save/reopen safety. The spreadsheet display prefix must remain intact.

Unknown files inside inaccessible locations are not invented. An incomplete or failed
CSV must not be treated as a complete inventory; valid media remain retained.

Cooperative application130 and best-effort reporting are separate from host closure,
force termination and stopped pipelines. Those boundaries can bypass reporting/cleanup;
synchronous filesystem work remains nonpreemptible.

Native lifetime, source-stability and duplicate matching retain their documented
practical/heuristic boundaries; reporting does not create content identity, mandatory
hashing, a hostile-filesystem sandbox or an unlimited Decimal aggregate.

The BAT and JPEG/colour/frame/size defaults are unchanged in this card. M3-T05 launcher
work, M3-T06/M3-T07 and owner quality/merge/release acceptance remain separate gates.

Fresh children may use process-only Bypass and clear inherited PSModulePath; never change persistent PATH/policy or redefine USERPROFILE. Keep raw paths, media, argv and tools ignored. Preserve actual failures. Use one successor draft if the owner merged this PR; normal scoped feature commits/pushes remain authorized. Merges, tags/releases, settings/deployment and owner acceptance remain separate gates. This session stops at M3-T04. Once requested, complete only M3-T05 and stop.
