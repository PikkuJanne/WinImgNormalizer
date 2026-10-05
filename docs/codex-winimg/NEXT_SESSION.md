# Next session handoff

Next task: **M2-T06 — Close the conversion-correctness milestone**. M2-T05 is complete.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; WinImgNormalizer-main remains the preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json, GIT_WORKFLOW.md, tasks/M2-T06.md and its referenced specs/cases, including T046. Recheck clean state, canonical fetch/push identity and the exact advertised feature SHA before editing. Tested implementation is 5ddbf4156f2c10534a045586a5b5e5e050255c30; verify the final evidence-checkpoint SHA from the prior response and draft PR #14 independently.

M2-T05 now rejects failed, unknown or truncated native diagnostics, permits one
validated known warning, and retries only explicit sharing/lock failures twice per
image. Review STATUS.md and evidence/M2-T05.json for the concrete policy, actual warning
provenance, retained failures and exact source bindings.

Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1: each normal run passed 269/269 with native exit 0 and zero skipped or incomplete
tests. Each control had 270 total, 269 passed, one exact T007 deliberate failure and
native exit 1. Summary/XML, parent native records and console hashes bind 26 tested
source paths to raw checkout and separate Git blobs, with only verified CRLF-to-LF
normalization where needed. Persistent execution policies were unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37267178588); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37267182413).
Evidence: evidence/M2-T05.json and evidence/M2-T05-ci.json.

M2-T06 closes the conversion-correctness milestone. Run representative colour, frame/page, literal-name, orientation and size cases together, bind comparison artifacts to the tested revision, and retain actual codec/build qualifications. JPEG derivatives are lossy and do not replace preserved originals. Preserve the current public forms, exact byte cap, scale/JPEG settings, source/no-overwrite safety, colour/white-alpha/strip order and byte-identical video copying.

General per-file deadlines, cancellation, native descendant lifecycle and final
reporting/BAT exit propagation remain later M3 gates. Fixed character capture is not a
CPU/disk/lifetime sandbox. Controlled warning, retry and process-local resource fixtures
remain qualified; prior codec/colour/provider limits and owner quality/default approval
are unchanged. Full limitations and raw artifact bindings are in evidence/M2-T05.json.

Fresh children may use process-only Bypass and clear inherited PSModulePath; never change persistent PATH/policy or redefine USERPROFILE. Keep raw paths, media, argv and tools ignored. Preserve actual failures. Use one successor draft if the owner merged this PR; normal scoped feature commits/pushes remain authorized. Merges, tags/releases, settings/deployment and owner acceptance remain separate gates. This session stops at M2-T05. Once requested, complete only M2-T06 and stop.
