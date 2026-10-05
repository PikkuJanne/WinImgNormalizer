# Next session handoff

Next task: **M3-T01 — Recover from inaccessible folders and expose ancillary failures**. M2-T06 and the M2 conversion-correctness milestone are complete.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; WinImgNormalizer-main remains the preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json, GIT_WORKFLOW.md, tasks/M3-T01.md and its referenced specs/cases T047-T050. Recheck clean state, canonical fetch/push identity and the exact advertised feature SHA before editing. Tested implementation is 11ee247154a00d79aaa49a438d1f526457b22fdf; verify the final evidence-checkpoint SHA from the prior response and draft PR #15 independently.

T046 exercises colour handling, frame/page selection, literal names, orientation and
exact byte-cap sizing together through the actual pinned native runner on approved
synthetic inputs. The combined checks preserve source bytes and recorded
creation/modified times, byte-identical video copying, deterministic output mapping and
earlier outputs across repeated runs. Concrete fixture/cap counts and outcomes are
recorded below rather than inferred from the prior task.

Each recorded mixed-tree batch retained 12 fully decoded JPEGs and 1 byte-identical
video, skipped 1 heuristic duplicate, and returned application code 0. The omitted
default cap was 1,048,576 bytes; the explicit exact cap was 65,537 bytes. Source state
and the complete earlier output/log tree stayed unchanged. Focused development and
immutable-I observations remain separately labeled in the comparison evidence.

Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1: each normal run passed 271/271 with native exit 0 and zero skipped or incomplete
tests. Each control had 272 total, 271 passed, one exact T007 deliberate failure and
native exit 1. Summary/XML, parent native records and console hashes bind 27 tested
source paths, including the integrated suite and all eight independent colour/reference
assets, to raw checkout and separate Git blobs. Only verified text CRLF-to-LF
normalization is allowed; ICC bytes match exactly. Persistent execution policies were
unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37272249221); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37272253545).
Evidence: evidence/M2-T06.json and evidence/M2-T06-ci.json.

M3-T01 recovers accessible siblings when a subtree is denied, applies skip-and-report to links/reparse points and makes log/timestamp failures visible. Preserve source/no-overwrite safety and the current conversion policies, public positional forms, default 1,048,576-byte cap, scale/JPEG settings and byte-identical video copying. The integrated comparison recipes and native/colour/quality qualifications are in evidence/M2-T06.json.

JPEG derivatives are lossy and do not replace preserved originals. The comparison
evidence applies to the approved synthetic corpus and pinned decoder/profile behavior;
codec/build and reference-intent qualifications remain explicit. Inaccessible subtrees
and ancillary reporting are next M3-T01; general deadlines, cancellation and native
descendant lifetime remain later M3 gates. Owner acceptance, releases and deployment
remain separate.

Fresh children may use process-only Bypass and clear inherited PSModulePath; never change persistent PATH/policy or redefine USERPROFILE. Keep raw paths, media, argv and tools ignored. Preserve actual failures. Use one successor draft if the owner merged this PR; normal scoped feature commits/pushes remain authorized. Merges, tags/releases, settings/deployment and owner acceptance remain separate gates. This session stops at M2-T06. Once requested, complete only M3-T01 and stop.
