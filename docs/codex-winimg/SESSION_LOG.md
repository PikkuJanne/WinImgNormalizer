# Session log

## 2026-10-04 — Bundle preparation

The planning bundle was prepared from the reviewed pinned GitHub baseline. No user
local checkout was inspected or modified, no GitHub branch/PR was written, and no
normalizer application tests were executed for this bundle. Helper self-tests and
bundle integrity checks are recorded in the external VALIDATION_REPORT.md.

Append one compact entry per Codex task with actual branch, implementation revision
or tested file hashes, commands/results, evidence file, unresolved blockers and next
task. A post-commit remote observation can be reported externally and logged at the
next session. Never insert an invented final checkpoint SHA or a prospective success.

## 2026-10-04 — M0-T01

Reconciled the requested workspace: it was a non-Git seven-file source snapshot,
all blobs equal to the historical baseline. Read-only discovery found no existing
WinImgNormalizer checkout, and the configured project confirmed no Git metadata.
Preserved the snapshot and used the documented fresh-clone fallback into the
previously absent `C:/projects/WinImgNormalizer`. No parent/root/nested AGENTS or
prior task progress existed there. Verified canonical origin fetch/push identity,
clean `main`, live refs and baseline ancestry; local and remote starting HEAD were
`4b918da50639d8fc7e3fddc6d6b3678880ea4098` with no source delta or divergence.
Created `codex/winimg-hardening` from that reconciled clean point.

Verified the immutable bundle (63 files), reviewed importer preview (56 adds,
zero conflicts), and applied with the exact reviewed HEAD. All 56 imports matched
the allowlist hashes before the intentional status/registry/handoff updates.
Validated the task graph (28 tasks, 75 cases, 18 review items). External bundle
tests ran on Windows: 48 total, 45 passed, 1 case-sensitive fixture failure,
2 unavailable-symlink skips; an explicit colliding-manifest control passed.
No application test, image conversion, dot-sourcing or software installation ran.

Evidence: `evidence/M0-T01.json`, including environment, baseline blobs, actual
command outcomes and precommit tested-file SHA-256 hashes. This is the one-commit
checkpoint model from GIT_WORKFLOW.md; committed content must match those hashes.
Push/remote comparison and final SHA are pending while this record is written;
their actual results belong in the final response and draft PR. Baseline had no
workflows, runs or PRs. M0-T01 is accepted; next task is M0-T02, with only that task
to run in a fresh thread. The inherited helper fixture limitation and missing
runtime tooling remain explicit; all application coverage is unrun.

## 2026-10-04 — M0-T02 partial checkpoint

Resumed the actual checkout on codex/winimg-hardening, clean and independently
equal to remote d323be72cbd61fe207af6b7876f52b2d809a2311. Verified origin fetch/push
identity. Read-only fetch showed the owner merged PR #1 into main at
2a51c0b6fa36370ee8341ff0d6c9b0b638a27a73; both trees are equal. No branch rewrite,
merge, reset, default push or application/source asset change occurred.

Added the fixture generator, shell inventory, contained snapshot wrapper, bounded
real-run controller and native fakes under tests/legacy, plus ignored scratch.
Executed bundled Python 3.12.14/Pillow 12.3.0 check_harness.py on actual Windows
PS5.1.26100.9444 and PS7.6.5: six instrumented fake-process runs, ten rejection
controls, six Unicode destinations, source preservation and persistent-policy
checks completed. Mocked nested planning executed without filesystem recursion.
Independent process controls returned exit9 with zero/23-byte invalid outputs;
legacy counted those outputs as converted. First-failure trace showed twelve
attempts against a, none against b, and a logged duplicate skip in both shells.

Infrastructure issues discovered and corrected: PS5.1 command discovery returned
multiple Git executables; inherited PS7 PSModulePath prevented PS5.1 module loading;
missing-marker control initially hit an earlier outside-source guard. The final
formal run exercises isolated marker rejection and edition-default modules.
Snapshots inject destination using an ASCII Base64 UTF-8 expression to preserve
Unicode under PS5.1's original no-BOM script encoding. Earlier failed previews are
kept in ignored scratch and are not reported as application passes.

Python syntax, generator help/missing-tool rejection and structural plan checks
completed. Full ImageMagick generator and T001/real collision/frame observations
remain unrun: no executable was found. The owner was asked about a pinned verified
official portable tool or an existing path; no reply/download/installation occurred.
M0-T02 stays blocked/incomplete, next task remains M0-T02. Sanitized commands,
transcripts, environments, fixture/input/code hashes and limitations are in
evidence/M0-T02.json and M0-T02-observations.json. This one-commit checkpoint binds
tested files by precommit hashes; final sync is pending until the actual commit,
normal push and independent remote comparison reported externally. CI unconfigured.

## 2026-10-04 — M0-T02 completed after owner tooling authorization

The owner explicitly requested download and continuation. Verified official
portable ImageMagick 7.1.2-32 Q16 x64 archive SHA-256 and 11984505-byte size, extracted
into ignored scratch and verified native version/executable hash. No installer,
persistent PATH, global policy, application-source or asset change occurred.
Resumed clean feature 38f1a33 independently equal to advertised remote; main 2a51c0b
and draft PR #2 remained unchanged. No pull, merge or history rewrite.

Initial full build failed the independent GIF canvas assertion before any app run.
Scoped per-frame -set page fixed the recipe, preserving assertions. Updated WebP
capability parsing for the actual native format listing; actual animated WebP ran.
The final generator/audit inspected 45 fixtures, source hashes/timestamps/properties,
with provenance and all-frame decode. Raw failure/debug artifacts were retained.

Full controller ran 12 observations (6 real / 6 controlled fake) and passed 22 harness
checks on PS5.1/PS7. Ordinary 4 JPEG + 3 video hashes/timestamps/tree/logs are recorded.
Collision output overwrite, directory false OK and duplicate interaction, numbered
GIF/TIFF/WebP output with reported errors, failed-first duplicate and weak-output
acceptance are actual known defects. All sources preserved and persistent policies
unchanged; twelve snapshots modify only the exact Pictures block. No real Pictures
write, unsafe recursion or dot-sourcing. Independent fixture audit passed 45/45.

T001-T004 characterization completed; M0-T02 accepted, next M0-T03. Corrected
regression tests 0, CI unconfigured, owner quality acceptance not requested. HEIC/
profiles/further failure fixtures remain explicit future coverage. Final evidence
M0-T02.json binds precommit source/file hashes; actual commands/transcripts/manifest
are in M0-T02-real-observations.json and M0-T02-fixtures.json. Prior partial evidence
is preserved separately. Final containing checkpoint push/sync is pending at write
and will be reported externally after normal commit/push/independent comparison.
