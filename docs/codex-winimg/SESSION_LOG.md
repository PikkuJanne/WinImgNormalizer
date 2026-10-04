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

## 2026-10-04 — M0-T03 completed

Resumed clean a3ef050 independently equal to the feature remote. Origin fetch/push
identity verified; owner-merged PR #2 main at 8e35b6a had the same tree. Continued the
feature without pull/merge/reset; created successor draft PR #3.

Added minimal callable orchestration and positional adapter with internal output/
process injection. Import performs no work or caller mutation and callable paths
do not exit. Existing algorithms, launcher and assets retained. Added six contained
Pester tests, pinned archive bootstrap, credible runner and read-only Windows CI.

Actual PS 5.1/PS 7 each passed 6/6. Ten local controls verified expected good/false-
assertion/zero/skip/discovery outcomes with real counts and unchanged persistent
policies/source hashes. Child import gate precedes Pester import; a copied app
prefixed with exit 0 was detected (six failures/container 1). Both -File argument
forms ran with only an outer scratch destination injection. Independent ordinary
parity matched sixteen complete Pillow JPEG decodes and twelve exact video copies,
including timestamps/tree/source preservation against M0-T02.

First CI at 08ea3c3 failed before tests: Get-Command returned Windows and Git tar
records. Explicit System32 libarchive selection fixed the default bootstrap;
ten controls reran with Git tar first on child PATH. Corrected implementation
3b80ee94b8471a22560ea90a7583283c1613195c passed real push/PR Windows matrices, each job six tests plus
an exactly identified false seventh assertion/exit 1 control. Actual runner versions
and job links are retained. PS 5.1 PSScriptRoot parameter-default and development
command-path failures were corrected infrastructure issues, not test passes.

Evidence M0-T03.json plus runner/parity/import-control/CI JSON binds tested code
and source hashes. This record-only checkpoint does not modify the tested runtime/
tests. Final containing commit/push/sync is pending at write and will be observed
externally after commit. M0-T03 accepted; next M1-T01. No private paths/media/tools
tracked, no installer/persistent policy/PATH/default-branch/release change.
Known bugs/full corpus/analyzer/codec/owner quality/publication remain later gates.

## 2026-10-04 — M1-T01 completed

Resumed clean 866aba0 independently equal to feature remote, with canonical origin
fetch/push. Owner-merged PR #3 main b153d70 had the same tree; read-only fetch, no
pull/merge/reset/default-branch update. Draft successor PR #4 continues the feature.

Implementation f57b8adbb6859e8812471a4cc123b73530a2b301: friendly arguments/FileSystem source/Int64 validation,
root semantics, actual executable selection, bounded version/format preflight and
reviewed 7.1.2-32 dependency floor; upstream sources recorded. Missing decoders get
no retries, supported items continue with code2. Video-only/empty ImageMagick
requirement documented/tested. Destination DeleteOnClose probe, practical decimal
space budgets and friendly mkdir failures; skipped-codec inputs excluded. Exact
Int64 extent quotient fixes valid large divisible caps without changing default units.

Actual desktop PS5.1/PS7 each passed 66/66, zero skipped, on clean committed revision;
failure controls each 67 total/66 passed/one identified false assertion/exit1. Tested
source hashes match git blobs; persistent policies unchanged. Real contained -File
forms/failures, dependency shadow identity, ordinary JPEG full decode and video/source
hash/timestamp preservation pass. Root/codec/denied-write/space failure controls are
explicitly distinguished from live UNC/real HEIC/ACL/exhaustion/manual launcher claims.
Initial development assertion expectation failures (55/59 each) are recorded and
corrected; final suite supersedes them. Independent review corrections are retained.

Implementation push/PR CI both passed Windows Server 2025 PS5.1.26100.33438 and PS7.6.6
matrices, 66 passing tests and deliberate67th failure/exit1 per job. Actual jobs/log
counts/versions in M1-T01-ci.json. M1-T01.json binds implementation and per-file hashes.
Code synchronized while clean at 2026-10-04T14:59:09.193774+00:00; this containing
record-only checkpoint awaits final commit/push/external remote comparison.

M1-T01 accepted; next M1-T02, stop. Four of28 accepted; T001-T012 exercised, of which
T001-T004 are legacy characterization. Original snapshot/launcher/license/assets
preserved. No private media/tools/raw local logs tracked, no installation/persistent
PATH/global policy changes. Containment/run isolation and remaining hardening, owner
acceptance/release/website gates stay in later tasks.


## 2026-10-04 — M1-T02 completed

Started at clean 95cbc7956897794bcebd558f8395fdc0391a3baa, independently equal to the
advertised codex/winimg-hardening remote. Fetch/push URLs both verified canonical
PikkuJanne/WinImgNormalizer. Owner had merged PR #4; origin/main c604ade had the same
tree. Read-only fetch preserved local history; no pull/reset/rebase/merge. Successor
draft PR #5 contains this task. The non-Git WinImgNormalizer-main snapshot is preserved.

Implemented canonical whole-segment nesting checks before enumeration/probes, linked
ancestor rejection including dangling reparse entries, one nonrecursive inventory,
visible link/incomplete-scan warnings and safety rechecks. Exclusive timestamp/GUID
run allocation retries at most eight collisions and never adopts existing files or
directories. All observed source top-level names reserve a disambiguated generated
work/reports namespace; log CreateNew preserves arrivals without a TEMP fallback.
Ordinary mirrored paths and conversion defaults are retained. General exits, user
output collisions, dedupe, transaction/frame/colour work remain later scope.

Added 29 traversal regressions alongside the existing 66 checks. Actual directory
loop/outside/dangling junctions, forced file/directory run collisions, simultaneous
same-stamp child applications and namespace/log conflicts pass. File-symlink creation
on this desktop was denied Win32 1314; mandatory controlled reparse-file coverage
asserts the exact skip diagnosis. Source/video hashes and timestamps remain intact.
Drive/UNC boundaries are lexical; denied enumeration/post-inventory changes use
controlled seams. No real Pictures, private media, live UNC or ACL modifications.

Development evidence retains the initial incorrect -File array invocation (zero-test
infrastructure failure), long fixture paths, legacy duplicate-key fixture collision,
and dependency test with a newly forbidden nested output parent. Assertions were
preserved; fixtures/mocks corrected. Initial c6157fb desktop passed 94/94, but both
push 37212183765 and PR 37212205276 pwsh CI had 88/94 passed, six setup failures; PS5.1
jobs passed. An owned parent 231 / child 272 reproduced the native mkdir path limit.
Correction 8c9b157338419958a14c662ca379ccb8c50e8825 adds internal extended drive/UNC
prefixes and a 272-character native exclusive/canonical regression. No user name
truncation or public device-path option. Initial failed CI logs/results remain history.

Corrected actual Windows 11 Pro build 26300, PS 5.1.26100.9444 / PS 7.6.5 with Pester 5.9.1
and verified ImageMagick 7.1.2-32 passed 95/95, zero skipped. Deliberate controls each
had 96 total / 95 passed / one identified false assertion / exit 1. Four summaries match exact
committed runtime/test hashes; persistent execution policy is unchanged. No installer,
persistent PATH change, ImageMagick global policy edit or runtime bootstrap.
Corrected push 37212453070 and PR 37212456334 Windows Server matrices each passed both
PS5.1/PS7 jobs and expected failure controls; actual versions/counts/job links and
initial CI failures are in M1-T02-ci.json. Desktop evidence is M1-T02.json.

Corrected implementation pushed and independently matched clean local/remote HEAD
at 2026-10-04T15:17:42.8747086Z. Final record-only checkpoint changes STATUS/TASKS/
NEXT_SESSION/SESSION_LOG and the two evidence files only, binding corrected code
and hashes. Its containing commit/sync fields remain pending_verification until
observed externally after commit/push. No merge, release, deployment or owner image
quality acceptance. M1-T02 accepted; next M1-T03, stop. Five of 28 accepted and
T001-T016 exercised, with T001-T004 retained as characterization.

## 2026-10-04 — M1-T03 completed

Started from clean e178acb, independently matching the verified feature remote.
Owner-merged PR #5 main cd64fee had the identical tree; read-only fetch preserved
the checkout/history. Continued the authorized feature branch and opened successor
draft PR #6 after normal push. Preserved the non-Git snapshot.

Implemented deterministic complete eligible-source mapping with total ordinal row
ordering and OrdinalIgnoreCase reservations. Mirrored directories, generated paths,
video names and every unique legacy image name precede suffix allocation. Colliding
images receive lowercase extension suffixes and checked numeric suffixes starting
at 2. Local PLAN lines persist the complete map before conversion, including later
missing-codec/duplicate/failed rows. Added neutral image candidates in exclusive GUID
work directories, no-overwrite File.Move and File.Copy(false), final availability/
ancestor checks, and partial status for new naming/staging/cleanup failures.

Added 24 mandatory naming checks; independent read-only review found no remaining
blocker. Actual JPG/PNG/BMP collisions full-decode; source hashes/timestamps and
unchanged video bytes pass. Controlled HEIC routing/case-only inventory, 80 seeded
shuffle/culture comparisons, complete logs, file/directory arrivals including
post-check races, allocation failure continuation and >260 native scratch pass.
Initial 95-test development runs passed 92/failed 3 per shell because native scratch
hit MAX_PATH. Internal extended output spelling corrected it without shortening
names; focused six-test ordinary invocation suites passed. Combined development
119-test runs passed 116/failed 3 because preflight controlled ownership guards
rejected that spelling. Removing only the extended prefix before the same guard
corrected fixture handling, without reducing behavior assertions. Raw failures and
sanitized counts/hashes remain in M1-T03.json history. A one-off evidence validator
syntax typo was corrected before any validation assertions or tracked writes.

Committed runtime/tests as da91c0d648b387b270bcdd7269a48ee1c63d11eb. Both actual Windows 11 shells passed
119/119 with zero skips; controls 120 total/119 passed/one identified false assertion/
exit 1. All four summaries match exact committed runtime/test bytes and unchanged
persistent execution policy. Both push and PR Windows Server matrices passed all
four jobs and failure controls; actual versions/counts/links are in M1-T03-ci.json.
No installer, persistent PATH/policy edit, global ImageMagick policy change, runtime
bootstrap, private media, owner visual acceptance or manual launcher claim.

Implementation push independently matched clean feature HEAD. Final record-only
checkpoint changes STATUS/TASKS/NEXT_SESSION/SESSION_LOG and two sanitized evidence
files, binding the tested implementation. Its own sync remains pending_verification
until externally observed after commit/push; final SHA/live sync belongs in the
response/PR. M1-T03 accepted; next M1-T04, stop. Six of 28 accepted, T001-T019
exercised with T001-T004 retained as characterization. Full candidate/native outcome
validation, video-copy transactions, frames/colour/duplicates/general lifecycle and
owner/publication gates remain later scope.

## 2026-10-04 — M1-T04 completed

Started from clean f73a4d0, independently matching the verified feature remote.
Owner-merged PR #6 main acd2959 had the identical tree; read-only fetch preserved
checkout/history. Continued the authorized feature and opened successor draft PR
#7 after normal push. Preserved the non-Git snapshot.

Each native image attempt now exclusively reserves its own neutral candidate,
including fallback and lower scales. Successful native outcome, nonempty regular
bytes, full single-frame JPEG decode and positive dimensions precede measured-byte
acceptance; validation rechecks length/time. Superseded scratch is cleaned before the
next attempt and cannot become stale best-effort success. Videos stream into owned
partials, flush/close and verify output/source length plus source modification time.
Same-volume no-overwrite File.Move preserves final arrivals. Exact owned cleanup
preserves unknown numbered outputs and unreserved neighbors. Item errors return 2.
Decision D22 and primary API reference W3 record choices and limits.

T020-T024 add invalid/ambiguous outcomes, actual truncated-header/full-pixel decoding,
real neutral native long-path validation/move, numbered outputs, fresh fallback/stale
attempts, final-name arrivals, stable opaque video hashes, controlled source changes/
partial writes/stream failures and exact ownership. Controlled cancellation/timeout
results do not certify actual process termination. No runtime video content hashing,
same-length/same-timestamp identity, crash durability or hostile-filesystem guarantee.

Development history: An exploratory development run passed 119/119 while runtime
cleanup and source line-ending edits were still in progress. Its counts are
retained as nonrevision-bound observations; runner hashes describe end-of-run
files and do not prove every import executed those bytes. Initial targeted 32-case
suites each passed 31 and failed one fixture-oracle assertion: converting a
truncated JPEG to null: returned zero instead of the expected nonzero. A further
resize/PNG probe emitted premature-end/corrupt warnings but returned zero and
wrote 250 bytes on the pinned build. Explicit identify +ping -regard-warnings
rejected the same truncated bytes with exit 1; the fixture oracle now checks
readable header dimensions against that explicit full-pixel mode. The runtime
already used it and its rejection regression had passed. Corrected targeted suites
passed 32/32; these remain uncommitted development observations. Initial clean
8f10472 desktop suites each passed 151/151 with 152-test controls proving one
false assertion/exit 1. All four initial hosted jobs failed with 150/151 tests
passing because test-side inspection of retained numbered siblings used ordinary
long paths. Test-only ed7fb76 applies the existing extended native path helper to
two inspection calls; runtime code is unchanged. Corrected final2/control2 suites
and hosted push/PR matrices above pass at the exact committed checkout binding.

Committed runtime/tests as ed7fb7691eca9dd55082338230657804c1865681. Actual desktop PS 5.1 / PS 7 each passed
151/151, zero skipped; controls 152 total / 151 passed / one identified false
assertion / exit 1. All four summaries match the same clean committed checkout bytes;
evidence records raw hashes, committed blob hashes and Git's CRLF/LF normalization.
Persistent policies are unchanged. Push and PR Windows Server matrices passed all four
jobs and failure controls; actual versions/counts/links are in M1-T04-ci.json.
Desktop evidence is M1-T04.json. No installer, persistent PATH/policy, ImageMagick
global policy, runtime bootstrap, private media or owner visual acceptance changes.

Implementation sync is a past exact-SHA observation. The final record-only checkpoint
updates STATUS/TASKS/NEXT_SESSION/SESSION_LOG/DECISIONS plus two sanitized evidence
files, binding the tested implementation. The primary reference update is in the
implementation checkpoint. Its containing SHA/sync remains
pending_verification until externally observed after normal commit/push. No merge,
release or deployment. M1-T04 accepted; next M1-T05, stop. Seven of 28 accepted and
T001-T024 exercised with T001-T004 retained as characterization. Duplicate early
registration/length guard, native lifecycle, frames/colour/general cancellation and
owner/publication gates remain later scope.

## 2026-10-04 — M1-T05 completed

Started from clean 8c9464b, independently matching the verified feature remote.
Owner-merged PR #7 main d8936872 had the identical tree; read-only fetch preserved
checkout/history. Continued the authorized feature and opened successor draft PR
#8 after normal push. Preserved the non-Git snapshot.

Replaced early duplicate seen-key registration with retained successful source/output
mapping after verified image/video finalization. The existing case-normalized filename
plus modification-time key now requires matching input length before skipping.
Deterministic total ordinal plan order selects retained relationships and every skip
logs its retained source/output/status. Failed conversions/copies/final moves leave no key;
different lengths continue independently. A validated above-target finalized image
may retain its existing success-with-warning note. Live source metadata precedes
lookup; changed image-source length/time after finalization keeps valid output with
a partial warning but prevents heuristic registration. Videos register the verified
copy helper's length/time snapshot. Same-key/same-length different
content can still match falsely; this remains a documented heuristic without runtime
content hashing, hash database or original removal. Decision D23 records the choice.

T025-T027 regress failed-first suppression, equal-length guarding and stable retained
links/heuristic limits. Existing mandatory M1 suites retain source hashes and creation/
modified timestamps, real full JPEG decode, unchanged synthetic video bytes, output
isolation, conflict preservation and exact owned cleanup. Public calls, byte cap,
scales, conversion settings and runtime dependencies are unchanged.

Development history: The first targeted development suites each passed 21/21 in
actual Windows PowerShell 5.1 and PowerShell 7, with zero failed or skipped tests.
Review strengthened the JPEG inspection helper to use the existing native
extended-path spelling and the per-length test to check both later mappings (short
C to short A, long D to long B). Targeted reruns again passed 21/21 in both
shells. These four runs retain end-of-run hashes and summary/XML artifacts as
uncommitted development observations. Initial clean committed implementation
4fce2f0e705a6a40fd2fa8f8804422c14bd4a5cb full suites each passed 171/172; controls
passed 171/173 with the intended T007 failure plus the same unexpected inherited
T013 failure. The source-reparse guard still prevented copying, but its broadened
diagnostic no longer matched the established assertion. Correction c9988e2
restores the original reparse-specific diagnostic and keeps a separate directory
guard. The existing T013 assertion is unchanged; initial failures and exact
raw/Git checkout binding remain recorded. The initial revision was superseded
before implementation CI; no CI result is claimed for it. Final acceptance relies
only on the corrected committed-revision normal suites, identified deliberate-
failure controls and actual hosted push/PR matrices above.

Committed runtime/tests as c9988e204d25bbb13616b6df601d0f5c400a73c1. Actual desktop PS 5.1 / PS 7 each passed
172/172, zero skipped; controls 173 total / 172 passed / one identified false
assertion / exit 1. All four summaries match the same clean committed checkout bytes;
evidence records raw hashes, committed blob hashes and any Git CRLF/LF normalization.
Persistent policies are unchanged. Push and PR Windows Server matrices passed all four
jobs and failure controls; actual versions/counts/links are in M1-T05-ci.json.
Desktop evidence is M1-T05.json. No installer, persistent PATH/policy, ImageMagick
global policy, runtime bootstrap, private media or owner visual acceptance changes.

Implementation sync is a past exact-SHA observation. The final record-only checkpoint
updates STATUS/TASKS/NEXT_SESSION/SESSION_LOG/DECISIONS plus two sanitized evidence
files, binding the tested implementation. Its containing SHA/sync remains
pending_verification until externally observed after normal commit/push. No merge,
release or deployment. M1-T05 accepted; next M1-T06, stop. Eight of 28 accepted and
T001-T027 exercised with T001-T004 retained as characterization. M1-T06 full milestone
audit, native lifecycle, frames/colour/general cancellation and owner/publication
gates remain later scope.

## 2026-10-04 — M1-T06 completed

Started from clean 91e7f4e, independently equal to the feature remote. Owner-merged
PR #8 main dc2a2a94 had the identical tree; read-only fetch preserved checkout/history.
Continued the authorized feature and opened successor draft PR #9 after normal
push. Preserved the separate non-Git snapshot.

Closed the covered output-safety milestone with T028 and combined mandatory M1
regressions.
One actual mixed synthetic source tree is normalized twice into the same
prepopulated output parent. After each run all 15 source files retain SHA256,
length, creation/modification ticks and attributes, and source directory
metadata/paths remain unchanged. Each run contains eight fully decoded JPEGs, two
byte-identical video copies and one log with the declared collision map and two
retained duplicate links. The second run allocates a separate directory and
preserves the first run media/log/tree and existing user entries.
M1-T06 changes test coverage and test documentation; runtime and batch bytes are
unchanged from the starting checkpoint.
The complete milestone diff from 866aba02 to 4d7f9f2e7681902971b584560211139d555bf011 was reviewed;
its exact diff hash/path list and finalization/collision/public-interface/dependency
findings are retained in sanitized evidence. No blocking audit findings remain.
No new product decision was needed; D20-D23 and their limits remain applicable.

Initial targeted PowerShell 7 development run M1-T06-target-PS7-1 failed 0/1 with
zero skips because the test's timestamp-inspection Get-Item omitted -Force for the
actual hidden source. The fixture and runtime behavior remained unchanged; the
correction adds -Force to the inspection while retaining all original preservation
assertions. Targeted PowerShell 7 rerun PS7-2 and first PowerShell 5.1 run PS51-1
each passed 1/1 with zero skips and exit 0. These three observations retain actual
end-of-run hashes and summary/XML artifacts as uncommitted development history.
Initial implementation push attempt 1 failed its PowerShell 7 development-
dependency bootstrap with observed HTTP 403 before test discovery: zero tests,
exit 1, control skipped. The failing download endpoint was not present in the log,
so neither dependency is identified as the failed request. Its PowerShell 5.1 job
and both PR jobs passed all 173 normal tests and identified
174-total/173-pass/one-failure controls. One full push-workflow retry used the
same implementation, pins and guards without code/test changes; the final recorded
attempt provides the required normal/control acceptance independently of the
retained initial failure. Earlier task failures remain in their original evidence.
Final acceptance uses only the clean committed-revision full desktop
normal/control suites and actual completed hosted push/PR matrices.

Actual desktop PS 5.1 and PS 7 passed 173/173 with zero skips at the clean
committed implementation. Controls each have 174 total/173 passed/one exact
T007 false assertion/exit 1. All four summaries bind the same source/test checkout
with raw hashes, Git blob hashes and verified CRLF/LF normalization recorded.
Persistent execution policies remain unchanged. Actual push/PR Windows Server
matrices passed all four jobs and expected controls; count/environment/log-hash
observations are in M1-T06-ci.json.

Implementation synchronization is a past exact-SHA observation. Six record-only
paths update STATUS/TASKS/NEXT_SESSION/SESSION_LOG and two sanitized evidence files.
The containing checkpoint SHA/live sync remains pending_verification until externally
observed after normal commit/push. No private media, raw logs or local tool binaries
are tracked. No merge, release, deployment or owner quality acceptance was performed.

M1-T06 accepted: 9/28 tasks, T001-T028 exercised with T001-T004 characterization.
Next M2-T01 — Implement deliberate first-frame and first-page handling; stop.
First-frame/page, colour/metadata, native lifecycle,
general cancellation/exits, complete corpus/analyzer, manual owner and publication
gates remain later scope. Synthetic source creation/modification timestamps are
covered; access times, heuristic content identity, live UNC/real Pictures, universal
long paths, hostile filesystem safety, batch atomicity and crash durability are not.
