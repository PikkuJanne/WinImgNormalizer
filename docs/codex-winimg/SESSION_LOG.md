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

## 2026-10-04 — M2-T01 completed

Started from clean b7fe275d, independently equal to the feature remote. Owner-merged
PR #9 main f90de6ce had the identical tree; read-only fetch preserved checkout/history.
Continued the authorized feature and opened successor draft PR #10 after normal
push. Preserved the separate non-Git snapshot.

GIF/TIF/TIFF/WebP/HEIC/HEIF inspection and conversion read the same exclusively
created owned snapshot with a neutral source basename and original extension.
Source regular-file/reparse checks and length/modification-time checks surround
the copy, and copied length must match. External snapshot arrivals and unknown
neighboring files are preserved; exact owned snapshots follow existing cleanup and
partial-warning rules.

A separate successful identify -ping probe counts every decoder-exposed image and
rejects malformed, inconsistent or ambiguous count/format/dimension observations.
Conversion selects image:frames=0 before the owned native input. Only actual
decoded GIF/WebP follows FirstDisplayedFrame coalescing onto its logical canvas;
+repage and existing auto-orientation follow. TIFF keeps one first page without
stacking; HEIC/HEIF keeps the decoder's primary/first image. Each finalized output
still must fully decode as one nonempty JPEG at the exact planned path.

SOURCE IMG records source count, selected count 1, omitted count, unit, policy and
actual decoder. Deliberate omission is the normal informational static-output
policy and retains exit 0 on successful runs. Source bytes/times, mirrored naming,
no-overwrite finalization, verified videos, success-based heuristic duplicate
links, size cap/scales/JPEG flags and public BAT/positional forms retain their
established behavior. The bootstrap, dependency pins and workflow were not
changed.

Decision D24 records the scoped frame/page policy.
T029 — passed: Two actual GIF variants, with one or two frames, independently
establish a transparent 24x20 first tile at +8+10 on a 64x48 logical canvas.
Literal bracket paths produce one exact mirrored 64x48 JPEG with white outside and
the first red region; later blue pixels and numbered outputs are absent. Source
count, selection and omissions are logged; source bytes, length, timestamps and
attributes remain unchanged.

T030 — passed: An actual two-page TIFF has an asymmetric 80x48 first page whose
first IFD Orientation tag is independently set/read as RightTop, followed by a
distinct blue page. One 48x80 JPEG contains the correctly oriented first-page
red/green/yellow samples, with no page overlay or second-page blue. FirstPage,
count 2 and omission 1 are logged; exact output, empty owned work and source
bytes/times are checked.

T031 — passed: An actual two-frame animated WebP requires sRGBA first-frame
channels and a transparent corner before asserting a white/red 64x48 first
displayed JPEG and omitted blue frame. A genuine 1328-byte HEVC collection with
two top-level still images is independently decoded through both .heic and .heif
aliases, exposing HEIC and HEIF labels respectively; one 64x48 JPEG retains the
asymmetric primary red/green/yellow image and omits the distinct blue image.
Counts, policy, exact single outputs, owned cleanup and source bytes/times are
checked. This verifies a still collection; timed HEIC animation remains
unverified.

Four uncommitted targeted development runs are retained. The first PS7 run passed
6/12 and failed six assertions: an HEIF label expectation, two mock source-path
errors, recursive copied-metadata inspection, a cleanup warning expected as
success, and an unsupported white/transparent WebP expectation. Independent RIFF,
pinned decoder and Pillow inspection found the original WebP alpha-free; its
optional white animation-background hint did not establish white rendered padding.
The second PS7 run passed 11/12 and failed an alpha expression applied to an RGB-
only second frame. Corrected fixture/mock/oracle checks passed 12/12 in PS7-3 and
PS5.1-1; no runtime fix was required for those test-development failures. Raw
summaries, XML, consoles and saved source snapshots stay ignored with hashes;
their end-of-run source hashes do not establish an immutable clean execution
revision. Separately, the native HEIC encode probe and first unpacked generator
DLL import failed during fixture tooling development; verified once-only
generation and independent actual two-image decoding then succeeded. These tooling
events are distinct from runtime assertion results and from the subsequent clean
committed full acceptance runs.

Committed runtime/tests as d64a9034ff0b4324b2515900fdf8b4cb82f80bbb. Actual desktop PS 5.1 and PS 7 each
passed 185/185, zero skips; controls 186 total/185 passed/one exact
T007 assertion failure/exit 1. All four summaries bind the same clean tested checkout;
raw hashes, Git blob hashes and verified CRLF/LF normalization are recorded. Persistent
execution policies are unchanged. Actual push/PR Windows Server matrices passed all
four normal/control jobs; environment/count/log-hash evidence is in M2-T01-ci.json.

Implementation synchronization is a past exact-SHA observation. The containing
record-only checkpoint updates STATUS/TASKS/NEXT_SESSION/SESSION_LOG and two sanitized
evidence files, plus the scoped D24 decision and primary references. Its own SHA/live sync stays pending_verification until externally
observed after normal commit/push. No raw logs, private media or local tools are tracked.
No merge, release, deployment or owner aesthetic approval was performed.

M2-T01 accepted: 10/28 tasks, T001-T031 exercised with first four characterization
and codec coverage/absence stated explicitly. Next M2-T02 — Convert colour profiles before metadata removal; stop.
Actual T031 HEIC/HEIF evidence is a genuine two-image top-level still collection
through two aliases. Timed HEIC animation, thumbnails, auxiliary images and
arbitrary codec builds remain unverified; container counts describe only images
exposed by the pinned decoder. The pinned native build reads HEIC/HEIF but cannot
generate the fixture; one verified official development wheel produced the
synthetic bytes in ignored scratch, with no installation, runtime/CI generator
dependency or distributed generator binary.

Snapshot length/modification-time checks and duplicate keys are stability
heuristics, not proof of content identity. General native argument/input-grammar
safety and source changes retaining identical metadata remain outside this narrow
neutral-snapshot policy. Owned cleanup deliberately preserves unknown entries and
may return partial exit 2 after a valid JPEG has finalized.

Tagged colour/profile conversion before metadata removal is next M2-T02. Broader
alpha/colour/reference fidelity, size-search quality, process timeout/resource
budgets, cancellation, reporting and publication remain later tasks. Automated
internal peer review does not grant owner aesthetic acceptance or authorize merge,
release, deployment, tags or default-branch changes.

## 2026-10-04 — M2-T02 completed

Started clean at ad56bdf9, independently equal to the feature remote. Owner-merged
PR #10 main 58d628ce had the identical tree; read-only fetch preserved checkout/history.
Continued the authorized feature and opened successor draft PR #11 after normal
push. Preserved the separate non-Git snapshot.

Every eligible image is copied once into a stable owned neutral snapshot;
frame/page inspection, source colour inspection and all conversion attempts use
those same preserved source bytes. Strict native inspection clears conflicting
free-form profile/colorspace properties without stripping actual ICC
characterization.

Tagged sources retain their source ICC through owned exact extraction and bounded
header/tag/model checks, then transform to a hash-verified embedded CC0 sRGB-v4
target. Requested Relative intent explicitly disables black-point compensation.
Decoded untagged sRGB is assumed sRGB; native linear RGB and gray are converted
explicitly. Untagged CMYK without source characterization is rejected with a
source-preserving partial error.

Auto-orientation precedes colour conversion. RGB transforms precede compositing
over white in encoded sRGB; alpha is then disabled and
ICC/EXIF/GPS/XMP/IPTC/comments are removed. Every 100, 90, 80, 70, 60 and 50
percent attempt uses that same order and white-alpha policy. The inherited
alpha-dropping 100-percent retry is removed so a retry cannot change the
background policy.

Malformed/mismatched profiles, ambiguous inspection and nonempty native diagnostic
output cannot claim accurate successful conversion. Existing transactional full
single-JPEG validation and no-overwrite finalization still decide whether output
is retained; later valid files continue after a source error, and exact owned
cleanup preserves unrelated entries.

The established positional/BAT interface, 1,048,576-byte default, scale sequence
and best-effort boundary, first-frame/page selection with informational omissions,
mirrored naming, source preservation, copied videos and success-only duplicate
links remain the regression baseline. No runtime dependency or dependency pin
changes; the runtime target profile is embedded rather than downloaded.

Three CC0 source/target profiles and an optional reproducible independent Pillow
12.3.0/LittleCMS 2.19 recipe provide synthetic references. Mandatory tests use
checked reference data, with distinct 3/255 pre-JPEG and 12/255 JPEG interior
patch tolerances, and require the actual T032-T036
RGB/CMYK/malformed-ICC/orientation/metadata/alpha cases.

Decision D25 records the tested colour/profile and metadata policy.
T032 — passed: Actual wide RGB, already-sRGB and spoofed-free-profile-property
sources use the real runtime conversion arguments; lossless pre-JPEG and final
JPEG patch values match independently precomputed Pillow/ImageCms references.

T033 — passed: Actual four-channel tagged CMYK uses its available forward A2B0
mapping, untagged RGB assumes sRGB, declared linear RGB encodes to sRGB, and
untagged CMYK fails before the conversion runner.

T034 — passed: Real malformed PNG and retained JPEG ICC, real profile/model
mismatches, structurally accepted invalid-curve and unsupported-version ICC all
fail without final output; the actual default native capture rejects a real zero
exit with visible diagnostics and decodable JPEG bytes. Controlled
zero-with-diagnostics coverage remains explicitly controlled.

T035 — passed: All eight actual EXIF orientations match independent displayed
corner maps/dimensions; synthetic GPS/EXIF/XMP/IPTC/comments/ICC are absent from
final JPEG segments. Source byte hashes, creation and modification times remain
unchanged.

T036 — passed: Actual tagged and untagged transparent/semitransparent patches
composite on white after managed sRGB handling; all six real native scales retain
profile/white/strip semantics and the same owned target path with trusted profile
bytes.

M2-T02 development history is retained separately from clean implementation
acceptance. The first prior-suite native -File call passed one comma-containing
Path, failed before discovery with zero tests/exit 1 and no XML/source hash
inventory; the corrected literal four-path child wrapper passed 111/111 in PS7.
Original Colour tests passed 23/23 in PS7 and failed six actual assertions in
PS5.1 (17/23) because a test-side typed-array alias changed JPEG argv when
building the pre-JPEG PNG oracle. Clone() corrected the test adapters without
runtime changes or weakened assertions. Intermediate 24/24 and final 25/25
targeted runs passed in both shells; the added cases distinguish real
invalid-curve failure from a structurally accepted unsupported ICC version
returning native zero with visible diagnostics and decodable JPEG, which the
actual default runtime capture rejects. Initial isolated policy rejection, shell
quotation failures, exploratory native/property/argv probes and successful
optional reference reproduction with deprecation output are tooling history, not
Pester or clean-I acceptance. A supplementary PS5.1 provider collector hung after
app return2/errors1 and its exact owned helper was stopped; no completed
supplementary-summary pass is claimed. Authoritative target3 case25 passed
independently. All eight development summaries and available XML/console/source
snapshots are preserved; observed end-of-run file hashes do not establish
immutable committed execution. Final full/control desktop and hosted gates bind
the separate clean tested implementation.

Committed runtime/tests as a30c456594a31368294588cc198444b38bc3e0e3.
Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444
and PS 7.6.5, Pester 5.9.1: 210/210 in each
desktop shell, zero skipped. Controls each have 211 total/210 passed/exactly
one T007 false assertion/exit 1. All four summaries and actual native child exits
bind the same clean committed checkout, with raw file hashes and separate Git blob
hashes for verified CRLF/LF normalization. Persistent execution policies are
unchanged. Actual hosted Windows Server matrices passed all four normal/control
jobs: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37226399665) and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37226403159).
Evidence: evidence/M2-T02.json and M2-T02-ci.json. Licensed profile/fixture provenance,
independent reference method, separate tolerances and retained history are explicit.

Implementation synchronization is a past exact-SHA observation. The containing
record-only checkpoint updates STATUS/TASKS/NEXT_SESSION/SESSION_LOG and two sanitized
evidence files, plus the scoped D25 decision and primary references. Its own SHA/live sync stays pending_verification until externally
observed after normal commit/push. No raw logs, private media or local tools are tracked.
No merge, release, deployment or owner aesthetic approval was performed.

M2-T02 accepted: 11/28 tasks, T001-T036 exercised with first four characterization
and prior codec limits explicit. Next M2-T03 — Make filesystem and native filename handling literal-safe; stop.
The selected CGATS001Compat CMYK profile contains only an A2B0 perceptual forward
display mapping. Requested Relative intent uses LittleCMS fallback to that
available mapping; evidence covers this source fixture rather than general
relative-colorimetric printing accuracy, all rendering intents or reverse CMYK
conversion.

The ICC guard checks bounded structure and decoded model consistency, not complete
ICC semantic conformance. Synthetic patch and metadata fixtures do not prove
arbitrary ICC profiles, every codec, production photography or subjective owner
quality.

Length/modification checks remain a source stability heuristic rather than content
identity. Duplicate matching remains the documented same-name/time/length
heuristic; mandatory hashing, content databases and original removal remain
excluded.

Actual HEIC/HEIF coverage retains the two-image still collection from M2-T01.
Timed HEIC animation, thumbnails and auxiliary-image behavior remain unverified.

General literal-safe filenames, supported long/UNC/root boundaries,
process/resource/cancellation controls, later quality/reporting/release work and
owner aesthetic/default acceptance remain later explicit gates. No merge, tag,
release, website deployment or owner quality approval is implied.

## 2026-10-05 — M2-T03 completed

Started clean at ecd0703a4c340741095a5b5b7875c1b3e9eb78ce, independently equal to the feature
remote. Owner-merged PR #11 main 10c21366085ad83814bbfda85533cbeead93b571 had the identical tree;
read-only fetch preserved checkout/history. Continued the authorized feature and
opened successor draft PR #12 after normal push. Preserved the separate
non-Git snapshot.

Every native image count, selected colour inspection, source ICC extraction, JPEG
conversion and full final JPEG decode sets registry:filename:literal=true before
the input or output. The process-local registry guard protects percent expressions
in parent directories as well as filenames, and cannot be overridden by a source
filename:literal=false property. Owned neutral snapshots and candidates are
finalized by literal no-overwrite filesystem moves to the deterministic original
identity.

The real selected-input regression places distinct red and blue numeric
counterparts beside a percent-template snapshot. Without the guard, image:frames=0
can return the wrong neighbour with native exit 0; the guarded operations read the
intended source, preserve first-frame/page and ICC policies, and retain literal
final JPEG/video names. Opaque video directory creation uses the literal .NET
Directory API.

The destination boundary retains canonical whole-segment and reparse checks and
adds read-only 128-bit FileIdInfo directory identities for the source and each
existing output ancestor. Local-drive/localhost-share aliases that would place
output inside the source reject before enumeration, probes or allocation. Missing,
unsupported or zero identities fail setup rather than fall back to an unverified
lexical alias boundary.

The BAT localizes DisableDelayedExpansion and conditionally appends a dot after a
trailing slash/backslash, preserving folder/root identity through quoted Windows
PowerShell 5.1 argv. Actual direct positional calls and byte-for-byte copied BAT
routes preserve the tested spaces, apostrophes, ampersands, parentheses, bangs,
percent/brackets and Finnish/German/emoji names. Caller CMD expansion that happens
before BAT entry remains documented; single-quoted direct PowerShell is the
alternative for paired percent expressions.

Seventeen mandatory Paths cases combine real pixels and exact output identity,
tagged ICC, animated first-frame, video byte/time fidelity, external final-name
arrivals, actual launcher argv/application returns, physical directory identities,
lexical roots and supported long paths. Separate owned localhost C$ probes cover
direct UNC input/output, BAT UNC source and both alias rejection directions. Full
suite counts are read from clean-I results, with no unavailable UNC capability
folded into mandatory pass counts.

Public positional forms, standalone PS1/BAT operation, byte cap, six-scale
sequence, JPEG flags, first-frame/page omission reporting,
ICC/orientation/white-alpha/strip order, deterministic naming, transactions,
source preservation and success-only heuristic duplicate links remain the combined
regression baseline. Workflow, bootstrap, dependency pins and trusted profile
assets are byte-unchanged; no host/share/credential/persistent PATH or policy
configuration is introduced.

Decision D26 records the tested literal filename and path-capability policy.
T037 — passed: Literal percent/%d/%03d/bracket/hash/at-sign/leading-hyphen names,
nested first-GIF-frame selection and opaque copied video retain exact output
identities and full JPEG pixels. Genuine ICC and filename:literal=false property
remain safe in percent ancestors. Real snapshot-delegating observer creates a
differently colored numeric counterpart; all native reads are real. External
final-name arrival is preserved.

T038 — passed: Actual Windows PowerShell 5.1 -File source-only/explicit-cap and
byte-for-byte copies of the frozen BAT under inherited CMD delayed expansion
preserve quoted spaces, apostrophes, ampersands, parentheses, bangs,
percent/brackets and Finnish/German/emoji characters. UTF8 argv receiver records
actual Desktop5.1 host and application return 0; BAT trailing separators retain
canonical folder identity.

T039 — passed: Actual 324-character local source file and longer intended output
succeed in both desktop shells with preserved source hashes/creation/mtime and
full JPEG validation. Native 128-bit directory identities agree for the same
physical folder, differ for siblings and fail closed for missing directories.
Existing owned localhost C$ supports direct UNC input/output and BAT UNC source;
local/UNC descendant aliases reject both directions before native probing/write.
Lexical root tests do not enumerate roots.

Eight uncommitted development Pester runs are retained separately from the clean-I
acceptance runs. The existing four-suite PS7 run had 111 total/110 passed/one
failed assertion/exit 1: a Preflight test observer still indexed the owned
snapshot at argument 2 after the explicit literal registry guard shifted it to
argument 4. Correcting this test seam preserved its source, JPEG and video
assertions. Paths PS7 attempt1 had 17 total/15 passed/two failed assertions plus
one failed container/exit 1: case-insensitive $Batch shadowed the batch path
before the BAT launched, and AfterAll generic List serialization failed. Renaming
the test variable and using ToArray corrected those test harness failures. Paths
PS7 attempt2 passed 17/17; PS5.1 attempt2 had 17 total/13 passed/four
failures/exit 1 because Get-Content decoded the UTF-8 receiver JSON as ANSI. The
actual receiver bytes and output preserved the characters; .NET ReadAllText
corrected the test inspection. Both attempt3 and both attempt4 runs passed 17/17.
Attempt4 strengthened real UNC output and alias checks. These summaries' raw
end-of-run source hashes remain observations, not proof of a clean committed
revision or exact execution binding. Genuine native percent-template output and
selected-input red-versus-blue substitution, baseline BAT bang/trailing-separator
corruption, and local/UNC containment gaps were preserved as research/prototype
observations and fixed before implementation acceptance. Caller CMD paired percent
expansion remains a documented limitation. Invalid preliminary probe
argv/driver/pixel inspection and preexecution parser mistakes are qualified
separately, with no application/acceptance pass inferred.

Committed runtime/tests as 79eb1379797cc1bbe4bbcf3cc4083cf83b0e17e8.
Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444
and PS 7.6.5, Pester 5.9.1: 227/227 in each
desktop shell, zero skipped. Controls each have 228 total/227 passed/exactly
one T007 false assertion/exit 1. All four summaries and actual native child exits
bind the same clean committed checkout, with raw file hashes and separate Git blob
hashes for verified CRLF/LF normalization. Persistent execution policies are
unchanged. Actual hosted Windows Server matrices passed all four normal/control
jobs: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37259886557) and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37259892278).
Evidence: evidence/M2-T03.json and M2-T03-ci.json. Literal-name matrices, actual
launcher/native observations, capability limits, retained history and artifact
hashes are explicit. Missing UNC or disabled-environment observations cannot be
inferred from a successful local path or lexical root test.

Implementation synchronization is a past exact-SHA observation. The containing
record-only checkpoint updates STATUS/TASKS/NEXT_SESSION/SESSION_LOG and two sanitized
evidence files, plus the scoped D26 decision and primary references. Its own SHA/live sync stays pending_verification until externally
observed after normal commit/push. No raw logs, private media or local tools are tracked.
No merge, release, deployment or owner aesthetic approval was performed.

M2-T03 accepted: 12/28 tasks, T001-T039 exercised with first four characterization
and actual path/codec limits explicit. Next M2-T04 — Clarify byte caps and benchmark quality without silent default changes; stop.
The actual 324-character local source and complete longer output succeeded in both
desktop shells in this configuration. A disabled-LongPathsEnabled environment was
not configured or tested and remains not_run. Root-boundary helper cases are
lexical; no entire drive/share root was normalized.

Live UNC evidence covers only the preexisting localhost C$ alias of marked owned
fixtures: direct input/output, BAT source, and both local/UNC containment
directions. It does not establish arbitrary remote SMB servers, ACLs, network
providers or unsupported FileIdInfo implementations. Hosted live-UNC status is
recorded separately from its mandatory test count.

The BAT/application receiver has owned output/dependency/argv/application-return
seams only at the final PS1 invocation. Explorer drag-and-drop, default Pictures
routing and general final BAT exit propagation were not certified. Caller CMD
expansion before entry cannot be reconstructed; use the documented direct
PowerShell literal form for paired percent expressions.

Directory identity and reparse checks provide the documented conservative
preflight boundary, not a hostile concurrent-filesystem sandbox or universal
cross-provider equivalence proof. Source metadata stability and duplicate matching
remain length/time heuristics; same-key same-length differing content can still
skip under the declared heuristic.

Prior codec and colour limitations remain: HEIC/HEIF evidence is a two-image still
collection, not timed animation/thumbnails/auxiliaries; the ICC guard is bounded
structure/model validation, not complete semantic conformance, and the chosen CMYK
A2B0-only display profile uses the declared requested-Relative fallback. Synthetic
patch tolerances do not grant owner photographic/aesthetic approval.

M2-T04 and later byte-cap/quality, process-resource lifecycle, cancellation,
reporting and release work remain explicit future gates. No owner quality/default
approval, merge, tag, release or deployment is implied.

## 2026-10-05 — M2-T04 completed

Started clean at f9b00d8befac93c4fa1116efcbc2e9d7144f4509, independently equal to the feature remote. Owner-merged PR #12 main 36d489f61cdaef93dca3f9afde95ce7ff6484bb7 had the identical tree. Continued the authorized feature and opened successor draft PR #13 after normal push. Preserved the separate non-Git snapshot.

The default remains 1,048,576 bytes (1 MiB); positive custom caps retain their
full Int64 values. Every JPEG attempt uses invariant decimal jpeg:extent=<bytes>B.
This corrects the previous binary-PowerShell-to-decimal-native MB/KB mismatch, so
encoded bytes or quality can change when the intended budget is restored.

Validated final file length is the authoritative Int64 compliance decision. A
final JPEG at or below its byte cap is ordinary Converted success. A fully decoded
valid last 50-percent attempt above its cap is retained as ConvertedWithWarning
under the existing best-effort policy. A failed or invalid last attempt cannot
reuse a superseded candidate.

After successful no-overwrite finalization, OK IMG or WARN IMG records exact
output and target bytes, width, height and selected scale. SizeWarnings counts
only successfully finalized warning outputs and is a subset of ConvertedImages. A
retained size warning returns the existing warning/partial code 2. A superseded
trial or failed finalization cannot add a retained size warning; later compliant
output remains ordinary success.

Successful duplicate registration retains the actual Converted or
ConvertedWithWarning status with source/output links. A later duplicate of a
warning output is skipped under the existing heuristic and does not count a second
size warning. Source bytes, creation/modification times, mirrored names, video
copying and no-overwrite transactions retain the regression boundary.

The existing 100, 90, 80, 70, 60, 50 percent sequence, no upscaling, encoder
quality search, 4:2:0 sampling, progressive JPEG and
orientation/ICC/white-alpha/strip operations remain. No quality floor, sharpening,
alternative encoder, chroma setting or more aggressive resizing algorithm is
introduced or proposed. The standalone PS1/BAT interfaces and development
dependency pins remain.

The scoped test additions use actual JPEG decode/file length and real native
colour/alpha attempts alongside explicitly qualified controlled encoder-result
boundary fixtures. The representative procedural benchmark recipe separately
compares baseline/current outputs, decoded RGB8 metrics, native conversion and
complete application batch times, including an unchanged-budget 65,537-byte
control. Its actual exact-revision results, source preservation and independent
metric audit are required before records can claim task acceptance.

Initial committed full desktop and hosted gates exposed two inherited Preflight
expectations still using MB/KB strings. Only that test name and its two expected
byte operands were corrected after all initial controllers completed; runtime and
benchmark algorithm bytes remained unchanged between those implementation commits.
Initial failed gates remain historical evidence, not accepted final results.

Decision D27 and SOURCES I9 retain the algorithm and distinguish intended/native budget from actual final compliance.

T040 — passed: Five exact invariant byte-extent argv cases cover omitted default
1048576B,1024B,1537B,1048577B and Int64 maximum under ar-SA with hostile
decimal/group separators. Genuine fully decoded COM-segment JPEG encoder results
at 1048575,1048576,1048577 actual bytes prove <= equality success versus above-cap
warning independently of identical rounded 1.00 MiB displays. A 100% above-cap
trial superseded by compliant 90% result finishes OK/exit0/SizeWarnings0 with
exact retained bytes.

T041 — passed: Actual tagged-alpha native conversion with cap1 executes all
six100-to50 percent attempts and retains a fully decoded32x24 JPEG only as WARN
IMG/exit2/SizeWarnings1 with exact bytes,cap,dimensions and scale. A later
duplicate links retained ConvertedWithWarning and counts the output once.
Truncated,wrong-format and nonzero last attempts fail with Errors1/SizeWarnings0,
six distinct candidates and no stale final output.

T042 — passed: The real default conversion retains64x48 geometry with one100%
attempt. Every tested attempt keeps the existing resize sequence,4:2:0
sampling,Line interlace,white alpha removal and strip operations without new
quality/sharpen flags. Tagged retries prove live trusted target ICC bytes,Relative
intent and BPCoff. Representative corpus byte/quality/latency measurements and
owner algorithm acceptance are separate benchmark evidence; these tests claim no
changed quality floor.

Five uncommitted targeted runs are retained. The initial sole JPEG corner
threshold failure was diagnosed as tiny-cap quantization; exact-white lossless
probes and specifically calibrated lossy stress tolerance corrected the test
oracle. The first three runs used 14 cases and the final two 15 cases. No
development run is substituted for clean-I acceptance. Separate seven exploratory
benchmark tooling observations remain qualified in the benchmark manifest. The
initial committed I also failed both desktop normals at 242 total / 240 passed / 2
failed and both controls at 243 total / 240 passed / 3 failed (T007 plus the two
stale Preflight unit expectations), all native exits 1. All four hosted normal
jobs failed with 242 total / 240 passed / 2 failed and their controls were
skipped. These histories bind that initial commit, not the corrected acceptance
revision. Only one existing Preflight name and its two extent expectations were
corrected; no runtime changes or initial-I representative benchmark are claimed.

Committed runtime/tests/benchmark recipe as 0e490cc5f972b5f22374d2599eed929e477b40ca.
Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5,
Pester 5.9.1: 242 / 242 in each desktop shell, zero skipped or incomplete.
Controls each have 243 total, 242 passed, one exact T007 false assertion and
actual native exit 1. Four desktop summaries/XML/native records bind 25 source
paths, including 8 colour assets and the sizing benchmark recipe, to raw checkout
and separate committed Git blobs; differences are verified CRLF-to-LF only.
Persistent execution policies remain unchanged.

Actual hosted Windows Server push/PR matrices passed all four normal/control jobs: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37262881421) and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37262886471).
Evidence: evidence/M2-T04.json and M2-T04-ci.json.

The separate exact-I benchmark uses actual sequential PS 5.1 and PS 7 hosts, five
synthetic images plus an opaque video copy, four byte caps and the unchanged
baseline/current algorithms. All 84 rows, actual child exits, source
bytes/timestamps, output dimensions/bytes/native conversion times and RGB8
MAE/PSNR are retained with recipe and artifact hashes. Independent NumPy
recomputation checks every saved output-grid and source-grid metric pair.

Application timing is the complete five-image batch wall time at a cap and is
repeated beside each fixture row; it is not per-image application latency.
Per-fixture native conversion latency is separately measured. One repeat, one
machine and sequential hosts produce descriptive observations, not confidence
intervals, perceptual acceptance or a statistically controlled speed claim.

Correcting decimal MB/KB to exact byte budgets can change encoded bytes and
quality without a new quality/resize algorithm. The 65537B comparison is an
unchanged-budget control. No alternative encoder, scale sequence, chroma policy,
quality floor or sharpening was introduced or recommended. Owner
photographic/aesthetic/default approval is not requested or granted.

PS 5.1 seeded noise at the 1048576-byte cap — baseline: 979962 bytes, 1600 x 1200,
100%, output-grid MAE 48.076/255, PSNR 12.391 dB, native conversion 1217.0 ms,
complete batch 5825.7 ms; candidate: 1015369 bytes, 1600 x 1200, 100%, output-grid
MAE 47.816/255, PSNR 12.437 dB, native conversion 1573.4 ms, complete batch 5910.4
ms.

PS 7 seeded noise at the 1048576-byte cap — baseline: 979962 bytes, 1600 x 1200,
100%, output-grid MAE 48.076/255, PSNR 12.391 dB, native conversion 1212.8 ms,
complete batch 5978.6 ms; candidate: 1015369 bytes, 1600 x 1200, 100%, output-grid
MAE 47.816/255, PSNR 12.437 dB, native conversion 1541.6 ms, complete batch 6309.2
ms.

Implementation synchronization is a past exact-SHA observation. This containing record-only checkpoint updates exactly STATUS/TASKS/NEXT_SESSION/SESSION_LOG, two sanitized evidence files, D27 and I9. Its own SHA/live sync remains pending_verification until externally observed after normal commit/push. No raw media/private paths/native argv/local tools are tracked; no merge, release, deployment or owner aesthetic/default approval was performed.

M2-T04 accepted: 13 / 28 tasks, T001-T042 exercised with first four characterization and prior codec/path limits explicit. Next M2-T05 — Capture useful errors and remove indiscriminate fallback; stop.

ImageMagick parses its extent through double and may round accepted Int64 budgets
above 2^53 internally. The exact argv regression does not claim massive-file
encoder precision. The validated final Int64 file length comparison remains
authoritative regardless of native search approximation or exit success.

Genuine JPEG COM padding controls encoder-result length for byte boundary
regressions; it does not establish metadata privacy or real resize quality.
Separate actual colour/privacy regressions retain their assertions. Exact-white
lossless probes and the specifically calibrated 24/255 JPEG corner tolerance apply
only to the impossible one-byte synthetic stress case, not a general perceptual
quality floor.

Representative benchmark rows are synthetic encoded-sRGB measurements on one
machine with one repeat per runtime/cap. Complete five-image batch wall time is
not per-image application latency; native conversion timing is separate.
Sequential host runs, tracing and validation overhead limit timing interpretation.
No statistical speed claim, confidence interval, arbitrary photographic/print
accuracy or subjective owner quality/default acceptance is granted.

Existing BAT pause/completion behavior and general final BAT exit propagation
remain future M3-T05 work; application warning code 2 does not certify a changed
launcher exit contract. Explorer/default Pictures routing is not inferred from
owned launcher tests.

Prior path/provider and codec limits remain qualified: local/localhost alias tests
do not certify arbitrary remote providers, disabled long-path policy is not
configured, HEIC evidence is a two-image still collection rather than timed
animation, and the selected CMYK A2B0 mapping is not general relative-colorimetric
printing accuracy. The ICC structure/model guard is bounded validation, not
complete semantic conformance.

Length/time stability and duplicate matching remain documented heuristics rather
than runtime content identity. General bounded diagnostic classification,
justified retries, process/resource lifecycle, cancellation, reporting and release
work remain later tasks. No owner merge, tag, release, deployment or
aesthetic/default approval is implied.

No new encoder/scale/chroma/default algorithm is proposed. Correcting decimal
MB/KB to exact byte budgets can change default or binary-divisible output
bytes/quality; 65537B provides an unchanged-budget comparison.

No arbitrary photography, print accuracy, owner aesthetic/default acceptance or
cross-machine performance claim.

The public function return is captured; the invoking parent must separately
preserve the fresh host native process exit. Raw paths, media and transcripts
remain ignored.

Hosts were measured sequentially after local Pester completion. One repeat per
machine/runtime/cap is descriptive only; no confidence interval or statistically
controlled speed claim.
