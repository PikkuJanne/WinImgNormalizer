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

## 2026-10-05 — M2-T05 completed

Started clean at 2e30cf9948669f7af32533efc9d842472693be6d, equal to the live feature
branch; owner-merged PR #13 main 590241944033c1332105e21a9dab49578a939941 had the same
tree. Continued without changing checkout/history. Successor draft PR #14 contains this
task. Implementation synchronization is a past exact-SHA observation; this record-only
checkpoint's SHA/live synchronization remains pending_verification until observed
externally after normal commit/push.

A native failure now stops that image at its current scale even if the encoder leaves a
decodable JPEG. Valid later files continue. Only an exact known JPEG notice with native
exit 0 and full validation can be retained as ConvertedWithWarning; it adds
NativeWarnings within ConvertedImages and returns application code 2. SizeWarnings
remains a separate retained-output attribute.

Stdout and stderr drain independently with bounded retained text; truncation fails. Only
explicit sharing/lock failures can retry twice total per image, after 100/200 ms, at the
same scale with identical flags and fresh owned scratch. Six smaller-size attempts are
reached only from valid above-cap output, giving at most eight conversions including
retries. Source/video preservation, no-overwrite safety, colour/white-alpha/strip
policy, exact byte cap and JPEG settings remain.

T043 — passed: Both actual Windows hosts capture native stdout/stderr/exit independently
under ErrorActionPreference Stop. Genuine pinned lossless-to-lossy warning provenance
distinguishes exit0 without regard-warnings from nonzero with regard-warnings, both with
decoded JPEGs. Controlled native re-emission after actual production conversion
finalizes explicit warning/exit2 with valid white/red JPEG and warning duplicate status;
nonzero, truncated, unknown or ICC-mixed outcomes cannot finalize.

T044 — passed: Actual owned unknown-coder, truncated-JPEG and exclusive-lock diagnostics
classify permanently; the native Permission denied text remains nonretryable. Controlled
native missing-codec, damage, denied, resource and unknown failures stop once despite
valid JPEG bytes. Only explicit sharing text or trusted Win32 HRESULT 32/33 retries;
another facility with low word 32 stops. Two 100/200 ms-backoff retries per file
preserve scale/ICC/alpha arguments and fresh candidates. The actual eight-call case
contains two controlled sharing failures and six real valid above-cap size conversions,
ending in a 541-byte 32x24 JPEG/size warning/exit2 with source/video preservation.

T045 — passed: Two actual concurrent newline-free child streams each emitted 4194330
characters and native exit9; retained prefix/tail were bounded 1024 characters per
stream, counts exact and streams complete without deadlock. At production capture
limits, a controlled zero-exit native process that wrote a valid JPEG while overflowing
streams is rejected once as OutputLimit; useful tail details remain in the bounded log
and console output stays below 4096 characters.

The initial Diagnostics target passed 10/23 and failed 13 because the owned managed
fixture could not handle generated long paths and the pixel oracle used invalid FX
syntax. Some rejection assertions could pass for the unintended fixture crash, so later
tests require the intended native exit, complete streams and actual candidate. Corrected
intermediate targets passed 26/26 in both shells; final targets passed 27/27 with the
actual eight-call boundary. Separately, both first focused runs passed 131/132: the
invalid-curve native work was safely rejected but the new category assertion expected
DamagedInput. A narrow classifier reason correction produced 132/132 reruns. Runtime
edits overlapped those first focused runs, so their end hashes do not prove immutable
import bytes.  The first clean committed implementation failed each desktop normal
268/269 and control 268/270 (T005 plus deliberate T007), and all four hosted normals
failed 268/269 with control steps skipped. T005 still expected an unexplained exit 9 to
retry. A three-line test-only correction supplied an explicit transient reason; both
focused Normalizer reruns passed 6/6. Runtime bytes were unchanged. The original commit,
raw source snapshots, summaries/XML/native records, hosted logs and independent
diagnosis are retained. Final acceptance uses a separate clean corrected revision.
Tooling-only inspection mistakes contribute no application or Pester acceptance counts.
The second clean committed implementation passed all four desktop gates and both hosted
matrices (269/269 normal; 270 total, 269 passed and exact T007 control). A separate
actual native result-shape probe then found that the unchanged development measurement
trace recorded absent DiagnosticOutput as null, bypassing its strict diagnostic guard.
Its narrow adapter now records bounded separate streams with legacy fallback and returns
the original result unchanged. Actual saved-AST checks passed five cases per host, with
full JPEG decode and genuine warning/quiet cases. Those probes observed I2 HEAD plus the
edited tool, whose raw bytes match final I3; they are not relabeled as clean-I3 runs or
a representative benchmark. I2 successful gates, all 26 raw source snapshots, original
CI/merge/review and both omission/correction proofs remain preserved. Final task
acceptance uses fresh full/hosted gates at I3.

Tested clean implementation: 5ddbf4156f2c10534a045586a5b5e5e050255c30.
Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1: each normal run passed 269/269 with native exit 0 and zero skipped or incomplete
tests. Each control had 270 total, 269 passed, one exact T007 deliberate failure and
native exit 1. Summary/XML, parent native records and console hashes bind 26 tested
source paths to raw checkout and separate Git blobs, with only verified CRLF-to-LF
normalization where needed. Persistent execution policies were unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37267178588); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37267182413).
Evidence: evidence/M2-T05.json and evidence/M2-T05-ci.json.

Accepted 14/28 tasks and exercised T001-T045, with the first four characterization cases unchanged. Next: M2-T06 — Close the conversion-correctness milestone. This containing checkpoint changes only handoff/evidence and reviewed D28/I10 when present; its own synchronization remains pending_verification. Raw media/private paths/argv remain ignored. No owner merge, release, deployment or subjective quality approval is inferred.

General per-file deadlines, cancellation, native descendant lifecycle and final
reporting/BAT exit propagation remain later M3 gates. Fixed character capture is not a
CPU/disk/lifetime sandbox. Controlled warning, retry and process-local resource fixtures
remain qualified; prior codec/colour/provider limits and owner quality/default approval
are unchanged. Full limitations and raw artifact bindings are in evidence/M2-T05.json.

## 2026-10-05 — M2-T06 completed; M2 milestone closed

Started clean at c062b835c4acb2734f42f87d81b8e47164af827a, equal to the live feature
branch; owner-merged PR #14 main e24461bb22155902984784aba7e30d4db0e63147 had the same
tree. Continued without changing checkout/history. Successor draft PR #15 contains this
task. Implementation synchronization is a past exact-SHA observation; this record-only
checkpoint's SHA/live synchronization remains pending_verification until observed
externally after normal commit/push.

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

T046 — passed: Both default1048576-byte and exact65537-byte actual unmocked mixed-tree
cases passed in every fresh PS5.1/PS7 clean-I normal/control invocation, plus focused
development. Eight exact-I comparison runs bind27 gate source files and actual
tool/fixture/output/log/buffer bytes; only T007 deliberately failed in controls.
All12JPEG+1video+1duplicate per normalization and mandatory
frame/ICC/privacy/literal-path/exact-cap/source/prior-run checks passed.

Focused development run 1, .scratch/M2-T06-target-ps7-1: 0/2 passed, 2 failed, actual
child exit 1; incomplete counts {'skipped_count': 0, 'not_run_count': 0,
'inconclusive_count': 0, 'failed_blocks_count': 0, 'failed_containers_count': 1}. Raw
summary/XML and case-note artifacts remain preserved. These precommit observations are
separate from clean-I acceptance.  Focused development run 2,
.scratch/M2-T06-target-ps7-2: 0/2 passed, 2 failed, actual child exit 1; incomplete
counts {'skipped_count': 0, 'not_run_count': 0, 'inconclusive_count': 0,
'failed_blocks_count': 0, 'failed_containers_count': 1}. Raw summary/XML and case-note
artifacts remain preserved. These precommit observations are separate from clean-I
acceptance.  Focused development run 3, .scratch/M2-T06-target-ps7-3: 0/2 passed, 2
failed, actual child exit 1; incomplete counts {'skipped_count': 0, 'not_run_count': 0,
'inconclusive_count': 0, 'failed_blocks_count': 0, 'failed_containers_count': 1}. Raw
summary/XML and case-note artifacts remain preserved. These precommit observations are
separate from clean-I acceptance.  Focused development run 4,
.scratch/M2-T06-target-ps7-4: 0/2 passed, 2 failed, actual child exit 1; incomplete
counts {'skipped_count': 0, 'not_run_count': 0, 'inconclusive_count': 0,
'failed_blocks_count': 0, 'failed_containers_count': 1}. Raw summary/XML and case-note
artifacts remain preserved. These precommit observations are separate from clean-I
acceptance.  Focused development run 5, .scratch/M2-T06-target-ps7-5: 0/2 passed, 2
failed, actual child exit 1; incomplete counts {'skipped_count': 0, 'not_run_count': 0,
'inconclusive_count': 0, 'failed_blocks_count': 0, 'failed_containers_count': 0}. Raw
summary/XML and case-note artifacts remain preserved. These precommit observations are
separate from clean-I acceptance.  Focused development run 6,
.scratch/M2-T06-target-ps7-6: 0/2 passed, 2 failed, actual child exit 1; incomplete
counts {'skipped_count': 0, 'not_run_count': 0, 'inconclusive_count': 0,
'failed_blocks_count': 0, 'failed_containers_count': 0}. Raw summary/XML and case-note
artifacts remain preserved. These precommit observations are separate from clean-I
acceptance.  Focused development run 7, .scratch/M2-T06-target-ps7-7: 0/2 passed, 2
failed, actual child exit 1; incomplete counts {'skipped_count': 0, 'not_run_count': 0,
'inconclusive_count': 0, 'failed_blocks_count': 0, 'failed_containers_count': 0}. Raw
summary/XML and case-note artifacts remain preserved. These precommit observations are
separate from clean-I acceptance.  Focused development run 8,
.scratch/M2-T06-target-ps7-8: 2/2 passed, 0 failed, actual child exit 0; incomplete
counts {'skipped_count': 0, 'not_run_count': 0, 'inconclusive_count': 0,
'failed_blocks_count': 0, 'failed_containers_count': 0}. Raw summary/XML and case-note
artifacts remain preserved. These precommit observations are separate from clean-I
acceptance.  Focused development run 9, .scratch/M2-T06-target-ps51-1: 0/2 passed, 2
failed, actual child exit 1; incomplete counts {'skipped_count': 0, 'not_run_count': 0,
'inconclusive_count': 0, 'failed_blocks_count': 0, 'failed_containers_count': 0}. Raw
summary/XML and case-note artifacts remain preserved. These precommit observations are
separate from clean-I acceptance.  Focused development run 10,
.scratch/M2-T06-target-ps51-2: 2/2 passed, 0 failed, actual child exit 0; incomplete
counts {'skipped_count': 0, 'not_run_count': 0, 'inconclusive_count': 0,
'failed_blocks_count': 0, 'failed_containers_count': 0}. Raw summary/XML and case-note
artifacts remain preserved. These precommit observations are separate from clean-I
acceptance.  The first four focused PS7 development runs each discovered two cases and
passed zero, failed two, recorded one failed BeforeAll container and returned native
host exit 1. They were fixture/setup failures before application execution and are not
acceptance runs.  Run 1: native multi-image fixture generation returned zero but
numbered the output instead of writing the intended literal source name. The generator
was changed to neutral owned files followed by literal byte copies into the declared
source tree; actual image count, geometry and pixels remain asserted.  Run 2: native
CRLF separators were compared to an LF-only expected string. Only comparison text
normalization changed; the original native process records remain retained.  Run 3: the
source identification assertion saw HEIC where the fixture expected HEIF. The initial
interpretation was an equivalent container-label alias, but retrospective inspection of
the retained native source/arguments shows this was the extensionless primary-heif
fixture caused by the missing concatenation operator later diagnosed in run 5. The
original notes are preserved with that initial interpretation; the final notes correct
it. The repaired literal .heic/.heif sources separately assert both images'
geometry/content and the runtime selection labels. This remains a two-image still
collection, not evidence of timed animation.  Run 4: the fixture helper incorrectly
prepended file-literal options to the standalone version query, which made ImageMagick
require an output filename. The version-only query is executed directly; file operations
retain explicit literal handling.  Run 5 reached two real application batches, each
returning 0, but the independent summary assertion failed before the later comparison
assertions completed. A missing concatenation operator omitted the HEIC/HEIF extensions
from two fixture names, so the media inventory excluded them. The fixture names were
repaired and a precondition now checks every declared image exists with an eligible
extension. The expected unsupported count was also corrected to the established
inventory policy: unrelated text files are excluded, preserved and checked in the
complete source state. No runtime behavior changed and no assertion of required image
content was removed.  Run 6 reached both actual application batches with 12 converted
images, one copied video, one duplicate and exit 0. The independent BMP sample assertion
then failed because PowerShell flattened the one-element nested expected/point arrays;
the actual native sample was 32/208/65. Unary-comma construction repairs that singleton
tuple, and preconditions now verify every expected RGB tuple and coordinate pair has the
declared shape. The colour tolerance is unchanged.  No runtime/default change, completed
independent quality comparison or immutable-I acceptance is inferred from these
development failures. Their actual summaries, XML, consoles, native outcomes, partial
comparison records and original source snapshots are retained in ignored scratch and
bound by the case notes.  PS7 run 7 and PS5.1 run 1 each discovered two cases, passed
zero, failed two, recorded no failed container and returned native exit 1. Both actual
application batches returned 0 and produced the required mixed corpus. The JPEG
structure assertion incorrectly required three RGB components for genuinely Gray seeded
noise, whose progressive JPEG correctly has one luminance component. The oracle now
permits that one-component SOF2 form only for the declared Gray noise fixture; coloured
outputs still require progressive three-component RGB 4:2:0. No runtime, encoder default
or colour tolerance changed.  PS7 run 8 and PS5.1 run 2 each passed both focused cases
with native exit 0, no skips and no failed containers. Their saved RGB buffers and
descriptive quality metrics were independently recomputed. These are
development/frozen-source observations, preserved separately from the subsequent clean
implementation-revision full-suite and deliberate-failure-control acceptance gates.
Focused counts above are actual retained development results and remain qualified by
their original source observations. Immutable-I desktop normal/control and hosted
push/PR gates are separate acceptance evidence supplied by root after execution.
Previous M2 task histories and benchmarks remain bound to their own task revisions; this
card does not relabel them.

Tested clean implementation: 11ee247154a00d79aaa49a438d1f526457b22fdf.
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

Accepted 15/28 tasks and exercised T001-T046, with the first four characterization cases unchanged. Next: M3-T01 — Recover from inaccessible folders and expose ancillary failures. This containing checkpoint changes only six handoff/evidence files; its own synchronization remains pending_verification. Raw media/private paths/argv remain ignored. No owner merge, release, deployment or subjective quality approval is inferred.

JPEG derivatives are lossy and do not replace preserved originals. The comparison
evidence applies to the approved synthetic corpus and pinned decoder/profile behavior;
codec/build and reference-intent qualifications remain explicit. Inaccessible subtrees
and ancillary reporting are next M3-T01; general deadlines, cancellation and native
descendant lifetime remain later M3 gates. Owner acceptance, releases and deployment
remain separate.

## 2026-10-05 — M3-T01 completed

Started clean at 5cbaa2246923f5bd94066feea34f71e0faafce00, equal to the live feature
branch; owner-merged PR #15 main 446683923f7d4b2aed7a5c9a00e1bf11d2924b68 had the same
tree. Continued without changing checkout/history. Successor draft PR #16 contains this
task. Implementation synchronization is a past exact-SHA observation. This record-only
checkpoint's SHA/live synchronization remains pending_verification until separately
observed after normal commit/push.

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

Actual development .scratch/M3-T01-target-ps7-1: 15/16 passed, 1 failed, native exit 1.
Only file-symlink fixture precondition failed before that traversal invocation. Default
P/Invoke bool marshaling falsely reported success for a BOOLEAN return; no link existed.
Byte-return MarshalAs(I1) probe established Win32 1314 privilege absence. All other 15
cases passed; original result and bytes retained.  Actual development
.scratch/M3-T01-target-ps7-2: 16/16 passed, 0 failed, native exit 0. Details and
original source observations are retained in the case notes.  Actual development
.scratch/M3-T01-target-ps51-1: 16/16 passed, 0 failed, native exit 0. Details and
original source observations are retained in the case notes.  Actual development
.scratch/M3-T01-target-ps7-3: 16/16 passed, 0 failed, native exit 0. Details and
original source observations are retained in the case notes.  Actual development
.scratch/M3-T01-target-ps51-2: 16/16 passed, 0 failed, native exit 0. Details and
original source observations are retained in the case notes.  Actual early focused
history remains in the original M3-T01-recovery-case-notes.json with original
summaries/XML/native/console and raw runtime/suite snapshots. PowerShell 7 attempt 1
discovered 16 tests, passed 15 and failed 1, with native exit 1. Its only failure was
the T048 file-symlink fixture precondition: default P/Invoke bool return marshaling did
not match native one-byte BOOLEAN and falsely reported success while no link existed. A
separately preserved byte-return probe established false/Win32 1314 (required privilege
not held). The narrow MarshalAs(I1) test correction left runtime bytes unchanged.
Corrected PS7 attempt 2 and fresh PS5.1 attempt 1 each passed 16/16 with native exit 0,
no incomplete tests and unchanged policies. These early focused results are separate
from committed acceptance gates and remain bound to their original source snapshots.
Initial committed implementation 9b95901f96afe033dd2b4836e4fdb15d0324a0e7 then ran four
actual desktop gates. Each normal discovered 287 tests, passed 284 and failed 3, with
native exit 1; each control discovered 288 tests, passed 284 and failed 4, with native
exit 1. The normal failures were both T006 positional variants expecting the obsolete
Completed log label, and the old T016 arriving-log assertion expecting setup code 1/no
copied video instead of the newly documented warning code 2/retained media policy.
Controls additionally failed the intended T007. All four summaries/XML/native/console
and 28 raw source snapshots were captured under M3-T01-initial-I-committed-history
before repair. These failed controls are not accepted runner-credibility gates because
they contain additional failures.  Automatic initial-I push 37276366241 and PR
37276387760 completed on both hosted shells. Each normal discovered 287 tests, passed
282 and failed 5, with native exit 1; the later control steps were skipped. The same
three stale assertions failed, plus both T047 fixture restoration comparisons: Set-Acl
added the DACL auto-inherited control flag (D:AI) to an original descriptor lacking it,
despite matching ACEs. The application denied-enumeration and sibling/outcome assertions
preceded this fixture finally-restoration failure. Original four job transcripts and the
original collection bytes remain preserved; no rerun or dispatch was substituted.  The
narrow test repair keeps complete T006 media/default/source/timestamp checks and asserts
Processing ended plus actual final clean reporting state. The arriving-log test now
requires warning code 2, one copied byte-identical video with original timestamps, one
log warning and authoritative degraded final state, exact unchanged foreign log
bytes/content/length, empty owned work and no TEMP fallback. Runtime remains frozen.
Recovery T047 retains exact ACL equality: owned native DACL probes on both desktop
shells compare full descriptor/SDDL/DACL bytes and control flags after real denial, for
natural and explicitly no-auto-inherited originals. The fixture restores original binary
DACL with SetFileSecurityW(mask 4) only when the original lacks DACL_AUTO_INHERITED; the
natural auto-inherited branch retains Set-Acl. The denied-only case establishes the
no-AI baseline before recording source state. This is fixture cleanup, not application
permission recovery or a change to global security policy.  Separate tiny peer exception
probes and fresh-host controlled PipelineStopped helper observations establish no
ordinary continuation and independently written finally state; native zero alone is not
completion. They do not establish real Ctrl+C, native deadlines, cancellation or
descendant cleanup. Optional actual file-symlink absence on the desktop remains
qualified; hosted logs do not establish a positive file-symlink capability. The first
prose drafts remain preserved as planning history. Final repaired focused histories and
subsequent clean-I normal/control plus automatic hosted matrices are bound separately
through actual case notes and evidence. Automated results do not grant owner acceptance.
Development end-of-run hashes are observations, not proof of immutable imported bytes.
Final clean-I desktop normal/control and hosted matrices are separate acceptance gates.
Prior card histories remain in their original evidence.

Tested clean implementation: 66f609ad32637d82ce1e9918ea1600400378422e.
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

Accepted 16/28 tasks and exercised T001-T050, with the first four characterization cases unchanged. Next: M3-T02 — Bound processing and own native process lifetime. This containing checkpoint changes only 6 reviewed handoff/evidence records; its own synchronization remains pending_verification. Raw media/private paths/argv remain ignored. No owner merge, release, deployment or owner approval is inferred.

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

## 2026-10-05 — M3-T02 completed

Started clean at fa4ab9fd609918500a5dcec1c5972c24d23c2c37, equal to the live feature
branch; owner-merged PR #16 main a2ef99ac52631e84988669905aa8e80d44682b79 had the same
tree. Continued without changing checkout/history. Successor draft PR #17 contains this
task. Implementation synchronization is a past exact-SHA observation. This record-only
checkpoint's SHA/live synchronization remains pending_verification until separately
observed after normal commit/push.

An image now shares a 120,000-ms elapsed budget across its snapshot, frame/colour/ICC
queries, conversion attempts, retries and complete JPEG validation. The final move
requires time remaining. Every real image command requests process-local pixel-cache
memory 512 MiB, map 1 GiB, disk 2 GiB, two threads and the remaining seconds; stricter
installed ImageMagick policies still apply. A timeout or resource failure rejects that
image and continues later usable image/video siblings with application warning/partial
code 2.

Native processes are created atomically inside a private kill-on-close Windows job and
start only after two bounded stream readers are ready. Cleanup targets that job,
confirms an empty tree and drains/cancels readers within one 3-second grace; it never
selects unrelated processes by name or PID snapshots. The separate exclusive plain TEMP
cache root is validated before output writes and kept outside the physical source tree.
Media snapshots, ICC data and JPEG candidates keep output-volume staging. Unconfirmed
tree termination rejects finalization and preserves all item scratch; ordinary cleanup
is nonrecursive, removes only regular magick-* cache files and preserves unexpected
entries with a warning.

D30/I12 record the implementation policy and dated official Windows/ImageMagick advisory
review. Public positional/BAT calls, exact size cap, six scales,
JPEG/colour/white-alpha/frame selection, literal paths, no-overwrite transactions,
source preservation, video bytes and successful duplicate retention remain within their
existing contracts. The development Measure-SizeQuality trace now follows the current
bounded native observer and separates leading resource-limit arguments from core image
flags. Tiny compatibility checks are separate from representative quality/performance
measurement and owner approval.

T051 — passed: Actual wrapper/job root-child-grandchild timeout and normal-root cleanup;
independent direct-parent plus creation-identity and held-file proof; flood and valid
JPEG before stall remain unacceptable; actual selected phase sleepers share a finite
file budget and later real media complete.

T052 — passed: Actual benign large-cache native exhaustion under lower process-local
memory/map/disk/thread ceilings; later tiny JPEG/video complete. Child-only environment
and stricter private policy preserve installed policy/parent environment. Short cache
ownership/path/source boundary and exact cleanup/refusal assertions pass.

T053 — passed: An independently started actual pinned ImageMagick process stays alive
with the same PID/start identity through another private job timeout. After a three-byte
synchronization write it exits zero and produces exactly one full-decoded 1x1 red PNG
from the independent first xc image; every decoded stdin image is discarded. The test
does not assert received stdin byte count or alter host encoding.

Actual development .scratch/M3-T02-target-ps7-1: 4/14 passed, 10 failed, native exit 1.
Fixture switch Input collided with the PowerShell automatic enumerator variable. Large
switch shadowed the large template path. GetNewClosure isolated callback command
resolution from locally imported functions. The combined top-level limit/list resource
probe is invalid on the pinned native CLI; identify subcommand fixes its actual
MissingArgument exit11.

Actual development .scratch/M3-T02-target-ps7-2: 8/14 passed, 6 failed, native exit 1.
Immediate normal-root cleanup correctly returned without timeout; the original
unconditional timeout expectation was incorrect. Actual generated ImageMagick cache
filenames beneath deep output work exceeded MAX_PATH, blocking later images and the
requested validation phase. The later sibling/cache/cumulative application assertions
failed honestly; source file hashes and times remained preserved.

Actual development .scratch/M3-T02-target-ps7-3: 13/17 passed, 4 failed, native exit 1.
Three strict directory-mtime baselines differed by milliseconds around fixture creation;
source hashes and file creation/modified times remained intact. Subsequent fixtures set
all directory times after writes, deepest first, and snapshot objects are refreshed;
exact equality remains mandatory. Source-contained cache rejection occurred after the
destination write probe and changed output-parent directory mtime. Runtime validation
moved before that probe; the original no-write assertion remains intact.

Actual development .scratch/M3-T02-target-ps7-4: 17/17 passed, 0 failed, native exit 0.
All actual assertions passed with complete summary/XML and native exit0.

Actual development .scratch/M3-T02-target-ps51-1: 17/17 passed, 0 failed, native exit 0.
All actual assertions passed with complete summary/XML and native exit0.

Actual development .scratch/M3-T02-target-ps51-2: 17/17 passed, 0 failed, native exit 0.
The narrow unrelated-process fixture repair passed: real ImageMagick stays alive with
the same creation identity during another private job timeout, then produces exactly one
independently generated red PNG after stdin synchronization.

Actual development .scratch/M3-T02-target-ps7-5: 17/17 passed, 0 failed, native exit 0.
The narrow unrelated-process fixture repair passed: real ImageMagick stays alive with
the same creation identity during another private job timeout, then produces exactly one
independently generated red PNG after stdin synchronization.

Actual M3-T02 focused PS7 development runs retained 14 total/4 passed/10 failed, then
14/8/6, then 17/13/4, each native exit 1 with zero excluded/container/block counts.
These are 20 failed assertions across three real runs, with fixture errors and genuine
runtime defects distinguished below. Initial-I1-focused PS7 attempt 4 and PS5.1 attempt
1 each passed 17/17 with actual native exit 0 and no incomplete tests. Their raw source
snapshots, summary/XML/native/console records and owned observations are preserved; no
earlier result was overwritten or relabeled as a clean implementation gate.

The first focused run exposed fixture Input/Large names shadowing automatic variables or
paths, GetNewClosure command-resolution isolation, and an invalid combined converter
limit/list-resource query. The actual native query returned MissingArgument/exit 11 even
though its exploratory script returned 0; the corrected identify query is the valid
route. Separate initial API/resource probes also retain a child Unicode-console mismatch
and query error. These are test/probe preparation failures, not invented runtime
success.

The second focused run incorrectly expected every normal-root/remaining-descendant
cleanup to time out; prompt normal-root cleanup was correct and that assertion was
repaired. It also exposed a genuine runtime path defect: ImageMagick's generated cache
files below deep output work exceeded ordinary path limits. An independently preserved
always-extended TEMP/TMP/MAGICK_TEMPORARY_PATH probe likewise failed native execution; a
later plain short cache-anchor probe completed native 0. Production now allocates an
exclusive WinImgNormalizer-cache-GUID directory beneath a validated existing process
TEMP root rather than depending on an extended cache prefix or universal 8.3 support.
The frozen early design artifact's output-tree cache description is superseded by this
final policy, with original source/probe bytes retained.

The third focused run retained three strict directory-mtime failures from fixture
creation order/cached metadata; fixtures now set directory times after all writes,
deepest first, and refresh snapshot objects while retaining exact equality. It also
exposed a genuine setup-order defect: source-contained native TEMP rejection occurred
after the destination write probe and changed the output-parent directory timestamp.
Root moved the physical temporary-root validation before that probe; the no-output-write
assertion remained intact. No source/hash/time preservation assertion was removed.
Initial-I1-focused executable runtime, suite and runner bytes stayed fixed through those
two green focused children; a later T053 test-only repair is recorded separately below.
Tests README prose/Markdown/UNC-sentence restoration happened afterward and is
explicitly bound to its frozen final SHA rather than claimed as the earlier focused
document bytes.

The measurement compatibility work retained an initial native-exit-1 tiny square-fixture
run: 128x128 landscape/portrait fixtures had identical bytes and made fixtureByHash
mapping ambiguous. The trace also exposed its existing automatic Input variable
collision. A rectangular 128x160 synthetic corpus and nativeInputs trace normalization
corrected the tooling. A first literal writer asserted the wrong substring count and
exited before editing. Successful fresh PS5.1 and PS7 smokes each returned native 0 with
ten actual image conversions and two opaque-video rows, preserving source/harness bytes
and policy. Both successful native records bind the same final runtime raw SHA256
791e8484851f13e6a340374cd6e913cdffa95e32e126de77a5ca4b6031357f3e and adapter SHA256
f7cf248119cb8c8f50a7cc891431ea04691e370b3bb2e5acf4f9152558208b94. An initial reviewer
message and the immutable adapter audit's first limitation mistakenly said the PS5.1
smoke preceded the setup-order correction; actual equal hashes and the final exact-I
review supersede that chronology. The original PS5.1 controller prose says 128x128 while
its saved dimensions/argv prove 128x160; original artifacts stay unchanged. These twelve
rows per host establish a compatibility smoke, not a representative benchmark rerun or
quality/performance result.

The reviewer authored the native C# block and measurement adapter under explicit root
delegation; root independently reviewed those blocks. The final exact-I peer artifact
discloses this boundary and does not present author review as independent third-party
authorship review. The saved current upstream advisories and tagged source/API bytes are
primary provenance, not malicious-exploit execution. Controlled unconfirmed-tree
preservation changes a result flag only after a real successfully killed tree; it is
separate from actual owned-tree timeout and actual unrelated-process survival.

Initial committed implementation 66780bb8f3584f0fd1b38edf55d71af53209f77a (I1) passed
both actual desktop normal suites 304/304 with native exit 0. Both desktop controls had
305 total, 304 passed and exactly T007 deliberate failure/native exit 1, with zero
incomplete tests. The four completed gates and all 29 raw tested-source snapshots are
preserved in M3-T02-initial-I-committed-history.json, with separate Git blob bindings.
These are genuine passing but superseded desktop observations, not final
repaired-revision acceptance.

Both initial automatic hosted workflows failed: push 37284113766 and PR 37284122322.
Each PS7 job passed normal 304/304 and control 305/304/one T007. Each PS5.1 normal job
instead had 304 total, 303 passed, one T053 failure/native exit 1; its control workflow
step was skipped after that failure and is not an observed passing control or a skipped
mandatory test. The unrelated actual ImageMagick process had passed
survival/PID-creation checks and returned native 0, but the expected single
unrelated.png was missing. Original metadata/raw logs, unchanged original collector JSON
and separately enriched actual I1 merge proof remain retained. No hosted success is
inferred from the passing desktop run or from native zero alone.

The original byte-level diagnosis reproduced a BOM appended during Close with an
explicit AutoFlush=false writer and suggested closing BaseStream directly. That
close-only suggestion is superseded by the preserved authoritative supplement. Microsoft
Framework Process.Start creates redirected StandardInput as a StreamWriter using
Console.InputEncoding and immediately enables AutoFlush. AutoFlush can emit an encoding
preamble before any raw payload; closing BaseStream cannot remove an already emitted
BOM. Actual forced-BOM AutoFlush controls on both shells received EF-BB-BF before
E0-20-20, and the original raw-only recipe produced two numbered PNGs at native exit 0.
Original hosted stdin encoding/bytes and numbered output files were not captured, so BOM
causation for those two hosted T053 failures remains a supported inference, not a
directly observed hosted byte dump or a production tree-isolation defect.

The narrow test-only repair supplies a declared xc:#E02020 image first, uses 1x1 RGB
stdin only as blocking synchronization, deletes all stdin-derived frames with -delete
1--1, then closes the binary pipe after writing and flushing the intended three bytes.
Its output oracle is the declared first image, not the raw stdin payload. Actual default
and forced-BOM candidate probes passed in fresh PS5.1 and PS7: the unrelated process was
blocked before payload, then returned native 0 and produced exactly one independently
Pillow-decoded 1x1 PNG with RGB 224/32/32. The test retains same process creation
identity/aliveness through owned-tree timeout and exact one-image/pixel checks.
Subsequent repaired focused, committed desktop and hosted gates must remain separately
revision-bound; these byte probes are not replacement acceptance gates. This final-ready
draft does not infer repaired full/hosted acceptance or a containing checkpoint
synchronization observation.

The byte-probe tooling originally accessed StandardInputEncoding, unavailable on .NET
Framework, after its first branch had run; the corrected explicit writer controls and
original script/artifacts are preserved. Two later PS5.1 candidate-probe launches failed
only because inherited PSMODULEPATH prevented Get-FileHash discovery. The first
process-only removal used case-sensitive lookup and missed the uppercase key; the
corrected fresh child excludes it case-insensitively without changing parent or
persistent environment. Both failed native records/logs and the successful corrected
candidate controls remain separately retained. Completed-job gh run view --log was
unavailable while the matrix was active, so completed-job API retrieval saved the actual
raw logs; later whole-matrix collection is separate. These are tooling limitations, not
runtime/process isolation outcomes. The root-owned final PR body should explicitly state
I1's passing desktop versus failed hosted result and the qualified fixture repair.
Accepted final-I gates will be reported separately, and the containing record checkpoint
synchronization must remain pending_verification until externally observed.

Corrected implementation 73bb2764ecf3ce3bf0fcb9e82feec434005f1e4a (I2) is actually
committed and pushed. Relative to I1, only the T053 synchronization fixture changes
(seven inserted and three deleted test lines); runtime raw SHA256
791e8484851f13e6a340374cd6e913cdffa95e32e126de77a5ca4b6031357f3e is unchanged. Fresh
focused PS7 attempt 5 and PS5.1 attempt 2 each passed all 17 cases with actual native
exit 0. Their corrected v2 notes retain all seven focused histories and the original
note hash, and the separate exact-I2 peer review binds this test-only scope. At v3 draft
preparation the M3-T02-attempt2 desktop/full-control children and I2 automatic push/PR
workflows were pending. They have since closed: both clean-I2 desktop normal suites
passed 304/304 with actual native exit 0; each control recorded 305 total, 304 passed,
exactly one intended T007 failure and native exit 1. All incomplete/skip counts are zero
and persistent policies unchanged. The exact-revision validator binds all 29 raw
tested-source paths to separate Git blobs. Both automatic I2 hosted push and PR matrices
passed all four Windows PS5.1/PS7 jobs with the same normal/control counts, actual
native exits and hashed raw transcripts. Those separate immutable gates establish
repaired-revision acceptance; no earlier failure or probe is relabeled.

Development end-of-run hashes are observations, not proof of immutable imported bytes.
Final clean-I desktop normal/control and hosted matrices are separate acceptance gates.
Prior card histories remain in their original evidence.

Tested clean implementation: 73bb2764ecf3ce3bf0fcb9e82feec434005f1e4a.
On actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1, each normal suite passed 304/304 with native exit 0 and zero skipped or
incomplete tests. Each control had 305 total, 304 passed, exactly one T007 deliberate
failure and native exit 1. Summary/XML, parent native records and console hashes bind 29
tested source paths, including the native lifetime suite and all eight colour/reference
assets, to observed raw bytes and separate Git blobs. Only verified text CRLF-to-LF
normalization is allowed; ICC bytes remain exact. Persistent execution policies were
unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37286018683); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37286024093).
Evidence: evidence/M3-T02.json and evidence/M3-T02-ci.json.

Accepted 17/28 tasks and exercised T001-T053, with the first four characterization cases unchanged. Next: M3-T03 — Implement and test clean cooperative cancellation. This containing checkpoint changes only 6 reviewed handoff/evidence records; its own synchronization remains pending_verification. Raw media/private paths/argv remain ignored. No owner merge, release, deployment or owner approval is inferred.

Windows 10/Server 2016 JOB_LIST support and compatible inherited-job restrictions are
required. Unsupported atomic assignment fails closed. ImageMagick memory/map/disk limits
govern pixel cache, not every decoder/delegate heap allocation or arbitrary write; these
controls do not certify an exploit-proof sandbox.

The shared clock counts image-copy/retry elapsed time, but synchronous filesystem calls
are not preemptible. Native termination/drain grace can add wall time. ImageMagick's own
time limit is cooperative and SOURCE_DATE_EPOCH can disable it; the external wrapper
deadline remains the native-hang authority. A noncooperative kernel driver can still
delay completion of an I/O cancellation request; incomplete capture never permits
finalization.

Cleanup assumes the exclusively allocated cache namespace remains trustworthy. A
concurrent regular magick-* arrival is indistinguishable from a native cache file;
foreign names, directories and reparse entries remain with warnings. Unconfirmed
termination preservation was tested with a controlled returned flag after an actual
successfully terminated tree, not by reproducing an OS failure to kill.

Development targets, saved API/resource probes and tiny 128x160 measurement smokes
retain their own raw source/artifact bindings. They are separate from clean-I full
normal/control and hosted acceptance. No representative benchmark rerun,
quality/performance improvement, owner aesthetic approval or codec-wide exploit test is
inferred. Existing HEIC still-collection and independent colour-reference qualifications
remain in their original evidence.

The unrelated-process fixture uses a declared first synthetic image as its completion
oracle; RGB stdin only keeps the process blocked until synchronization ends, and all
stdin-derived frames are discarded. Actual default and forced-BOM controls passed on
both shells. The original hosted stdin bytes were not recorded, so its BOM explanation
remains a supported inference. Repaired-revision full normal/control and hosted gates
require their own immutable evidence.

Real Ctrl+C/cooperative cancellation, interrupted-copy exit behavior and
force-termination recovery remain M3-T03. This card stops after M3-T02 acceptance;
releases, deployment and owner acceptance stay separate.
