# Project status

Updated: 2026-10-06 — M4-T02 portable ZIP and clean Windows extraction verified.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested M4-T02 implementation: 69822554c69223d4e048b2a58d1f3203cd01428c
Application version: **1.0.0 (unreleased; unsigned local package prepared)**
Owner-accepted normalization implementation: 5a32fc6533c6f9fae1d36c271b515d1d1dd4d8ad
Completed task: **M4-T02 — Build a portable release ZIP from reviewed tracked content (done)**
Next intended task: **M4-T03 (pending; not started)**
Task progress: **24 / 28 accepted**
Specified cases: **T001-T071 exercised (71 / 75); T069-T071 passed**
Pester suite: **512/512 passed in each desktop shell, zero skipped/unrun; 25 mandatory suites**
Failure controls: **513 total, 512 passed, one intended failure; native exit 1**
Static analysis: **nine scripts, zero normal findings; one intended control finding/native exit 1**
CI: **Implementation push and PR Windows matrices passed; four jobs and sanitized artifacts independently verified**
Containing evidence checkpoint synchronization/CI: **pending external post-commit observation**

## Accepted M3-T06 baseline and its dated verification

The mandatory gate inventories all 23 maintained suites, requires actual passed
T005-T066 coverage except the separate T007 failure control, and rejects zero
discovery, skips, unrun/inconclusive tests, failed blocks/containers, missing suites,
source changes and absent real codec outcomes. FileInfo file containers and
ScriptBlock controls are handled explicitly and covered by a fresh pinned-Pester
child regression. Focused runs cannot replace full acceptance.

All ten advertised extensions have real codec results: jpg/jpeg/png/bmp/tif/tiff/
gif/webp plus the existing HEIC/HEIF collection cases. Reader-absence controls mask
the capability table while actual supported sibling conversion and video copying
complete; the missing JPEG writer fails setup. Real positives establish support for
the pinned ImageMagick 7.1.2-32 build, not every installed codec build.

Pester 5.9.1, ImageMagick 7.1.2-32 and PSScriptAnalyzer 1.25.0 remain pinned development
dependencies. Analyzer archive/manifest verification and safe fresh extraction precede
use. Fifteen scoped security/defect/syntax rules scan seven maintained application/
infrastructure scripts; normal results have zero findings. Each unsafe-expression
text control produces exactly one intended finding and native exit 1. The application
change only renames a caught exception variable that shadowed automatic $Error.

The workflow uses Windows Server 2025, explicit powershell/pwsh jobs, contents:read,
reviewed full action SHAs and persist-credentials:false. Native failures remain
failures. Exactly two allowlisted synthetic JSON files are uploaded per shell after
reparse-path rejection; raw logs, XML, diagnostics, test names and media stay local.
Retained environment and text fields reject structured values before they can retain
private text. UTC timestamp kind and fractional seconds survive both JSON readers.

Clean implementation 17a405bfe0150607486248e845f93ba82e1d0d52 passed **468/468** with zero
skips/unrun tests on actual Windows 11 Pro build 26300, Windows PowerShell 5.1.26100.9444
and PowerShell 7.6.5. Both full controls had 469 total/468 passed/one exact intended
T007 failure/native 1. Standalone normal/control analyzer gates passed in both shells.
Native parent records and summary/XML/console hashes bind these observations. All 45
source bindings verify exact bytes or CRLF/LF-only equivalence against the tested Git
blobs; source/commit/persistent policies remain unchanged.
Evidence: [desktop](evidence/M3-T06.json).

Hosted [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37333526329)
and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37333528830)
passed all four jobs: 468/468 normally, 469/468/1 assertion controls, zero normal
analyzer findings and one exact analyzer-control finding. Actual Server 2025
Datacenter 10.0.26100, image win25-vs2026 version 20260925.250.1, runner 2.337.0,
PS 5.1.26100.33438/PS 7.6.6, x64 en-US are recorded separately from desktop automation.
Downloaded artifact digests, two-file content/privacy checks, source/Git bindings
and synthetic PR merge provenance are independently verified in
[CI evidence](evidence/M3-T06-ci.json).

## Handoff qualifications

Started clean at fd4b60cd0e9a3107ee48e85b0589c837e4084f7e; owner-merged PR20 main
e6b50d60937e81409dbb481e09bd42fa69c25111 has the same tree. Continued on the feature
branch without changing the preserved non-Git snapshot or Git history. Successor
[draft PR21](https://github.com/PikkuJanne/WinImgNormalizer/pull/21) contains M3-T06.
This record-only checkpoint's final SHA/live synchronization requires an independent
post-commit observation; its own hash is not embedded here.

Earlier fixture/export failures and superseded desktop runs remain recorded. I2
hosted assertions passed 459/459 but the string-only container guard rejected actual
FileInfo suites; both matrices failed. I3 passed 461/461 and both controls in all
environments, then review found nested-array retention and lossy UTC serialization
in the exporter. Actual I3 uploaded payloads were clean; its exported timestamps are
qualified. I4 repairs those boundaries and adds seven regressions. These earlier
observations are separate from final acceptance, never relabeled successful I4 runs.

Live UNC/network-share validation and support are **out_of_scope** by the owner's
2026-10-05 decision (D35), with no suitable server available. This is no longer an
outstanding acceptance/release gate. Local long-path, lexical drive/UNC-root and
local identity checks remain mandatory; historical localhost results and `not_run`
observations are preserved without claiming live network-share support. See
[scope record](evidence/M3-T06-scope.json). Analyzer scope excludes Pester DSL bodies, legacy and vendor
code. Package checksums bind bytes without claiming signatures/publisher identity.
Desktop automation and Server CI do not supply physical Explorer gestures or owner
workflow/quality approval. JPEG/colour/frame/size defaults, source/no-overwrite safety,
video fidelity, launcher pause/exit behavior and owned native processing remain
covered. M3-T07 owner approval, merge and release are separate.

## M3-T07 D36 repair and owner gate

Original review 4618a83cb2aa893aea9d6a1587987612624e1cf7 and actual owner smoke at
that code remain dated evidence. The owner reported clean BAT completion, smooth
landscape, intact detail and the white transparency area. Independent checking found
the installed Q16-HDRI semi-transparent green red channel 0 instead of 127. These
partial observations supplied no named-commit acceptance; native BAT exit was not
captured. Original packet/output/source/report bindings remain preserved.

The owner explicitly authorized the bounded fix and short recheck (D36). At clean
base 45c6795554b70a1fc23dea04b1b4339a7d69c430 actual installed-HDRI Pester runs fail
two T036 checks in each shell: 25 total/23 pass, PNG/JPEG channel error 127. Earlier
maintained-runner attempts rejected the non-pinned executable before discovery;
those zero-test infrastructure failures are retained, not relabeled colour runs.

Clean repair 5a32fc6533c6f9fae1d36c271b515d1d1dd4d8ad adds only converted-sRGB
clamping before white composition. T036 covers tagged/untagged real references,
two controlled sharing-start failures and six actual scale conversions, checking
conversion/clamp/white/strip order on all eight calls. Independent reference/profile/generator,
PNG tolerance 3/JPEG tolerance 12, BAT and defaults are unchanged.

[Repair evidence](evidence/M3-T07-repair.json): actual PS5.1/PS7 portable full gates
pass 469/469, zero skips/unrun/blocks, all 23 suites and ten real codecs. Both controls
have 470 total/469 pass/one exact T007 failure/native 1. Normal analyzer scans seven
scripts with zero findings; each unsafe-expression control yields one intended
finding/native 1. Colour-only runs pass 26/26 for each build/host (four runs). The
separate installed-HDRI driver verifies Pester and executable/version; the maintained
portable pin is intact. Counts, native results, source/Git and raw artifact hashes bind
the clean implementation; all policies and sources remain unchanged.

[Repaired artifact checks](evidence/M3-T07-repair-artifacts.json) cover two installed-
HDRI Command runs using disposable destinations. Each returns 0 with 24 inputs:
20 JPEGs, one equal-byte video, one duplicate, two ignored, no errors/warnings and
five frame omissions. Forty JPEGs independently decode; source files/dirs, metadata,
geometry, collision/duplicate mapping, mirrors and output timestamps pass. Nineteen
JPEGs equal the old owner output, one alpha is corrected; all 20 match preparation
and each other across hosts. Corrected alpha maximum error is 1, within tolerance 12.

[Implementation CI](evidence/M3-T07-repair-ci.json) independently verifies the
[push](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37476651394)
and [PR](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37476659134)
jobs, exported counts/controls, actual codec results, provenance and privacy. CI
does not provide owner consent. This containing record checkpoint has unchanged
runtime/tests and requires post-commit live synchronization observation.

## Completed repaired owner recheck and acceptance

The first one-image attempt used the older launcher: the live paused CMD command
proves the original packet BAT ran on the new Colour recheck source. Its green
error 127 and plausible visual Yes remain dated [old-launcher evidence](evidence/M3-T07-recheck-old-launcher.json).
The two Launcher window titles were insufficiently clear; this is not a repair regression.

A distinct **Fixed launcher 5a32fc6** folder holds exact tested copies. The repeated
owner launch reports completion/pause and its live command binds that fixed BAT,
adjacent PS1 and approved source. [Fixed-run readback](evidence/M3-T07-recheck-fixed-launcher.json)
fully decodes the one JPEG, matches repaired automated bytes, verifies reference
error 1 within tolerance 12, metadata removal, balanced 1/1/no warnings and preserved
source bytes/times/attributes against the preceding snapshot. Output timestamps match.
Native conversion 0 is logged; native BAT exit remains independently uncaptured.

The owner directly replied **Accepted** to “Do you accept fixed version 5a32fc6 for
the drag-and-drop workflow and image appearance?”.
[Acceptance record](evidence/M3-T07-acceptance.json) binds the exact prompt/reply to
**5a32fc6533c6f9fae1d36c271b515d1d1dd4d8ad**. No earlier area Yes supplies this consent.
Not all original comparison images were individually visually reviewed; their
existing automated checks and earlier visual statements remain separate.

**M3-T07 done; T067 passed; 22/28 accepted; 67/75 specified cases exercised.**
M4-T01 was pending at that M3-T07 acceptance checkpoint. Preceding record checkpoint
c268e392956b4b1c229e68f184cd2dd436e56d23 has unchanged runtime/tests/workflow and
independently observed successful push/PR CI metadata (37479656551/37479662600).
Detailed implementation artifact verification remains separately bound to 5a32.
That acceptance checkpoint required an external post-commit synchronization
observation. No further task or publication work had started at that observation.

## M4-T01 — one application version and unreleased release notes

At clean implementation **3109a1b27d73f4084076cdd14ba06447c6243c0f**, the sole semantic-version
literal is `Get-WinImgVersion` in the self-contained application. Existing usage
and run-log headers use it. The development generator derives CHANGELOG.md and
release-metadata.json from it and the version-free release prose. Exact checks
reject stale output without writes; compact ordered JSON and scoped LF attributes
preserve bytes across Windows PowerShell 5.1/PowerShell 7 and Windows checkouts.
No runtime support file or new invocation form is introduced.

Successful local/remote tag and GitHub release inspection found zero existing
records. **1.0.0 is the first managed proposed version and remains unreleased**;
no tag, release or ZIP was created. Published/date/download/asset fields are null.
The changelog distinguishes fixes, retained defaults/safeguards, third-party
requirements, supported/tested environments and limitations. LICENSE and original
Janne Vuorela authorship remain intact; the BAT and normalization defaults are
unchanged. Live UNC remains owner-excluded D35.

[Desktop evidence](evidence/M4-T01.json) records **482/482** in each actual desktop
host, no skips/unrun/inconclusive/failed containers, all 24 suites and ten real
codec outcomes. Six T068 regressions cover real help without ImageMagick/output,
real video/log preservation and substituted version, exact non-writing checks,
stale rejection and changed-source/repeat generation. Seven exporter regressions
admit only the exact public release inputs/outputs while rejecting nearby private
and traversal names. Both assertion controls are **483 total/482 pass/one exact
T007 failure/native 1**. Analyzer checks eight scripts under the unchanged 15 rules:
zero normal findings, one expected unsafe-expression control/native 1. Both actual
full exports complete. Exact source/Git/dependency/parent/artifact hashes and
persistent policy preservation bind these results to the clean implementation.

The first outer PS5.1 capture wrapper escalated stderr from an intentional logging
failure fixture and ended without a completed suite summary. That infrastructure
attempt is preserved separately; only the ignored wrapper was corrected before
fresh clean runs. It is not a failed application result or passing test run.

CI: Push and PR Windows CI passed for the tested implementation; exported artifacts
independently verified. The final containing governance/evidence checkpoint's SHA and live
feature synchronization are observed externally after commit. M4-T01 is done,
23/28 accepted and T001-T068 exercised; **M4-T02 remains pending**. Packaging,
publication and deployment retain their separate task/approval boundaries.

## M4-T02 accepted portable packaging

Build-Release.ps1 reads seven explicitly mapped regular Git blobs from the exact
clean current full commit, preserves PS1/BAT bytes and refuses dirty/stale/missing/
linked inputs, linked output ancestors and destination replacement. It produces
an eight-file ZIP (seven inputs plus generated manifest), outer source/asset
provenance and SHA256SUMS. Compact setup guidance and full MIT/CC0 notices accompany
the self-contained runtime; tests/governance/tools/media and ImageMagick are excluded.
See [packaging instructions](../release/PACKAGING.md).

At clean implementation 69822554c69223d4e048b2a58d1f3203cd01428c, actual Windows 11 Pro
build 26300, PS5.1.26100.9444 and PS7.6.5 full gates each passed 512/512 with zero
skips/unrun tests, 25 suites and ten real codec outcomes. Both controls produced
513/512/one exact T007 failure/native 1. Both normal/control analyzer and sanitized
exports completed. Four actual clean-commit ZIP extractions (normal/control in each
shell) passed exact BAT setup/pause and real extracted PS1 image/video smoke via
the existing owned OutputParent seam. Actual children were non-admin; source state,
video bytes and persistent security state remained unchanged. These are automated
checks, not a physical Explorer gesture or new subjective owner acceptance.

Independent audits recompute all source/file/archive/checksum bindings and verify
legal/ICC provenance and package-local links. The PS7 prepared ZIP is 181042 bytes,
SHA256 18754f1e69a76460dee60166a4f910d04e93b669b1782c01aad12c033f0b28fd. The PS5 ZIP is 181102 bytes:
Framework NoCompression uses DEFLATE stored-block framing, whereas the tested PS7
writer uses ZIP Stored entries. All eight decoded files agree across hosts; each
host repeats all three external assets byte-identically. No universal ZIP-writer
byte equality, signature, public download or published release is claimed.

The first clean I1 PS7 normal run passed 512/512. I1 PS5 had only the incorrect
stored-size equality assertion (511/512); extraction and real media passed. The
superseded owned PS7 control was stopped, supplies no control counts and is retained
separately. I2 f4da7b8 qualified the framing contract and both desktop normal gates
passed 512/512. All four I2 hosted jobs then failed the same independent T070 identify
inspection (normal 512/511/1; control 513/511/2), so I2 local controls were stopped and
supply no passing counts. A partial PS5.1 NativeLifetime observer FormatException is
retained separately. I3 e0731fc corrected the independent JPEG inspection path;
both desktop validation parents were stopped before normal completion when the
separate fixture identity-publication race was diagnosed. I4 publishes complete
owned fixture identity bytes with an atomic move; these two corrections change
tests only, preserving runtime, builder and package inputs. The earlier hosted
stderr was not uploaded, so its underlying native failure is not retrospectively
proved. Focused developer replicas and the initial source-change gate rejection
are qualified in [desktop evidence](evidence/M4-T02.json), alongside final I4
acceptance. [CI evidence](evidence/M4-T02-ci.json) binds actual implementation runs,
downloads and PR synthetic merge provenance, separately from desktop observations.

Prior checkpoint 958b387 now has both push/PR CI successes. The owner merged PR22
into main 1eaaa356435b5d15cc25c909a9dd0c82635d479e with an identical prior feature
tree. This task continued the same feature branch in [draft PR23](https://github.com/PikkuJanne/WinImgNormalizer/pull/23).
Application/BAT/license bytes and dependency pins remain unchanged. This containing
governance checkpoint needs its external SHA/live sync and CI observation after
commit; it cannot contain its own hash. M4-T03 has not started. Future tag/release
publication and website deployment remain separate owner gates; live UNC remains D35.
