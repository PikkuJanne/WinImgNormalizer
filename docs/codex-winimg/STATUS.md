# Project status

Updated: 2026-10-05 — M2-T04 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 0e490cc5f972b5f22374d2599eed929e477b40ca
Next task: **M2-T05 — Capture useful errors and remove indiscriminate fallback**
Task progress: **13 / 28 accepted**
Specified cases exercised: **T001-T042 (42 / 75); T001-T004 are characterization; prior path/codec limits remain explicit**
Pester suite: **242 / 242 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **243 total, 242 passed, one intended assertion failed; exit 1 in each shell**
CI: **implementation push and PR Windows Server PS 5.1 / PS 7 matrices passed, including expected controls**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and evidence

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

Decision D27 and primary references I9 record the intended byte target, exact final-length status and retained algorithm. Source preservation, deterministic mirrored names, literal native/launcher inputs, no-overwrite transactions, copied videos and success-only duplicate links remain combined regression gates. Frame/page and ICC/orientation/white-alpha/metadata policy remain intact; prior qualified HEIC, CMYK and path-provider evidence is not broadened by this task.

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

## Handoff and remaining scope

Started clean at f9b00d8befac93c4fa1116efcbc2e9d7144f4509, equal to the feature remote. Owner-merged PR #12 main 36d489f61cdaef93dca3f9afde95ce7ff6484bb7 had the identical tree. Read-only fetch preserved checkout/history. Successor draft PR #13 contains M2-T04; the preserved non-Git snapshot remains separate.

Implementation synchronization is a past exact-SHA observation. This record-only checkpoint binds that implementation; its own SHA/live remote state is reported externally after normal commit/push. Raw paths, media, native argv and local tools remain ignored. No merge, tag, release, deployment or owner quality/default approval is implied.

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

Stop after M2-T04; start M2-T05 only when next requested.
