# Project status

Updated: 2026-10-04 — M2-T02 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: a30c456594a31368294588cc198444b38bc3e0e3
Next task: **M2-T03 — Make filesystem and native filename handling literal-safe**
Task progress: **11 / 28 accepted**
Specified cases exercised: **T001-T036 (36 / 75); T001-T004 are characterization; codec limits remain explicit**
Pester suite: **210 / 210 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **211 total, 210 passed, one intended assertion failed; exit 1 in each shell**
CI: **implementation push and PR Windows Server PS 5.1 / PS 7 matrices passed, including expected controls**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and evidence

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

Decision D25 records the tested colour/profile and metadata policy. Source preservation, deterministic mirrored names, transactional finalization,
verified copied videos and success-only duplicate links retain the M1 regression baseline.
The deliberate first displayed GIF/WebP frame, first TIFF page and decoder-primary
HEIC/HEIF selection retain M2-T01 coverage and informational omission reporting.
Actual HEIC evidence remains a two-image still collection, with timed animation,
thumbnails and auxiliary-image behavior unverified.

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

## Handoff and remaining scope

Started clean at ad56bdf9ae991812a819884e50647e454d443a4f, independently matching the
feature remote. Owner-merged PR #10 main 58d628ce189a3dd508d05fdce24937e71f74a222 had
the identical tree. Read-only fetch preserved the checkout and branch.
Successor draft PR #11 contains M2-T02; the preserved non-Git snapshot remains separate.

Implementation synchronization is a past exact-SHA observation. This record-only
checkpoint binds that implementation; its own SHA/live remote state is reported
externally after normal commit/push. Raw logs, private media and local tools stay ignored.

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

Stop after M2-T02; start M2-T03 only when next requested.
