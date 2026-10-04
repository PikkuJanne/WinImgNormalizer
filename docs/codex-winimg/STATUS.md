# Project status

Updated: 2026-10-04 — M2-T01 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: d64a9034ff0b4324b2515900fdf8b4cb82f80bbb
Next task: **M2-T02 — Convert colour profiles before metadata removal**
Task progress: **10 / 28 accepted**
Specified cases exercised: **T001-T031 (31 / 75); T001-T004 are characterization; codec observations remain explicit**
Pester suite: **185 / 185 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **186 total, 185 passed, one intended assertion failed; exit 1 in each shell**
CI: **implementation push and PR Windows Server PS 5.1 / PS 7 matrices passed, including expected controls**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and evidence

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

Decision D24 records the scoped frame/page policy. Covered M1 output safety remains the established baseline.
Original sources, mirrored directories, deterministic plans, no-overwrite moves,
full JPEG validation, copied video bytes and successful duplicate links retain their
combined regression coverage. Deliberate static output does not preserve animation
or subsequent document pages; source counts and omissions are recorded honestly.

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

Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and
PS 7.6.5, Pester 5.9.1: 185/185 each, zero skipped.
Controls each have 186 total/185 passed/exactly one T007 deliberately false
assertion/exit 1. All four summaries bind the same clean committed checkout, with raw
file hashes and separate Git blob hashes for verified CRLF/LF normalization.
Persistent execution policies remain unchanged. Actual hosted Windows Server
matrices passed all four normal/control jobs: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37223580015) and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37223613778).
Evidence: evidence/M2-T01.json and M2-T01-ci.json. Codec fixture counts, capability
observations, exact artifacts and retained development history are separate fields.

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

## Handoff and remaining scope

Started clean at b7fe275dfea31e7954df0ebbeeaff02189cf6f8a, independently matching the
feature remote. Owner-merged PR #9 main f90de6ce37fdd3d760a32ce028fe3c40cc0290c9 had
the identical tree. Read-only fetch did not pull, merge, reset or change the branch.
Successor draft PR #10 contains M2-T01; the preserved non-Git snapshot remains separate.

Implementation synchronization is a past exact-SHA observation. This record-only
checkpoint binds that implementation; its own SHA/live remote state is reported
externally after normal commit/push. No raw logs, private media or local tools are tracked.

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

Stop after M2-T01; start M2-T02 only when next requested.
