# Project status

Updated: 2026-10-05 — M3-T02 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 73bb2764ecf3ce3bf0fcb9e82feec434005f1e4a
Next task: **M3-T03 — Implement and test clean cooperative cancellation**
Task progress: **17 / 28 accepted**
Specified cases exercised: **T001-T053 (53 / 75); T001-T004 are characterization**
Pester suite: **304/304 passed in each desktop shell, zero skipped**
Failure controls: **305 total, 304 passed, one intended failure; native exit 1**
CI: **implementation push/PR Windows matrices passed, including controls**
Containing evidence checkpoint synchronization: **pending_verification**

## Current behavior

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

## Verification and handoff

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

Actual development results and their source/artifact qualifications are retained in
evidence/M3-T02.json and SESSION_LOG.md. Clean implementation desktop and hosted gates
are separate acceptance observations; earlier task histories retain their original
revisions.

Started clean at fa4ab9fd609918500a5dcec1c5972c24d23c2c37, equal to the live feature
branch; owner-merged PR #16 main a2ef99ac52631e84988669905aa8e80d44682b79 had the same
tree. Continued without changing checkout/history. Successor draft PR #17 contains this
task. Implementation synchronization is a past exact-SHA observation. This record-only
checkpoint's SHA/live synchronization remains pending_verification until separately
observed after normal commit/push.

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

This session stops at M3-T02. Begin M3-T03 only when requested.
