# Project status

Updated: 2026-10-04 — M1-T03 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: da91c0d648b387b270bcdd7269a48ee1c63d11eb
Next task: **M1-T04 — Validate temporary outputs and finalize without overwriting**
Task progress: **6 / 28 accepted**
Specified cases exercised: **T001-T019 (19 / 75); T001-T004 are characterization**
Pester suite: **119 / 119 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **120 total, 119 passed, one intended assertion failed; exit 1 in each shell**
CI: **implementation push and PR Windows Server PS 5.1 / PS 7 matrices passed, including expected controls**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and evidence

M1-T03 builds the complete eligible-source map before processing. Total ordinal
relative-source ordering is independent of culture and enumeration; case-insensitive
reservations cover mirrored directories, generated namespaces, videos and all unique
legacy image names before extension/numeric suffixes are assigned. Unique names stay
familiar; same-stem formats use __sourceext and checked __N suffixes starting at 2.
The complete ordered map remains in memory and local PLAN log lines, including later
skipped, unavailable-decoder and failed items.

Images use absent neutral candidates under exclusive GUID work directories. Final
two-argument File.Move and video File.Copy(overwrite=false) refuse file/directory
arrivals even after the availability check. Paths are rechecked before finalization;
new naming/allocation/finalization/cleanup warnings return 2. Long generated scratch
receives an internal extended Windows prefix for ImageMagick; names are not shortened.
Full native outcome/content validation and staged video copies remain M1-T04.

Actual Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester 5.9.1 and verified
ImageMagick 7.1.2-32: 119/119 each, zero skipped; controls 120 total / 119 passed / one identified
false assertion / exit 1. All four summaries bind exact committed runtime/test hashes;
persistent execution policy is unchanged. T017-T019 cover real JPG/PNG/BMP collisions
with full JPEG decode, source hashes/timestamps and unchanged video bytes, controlled
HEIC/case variants, 80 seeded shuffle/culture comparisons, complete PLAN logs before
conversion, external arrivals/post-check races, allocation continuation and native
scratch longer than MAX_PATH. HEIC coverage proves routing, not an actual codec run.

Evidence: evidence/M1-T03.json and M1-T03-ci.json; prior evidence remains retained.
All four implementation CI jobs passed: push https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37213722521; pull_request https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37213726939.

## Handoff and remaining scope

Started clean at e178acb6228e35ed1ba71393d5322f884e5e70e1, independently equal to the
feature remote. Owner-merged PR #5 main cd64fee had the same tree; read-only fetch did
not pull/merge/reset. Successor draft PR #6 contains M1-T03. Initial 95-test development
checks passed 92 / failed 3 in each shell due to long native scratch; extended prefixes
corrected that regression. Then combined 119-test development checks passed 116 / failed 3
because controlled large-cap/HEIF guards rejected the extended spelling. Those guards
now remove only that spelling before retaining their owned-path checks. Initial
failures remain in sanitized evidence and ignored raw artifacts.

Implementation was independently synchronized after normal push. This record-only
checkpoint binds that implementation; its own SHA and live remote observation are
reported externally after commit/push. No live UNC, real Pictures, private media,
NTFS case-sensitivity change or manual launcher/owner quality result is claimed.
Inherited file-symlink denial and controlled reparse coverage remain distinct from
real junction checks. No hostile-filesystem sandbox or universal long-path guarantee.

Candidate validation, staged video-copy/source-stability recovery, native execution,
frames/colour, duplicate safeguards, lifecycle/general exits, complete corpus/analyzer
and owner quality/publication gates remain later tasks. No release readiness claim.
Stop after M1-T03; begin M1-T04 in a fresh thread.
