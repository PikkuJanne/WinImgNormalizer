# Next session handoff

Next task: **M3-T03 — Implement and test clean cooperative cancellation**. M3-T02 is complete.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; WinImgNormalizer-main remains the preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json, GIT_WORKFLOW.md, tasks/M3-T03.md and its referenced specs/cases T054-T056. Recheck clean state, canonical fetch/push identity and the exact advertised feature SHA before editing. Tested implementation is 73bb2764ecf3ce3bf0fcb9e82feec434005f1e4a; verify the final evidence-checkpoint SHA and draft PR #17 independently.

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

M3-T03 implements and tests clean cooperative cancellation: stop taking new files, finish or stop only owned native work, and remove only owned incomplete temporary data. Preserve finalized outputs, emit one clear cancelled summary, and check real Ctrl+C behavior in both Windows shells as required by that card. Retain source/no-overwrite safety, reporting, public positional forms, the default 1,048,576-byte cap, scale/JPEG settings, byte-identical videos and the accepted timeout/resource policy.

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

Fresh children may use process-only Bypass and clear inherited PSModulePath; never change persistent PATH/policy or redefine USERPROFILE. Keep raw paths, media, argv and tools ignored. Preserve actual failures. Use one successor draft if the owner merged this PR; normal scoped feature commits/pushes remain authorized. Merges, tags/releases, settings/deployment and owner acceptance remain separate gates. This session stops at M3-T02. Once requested, complete only M3-T03 and stop.
