# Project status

Updated: 2026-10-07 — M4-T05 candidate audit and manual T074 preparation complete.

Repository: PikkuJanne/WinImgNormalizer. Checkout: C:/projects/WinImgNormalizer.
Branch: codex/winimg-hardening. Reviewed/tested/tag-target candidate: **8edbcbaeb3425ec3a52eeafde212c32553755af1**.
Version: **1.0.0, unreleased and unsigned**. Selected WinImgNormalizer-1.0.0-portable.zip:
**184659 bytes; SHA-256 251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d**.

M4-T05 is done; progress **27/28 tasks, 74/75 specified cases**. T074 is a manual
proposal/audit gate; **T075/M4-T06 remain pending**, requiring exact explicit owner
publication approval. No application/BAT/builder/test/inventory/dependency or public
metadata behavior changed. The 27 mandatory suites/67 automated IDs remain 590/590
in each actual Windows desktop host; controls are 591/590/one exact T007/native 1.
Ten-script analyzer normal/control checks and both exact candidate CI matrices
passed. Four builds, measured same-host reproducibility, exact ZIP/license/provenance
checks and selected ZIP ordinary-account extracted PS1 media/setup checks passed.
See [audit](evidence/M4-T05.json), [CI](evidence/M4-T05-ci.json), and
[pending exact approval proposal](RELEASE_CANDIDATE.md).

All 18 items are reconciled without falsely completed or unresolved required
coverage. ACCEPTANCE.md now points its earlier M0-M3 gates to their actual evidence.
Owner D37/T067 acceptance stays at 5a32fc6533c6f9fae1d36c271b515d1d1dd4d8ad; automation grants no new consent.
Windows 10 is untested; live UNC remains owner-excluded D35. BAT setup is separate
from PS1 synthetic media; no physical gesture or real Pictures processing is claimed.
Checksums bind bytes, not signatures. Website download unavailable/render false,
publication fields null, deployment context null and screenshot/comparison arrays
empty; no public reuse approval or deployment context exists.

Owner-merged PR25/main is a separate dated starting observation. This record-only
checkpoint cannot contain its own final SHA; final clean/live synchronization,
successor draft PR URL/head and checkpoint CI are pending external post-commit
observation. No merge/tag/release/settings/deployment is performed by this task.
Only explicit owner approval of the named final PR/head, candidate and asset can
authorize M4-T06; actual post-merge/tag/source/license/download verification remains
required, including reviewed publication-state guard/projection changes if needed.
