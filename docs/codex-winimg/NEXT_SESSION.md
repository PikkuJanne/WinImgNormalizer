# Next session handoff

Current task: **M3-T07 — D36 bounded repair in_progress**. The owner ran the everyday Windows smoke
on 2026-10-06 and supplied partial appearance observations. **T067 is blocked/incomplete**
by an objective installed-HDRI alpha colour defect and absent explicit named-commit
acceptance. Accepted tasks remain 21/28; T001-T066 are previously exercised, T067
partly exercised. M4-T01 stays pending.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; preserve the non-Git
WinImgNormalizer-main snapshot. Read AGENTS.md, STATUS.md, TASKS.json,
GIT_WORKFLOW.md, tasks/M3-T07.md, OWNER_ACCEPTANCE.md and the three M3-T07 evidence
records. Recheck clean state, canonical fetch/push identity, exact live feature SHA
and draft PR22. This record-only checkpoint's final SHA/sync is reported externally
after commit. Its preceding checkpoint dd6753991edfb7dec7454d1af5c2fe43df75d103 was
clean/live-equal at the follow-up start; push 37345811050 and PR 37345853641 CI passed.

The owner dragged the literal `Drop this folder` onto the copied unchanged BAT,
reported clean final reporting and completion/pause, opened the actual Pictures
result, described the landscape as a smooth gradient and hard edges as clear/intact,
and confirmed the upper-left white transparency area. The last `Yes` answered
that area check only. No overall appearance/default/workflow approval was supplied;
the native BAT exit was not independently captured. Additional comparison images
were not all visually reviewed.

Only the approved synthetic source/output tree was independently inspected. All
24 originals and eight source directories retain bytes/times/attributes. All
20 JPEGs decode, geometry and metadata policy match, one video has equal bytes,
mirrors/collisions/duplicate mapping/output timestamps and complete balanced reports
are verified. Zero report warnings coexist with a real pixel defect. Nineteen
outputs match the pinned preparation; the alpha JPEG differs.

Actual owner ImageMagick 7.1.2-32 Q16-HDRI retains a negative red channel after ICC
conversion. White composition then gives the semi-transparent green patch [0,213,169]
instead of independent [127,213,169], exceeding established JPEG tolerance 12.
The native diagnosis reproduces actual owner output bytes. `-clamp` immediately
after either sRGB conversion branch and before white alpha removal fixes the fixture;
clamping before conversion or after composition fails. Portable non-HDRI bytes
remain unchanged in these diagnostic controls. No runtime/tests/defaults were edited.

Next safe action is a bounded tested colour repair if the owner requests it, as the
M3-T07 task card requires for material changes. Reuse independent T036 reference,
check clamp order on every retry and retain profile intent, alpha, CMYK, orientation,
stripping and JPEG defaults. Require meaningful Q16/HDRI regression and the normal
PS5.1/PS7 full gates/controls before renewed owner review at the exact repaired commit.
Native diagnostic success alone is not production validation or owner approval.
Do not continue M4 or sign acceptance on the owner's behalf.

The local packet location remains in ignored .scratch/M3-T07-location.json. Copied
BAT/PS1 match review commit 4618a83cb2aa893aea9d6a1587987612624e1cf7. Keep original
smoke/preparation outputs separate from any repaired run. Raw media, logs and private
Pictures paths remain local. M3-T06's 468/468 desktop and hosted non-HDRI results
remain dated evidence, qualified for that build. Preserve BAT CRLF/-text. Live UNC
remains owner-excluded D35. No merge/tag/release/settings/deployment is authorized.


D36 update: the owner explicitly authorized the colour fix and short guided recheck
on 2026-10-06. Complete only the amended task-card repair. Real pre-fix installed
HDRI colour runs in both hosts fail 2/25 T036 assertions (23 pass, zero skips).
Runner-pin rejection attempts are distinct zero-test infrastructure failures. Keep
the maintained portable pin; use separately verified installed-HDRI colour driver.
Validate the exact clean implementation, update durable repair evidence, and prepare
a fresh one-image recheck packet before returning M3-T07 to awaiting_owner.
