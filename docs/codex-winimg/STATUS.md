# Project status

Updated: 2026-10-06 — owner smoke recorded; HDRI colour defect blocks acceptance.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 17a405bfe0150607486248e845f93ba82e1d0d52
Current task: **M3-T07 — Obtain owner acceptance of the familiar workflow (awaiting_owner)**
Task progress: **21 / 28 accepted**
Specified cases: **T001-T066 previously exercised (66 / 75); T067 partly exercised, blocked**
Pester suite: **468/468 passed in each desktop shell, zero skipped**
Failure controls: **469 total, 468 passed, one intended failure; native exit 1**
CI: **implementation push/PR Windows matrices passed, including both controls**
Containing evidence checkpoint synchronization: **pending_verification**

## Current behavior and verification

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

## M3-T07 owner gate — actual smoke, colour repair needed

At the 2026-10-05 preparation observations, clean review commit
4618a83cb2aa893aea9d6a1587987612624e1cf7 was live-equal to the feature branch. Owner-merged PR21/main
4d0bc6dd403a84b4cad770fdacfd394066c9fbd1 has the same tree. No application, BAT,
test or default changes were made. M3-T06's tested implementation and full 468/468
gates above remain separate evidence. Prior review-checkpoint push 37341068495 and
PR 37341075972 metadata report success. The new containing checkpoint still needs
post-commit live synchronization observation.

[Owner checklist](OWNER_ACCEPTANCE.md) and [prepared evidence](evidence/M3-T07.json)
are ready. The local packet outside the repository has 24 synthetic inputs, original/
current displays, five hash-bound historical baseline comparisons and real tiny-cap
warning examples. Fresh real Command preparation in PS5.1/PS7 returns 0 by default:
20 converted/1 video/1 duplicate/2 ignored, zero errors, five omitted frames. Each
1-byte run returns 2 with two valid above-target JPEGs. All 44 JPEGs fully decode;
source state and video bytes are preserved and reports complete/balance.

The 2026-10-05 preparation injected a disposable destination and explicit portable
Q16 ImageMagick without HDRI. It did not establish Explorer/Pictures/tool discovery
or owner appearance approval. Its driver failures and final v3 results are retained
as dated evidence; two historical quality differences follow the earlier byte-extent
correction. No new default is proposed.

On 2026-10-06 the owner followed the guided physical folder-to-BAT flow and reported
`Completed successfully. Press any key to close.`, with LogWarnings=0,
DiskLogIncomplete=False and FallbackDropped=0. The owner opened the actual Pictures
result, described the landscape as a smooth gradient and hard edges as clear/intact,
and confirmed the transparent upper-left area is white. These are partial observations,
not explicit named-commit acceptance; the BAT's native exit was not independently captured.

[Actual smoke evidence](evidence/M3-T07-owner-smoke.json) binds the copied launcher
to review commit 4618a83cb2aa893aea9d6a1587987612624e1cf7 and records independent
checks of only the approved synthetic source/output tree. Reports reconcile 24 inputs:
20 JPEGs, one equal-byte video, one duplicate and two ignored files; zero logged
errors/warnings, five omitted frames, complete balanced accounting. All 20 JPEGs fully
decode with correct geometry and stripped metadata. All original bytes/times/attributes,
mirrors, collision names and output timestamps are intact; 19 outputs match preparation.

The actual run used installed ImageMagick 7.1.2-32 **Q16-HDRI**, unlike preparation.
The tagged alpha output fails the independent colour reference: the semi-transparent
green patch is [0,213,169] instead of [127,213,169], maximum channel error 127 versus
JPEG tolerance 12. The white patch is correct. Zero process/report warnings do not
establish pixel correctness. [Native diagnosis](evidence/M3-T07-hdri-diagnosis.json)
reproduces the actual JPEG exactly. Clamping converted RGB before white composition
repairs this fixture in isolated native experiments; no production fix was applied.

M3-T07 stays **awaiting_owner**, T067 **blocked/incomplete**. Accepted tasks remain
**21/28**; T067's partial smoke does not add a passed case. Recommend a bounded tested
colour repair, followed by owner review at its new tested commit. Stop at this gate;
M4-T01 cannot start. Review/default approval, merge and release remain separate.

At the clean start of this follow-up, feature HEAD and live branch were
dd6753991edfb7dec7454d1af5c2fe43df75d103; draft PR22 remains open.
[Push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37345811050)
and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37345853641)
for that preparation checkpoint passed. Those pinned non-HDRI gates do not cover
this installed-HDRI failure. This containing record checkpoint needs its own
post-commit synchronization observation; its SHA is not embedded here.
