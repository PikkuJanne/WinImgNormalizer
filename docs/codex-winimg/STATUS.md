# Project status

Updated: 2026-10-05 — M3-T06 completed; owner excluded live UNC validation.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 17a405bfe0150607486248e845f93ba82e1d0d52
Next task: **M3-T07 — Obtain owner acceptance of the familiar workflow**
Task progress: **21 / 28 accepted**
Specified cases exercised: **T001-T066 (66 / 75); T001-T004 characterization, T007 control**
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
covered. M3-T07 owner approval, merge and release are separate. Stop at M3-T06.
