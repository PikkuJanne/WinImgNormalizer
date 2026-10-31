# WinImgNormalizer 1.0.1 delivered - 2026-10-31

<!-- hotfix-1.0.1-checkpoint 2026-10-31T20:01:14.6229881Z -->

The authorized hotfix is complete: implementation, normal merge, stable/latest publication and fresh public-download verification passed.

Release: https://github.com/PikkuJanne/WinImgNormalizer/releases/tag/v1.0.1
Portable ZIP: https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.1/WinImgNormalizer-1.0.1-portable.zip
SHA-256: 0f5ef869fb2a868a49754fef3116c7b9c9618f448edf1e30f8413d5f5c93dee5
Source and tag commit: 25cd8da78595abd256d96b2ff7d117e2b3a02c18. This identifies the release bytes, separately from any later publication-metadata commit.

Actual local e04180e0 coverage: Windows 11 Pro build26300, Windows PowerShell5.1 and PowerShell7,716/716 in each host with no skips; both analyzers reported zero findings.
Premerge feature CI: https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/41231071132; head 0168069f62d9bc5db3aeb0248b5d6172ae585d0d.
Exact candidate-byte workflow: https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/41233089750; status completed/success.
Exact stable public-byte workflow: https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/41234831588; status completed/success.

The automated workflow uses genuine7.1.2-5 on PATH, the private bundled7.1.2-35 Beta engine and authenticated app-local Microsoft OpenMP. It covers real JPEG/PNG/WebP/GIF/TIFF/HEIC/HEIF inputs, byte-identical valid video copies, private runtime integrity, policy denial, empty/video-only work and the real BAT/private-console paths.
Windows 10 itself was not executed. Windows 11 and Windows Server2025 are separate observations. No owner test or separate appearance sign-off is pending. No live website deployment was performed.

Evidence: evidence/HOTFIX-1.0.1.json. Earlier records below are historical and remain unchanged.
This checkpoint cannot contain its own future metadata-commit SHA; observe and report that later SHA separately after committing.

---

# Authorized 1.0.1 hotfix — 31 October 2026

The maintainer's attached HOTFIX.md and ACCEPTANCE.md are one authorized delivery:
implement the scoped runtime/console repair, automate Windows acceptance, use the
normal merge route, publish a stable patch, and verify its public download. This
patch supersedes the historical owner-test, appearance sign-off, one-task-per-chat
and separate merge/release approval stops. It does not authorize website deployment,
settings changes, protection bypasses, or changes to the installed system engine.

Work starts at current public main `e8d7b6c2b4566f408c22ed998f155e1baf1a7dc8`
in a clean dedicated clone/branch. Earlier dated records below remain historical.
Current acceptance and release observations will be recorded in
`evidence/HOTFIX-1.0.1.json`. No result is certified before it is actually observed.

---

# Website asset cleanup checkpoint — 10 October 2026

This separately authorized cleanup starts from public main `e5376d41879a2e4ffb5da5e607bfb634f1598621` in a fresh
isolated clone. The obsolete website-only paths and direct references are prepared
on a focused branch; application/runtime/package code and existing tags/releases
remain unchanged. Historical engineering acceptance is retained.

See [the cleanup record](../WEBSITE_ASSET_CLEANUP.md) for scope,
classification and checks. GitHub PR/check/default-branch delivery must still be
observed before treating current-main cleanup as complete.

---

# Project status

Updated: 2026-10-07T06:34:05.769146Z — M4-T06 approved publication complete.

Progress: **28/28 tasks, 75/75 specified cases**. T075 passed through actual
approved publication and public-byte/canonical/website consistency verification.
Version **1.0.0** is public at [v1.0.0](https://github.com/PikkuJanne/WinImgNormalizer/releases/tag/v1.0.0), tag source
**8edbcbaeb3425ec3a52eeafde212c32553755af1**. The unchanged ZIP is **184659 bytes**, SHA-256
**251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d**, with both exact reviewed sidecars.

Owner statement **“v1.0.0 Approved”** was recorded after the exact tag/public-release
proposal. [Publication evidence](evidence/M4-T06-publication.json) records scope,
actual release/asset IDs, anonymous downloads, source/license/package parity and
safe local-main fast-forward to the owner's PR27 merge ccbf6626. No remote default
branch push or agent merge occurred.

Published metadata implementation: **9dcc504071540302d6cbbde2de985b25cc4d1929**. Both desktop hosts
passed **669/669** normal assertions, all 27 mandatory suites/67 required IDs,
with only the intended T007 failure in controls and the intended analyzer control.
[Desktop evidence](evidence/M4-T06-validation.json) and
[exact implementation CI](evidence/M4-T06-publication-ci.json) record actual outputs,
source 71 bindings, analyzer 12 bindings/15 rules and private-safe exports.

The approved ZIP's embedded unreleased wording remains its preparation snapshot;
the public release explains it. The generator/website guard now preserve and
validate the dated published record, including exact URLs and byte identities.
The builder remains unreleased-only; new tests distinguish clean published-head
refusal from explicitly draft fixture-package smoke. Runtime/dependencies/branding
and the approved files remain unchanged.

Windows 10 testing is **out of scope (D43)**, with no Windows 10 pass claimed.
Live UNC exclusion D35 and exact historical T067/D37 acceptance remain unchanged.
The release is unsigned. Website deployment and public owner-media reuse remain
separate unapproved work, with no supplied provider context.

PR28 carries the final publication metadata and records for owner review. Its merge
requires separate owner action; release approval does not name that new merge.
This containing record checkpoint's final SHA, push, CI and live synchronization
must be observed externally after commit; no future result is certified here.
