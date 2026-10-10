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
