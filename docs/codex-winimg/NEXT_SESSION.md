# Next session handoff

**M4-T01 is complete; T068 passed.** Progress 23/28, T001-T068 exercised 68/75.
Next intended task **M4-T02 is pending and not started**; load its card when requested.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; preserve the non-Git
WinImgNormalizer-main snapshot. Read AGENTS/STATUS/TASKS, GIT_WORKFLOW,
RELEASE_AND_WEBSITE and M4-T01 evidence. Recheck clean worktree, canonical fetch/push
identity, exact live feature SHA and open draft PR22. Implementation
**3109a1b27d73f4084076cdd14ba06447c6243c0f** was independently observed clean/live-equal; its dated
observation is in task evidence. The containing final record SHA/sync is external
after commit and must be checked live again next session.

Application **1.0.0 is unreleased**, chosen after actual tags/releases were empty.
`Get-WinImgVersion` in the PS1 is the sole version source; the application retains
its self-contained two-file workflow. Existing usage/log headers consume it.
`tools/release/Update-ReleaseMetadata.ps1` derives CHANGELOG.md and root
release-metadata.json from the source and docs/release/NOTES.md. Use `-Check` for
exact drift detection. Compact JSON/UTF-8 without BOM/LF and Git attributes are
required for cross-host/Windows-checkout agreement. Actual publication and asset
fields remain null. No package ZIP, tag or release exists from this task.

Clean desktop full gates 482/482 in PS5.1.26100.9444 and PS7.6.5, zero skips/unrun,
24 suites and ten actual codec outcomes; controls 483/482/1/native 1. Analyzer now
includes the development generator: eight scripts, unchanged 15 rules, zero normal
findings and one expected control/native 1. Both sanitized full exports succeeded;
their exact seven new public source bindings have positive and privacy negatives.
Six T068 regressions prove actual help/log and changed-source derivation. The
superseded outer PS5.1 capture infrastructure failure is retained separately.
CI: Push and PR Windows CI passed for the tested implementation; exported artifacts
independently verified.

License/BAT/authorship/normalization defaults and dependency pins remain intact.
Owner workflow/appearance acceptance D37/T067 remains bound to repaired 5a32fc6;
the installed Q16-HDRI colour and physical owner observations keep that exact
binding. Version/help/log/documentation preparation supplies no new subjective
quality approval. Live UNC/network-share support remains owner-excluded D35.

For later packaging, use an explicit clean-commit allowlist, retain embedded ICC
provenance/license and choose package-facing documentation so maintenance/governance
links excluded from a ZIP do not become broken. This is handoff guidance, not
M4-T02 implementation or evidence. Do not bundle development tools/ImageMagick by
default. Merge/tag/release/settings and deployment still require their own approval.
