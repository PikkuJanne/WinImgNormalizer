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

# Next session handoff

**M4-T06 is done; project task progress is 28/28 and specified cases 75/75.**
The approved public [v1.0.0 release](https://github.com/PikkuJanne/WinImgNormalizer/releases/tag/v1.0.0) and unchanged three
assets were anonymously verified twice. Read [publication evidence](evidence/M4-T06-publication.json),
[exact clean desktop validation](evidence/M4-T06-validation.json) and
[implementation CI](evidence/M4-T06-publication-ci.json).

Publication source remains **8edbcbaeb3425ec3a52eeafde212c32553755af1**, ZIP SHA-256
**251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d**. Do not move its tag, rebuild/relabel the release
or replace released bytes. Embedded unreleased wording is the disclosed preparation
snapshot. Public metadata implementation is **9dcc504071540302d6cbbde2de985b25cc4d1929**.

Reverify the containing checkpoint's actual clean/live feature SHA and exact CI;
these records cannot self-name their final commit or certify future pushes.
The actual owner PR27 merge/local main is ccbf6626d280d3864622c7d442e1ff68e4e55397.
[PR28](https://github.com/PikkuJanne/WinImgNormalizer/pull/28) needs separate owner
merge action after its actual checks; do not infer that permission from release
approval. Use safe clean fast-forward synchronization after any actual owner merge.
No next implementation task is selected.

Windows 10 testing is expressly outside project scope (D43); do not request it or
make it a release gate. Preserve D35 live-UNC exclusion and historical T067/D37
consent at 5a32fc6533c6f9fae1d36c271b515d1d1dd4d8ad. Website deployment/public
owner-media reuse remain unapproved and need actual repository/provider context.
Preserve the separate WinImgNormalizer-main snapshot and all raw ignored evidence.
