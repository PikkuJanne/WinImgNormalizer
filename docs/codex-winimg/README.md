# Codex hardening workspace

**Start with STATUS.md, TASKS.json and NEXT_SESSION.md.** Work on one task at a time.

`REQUIREMENTS.md` is the 18-item traceability map. `ROADMAP.md` gives the milestone
order. Each `tasks/Mx-Txx.md` is a self-contained execution card with scope,
acceptance, test cases and stop conditions. `IMPLEMENTATION_SPEC.md` settles
technical boundaries without replacing the current application architecture.

`TEST_CASES.json`, `TEST_MATRIX.md` and `FIXTURE_SPEC.md` define tests to implement.
They are not claims of completed tests. `GIT_WORKFLOW.md` defines synchronized
checkpoints. `RELEASE_AND_WEBSITE.md` and `SECURITY.md` govern distribution and scope.
`REVIEW_BASELINE.md` distinguishes static findings from runtime cases to reproduce.
External primary references are indexed in `SOURCES.md`. Import conflicts and
recovery are covered in `IMPORT_RECOVERY.md`.

Mutable records: STATUS.md, TASKS.json, NEXT_SESSION.md, SESSION_LOG.md, DECISIONS.md,
ACCEPTANCE.md and sanitized evidence files. Do not overwrite them with fresh bundle
copies after work has started. Immutable initial-bundle checksums are for the
extracted bundle, not for these intentionally evolving repository records.

From the repository root, the optional Python 3.9+ development helper can run:

```text
python tools/codex-winimg/handoff.py plan --repo .
python tools/codex-winimg/handoff.py status --repo .
python tools/codex-winimg/handoff.py status --repo . --check-remote
```

The status command reads Git and, with `--check-remote`, contacts the configured
remote using `git ls-remote`. It never fetches, pulls, switches, commits or pushes.
Its JSON distinguishes a clean local checkpoint, a remote match, and true observed
synchronization. It does not infer CI or owner approval.
