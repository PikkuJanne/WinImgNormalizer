# Local and GitHub synchronization protocol

## Authorization and identity

This project uses one feature branch, normally `codex/winimg-hardening`, and one
draft PR. Starting a new Codex thread does not require a new Git branch. The owner's
startup prompt authorizes scoped commits and normal feature-branch pushes to
`PikkuJanne/WinImgNormalizer` and maintenance of that draft PR. It does not authorize
merges, default-branch pushes, release/tag publication, settings changes, website
deployment, force push, history rewriting, cleanup of unknown local work or uploads
of private media. Honor any actual Codex/tool permission prompts. [G1,G2,O1]

Inspect both fetch and push URL, not only the remote's name. The helper accepts
canonical GitHub SSH/HTTPS identity and rejects unexpected remotes, extra push URLs
and ambiguous mappings. A custom SSH alias may require documented manual identity
verification; never bypass the identity check merely because the remote is named
origin. Never print an embedded token in a report. Preserve existing Git config.

## At session start

Read the current instructions/status/task records, then inspect `git status`, the
branch, local HEAD, both remote URLs and current remote refs. Network reads/fetches
are allowed, but a fetch is not a pull or a permission to rewrite working files.
Do not start from a stale remote-tracking ref and call it current GitHub truth.

The baseline in BASELINE.json is historical. Current local code can be newer.
Reconcile newer source and existing work in a small audit note; adjust task scopes
without discarding completed changes. If local work is dirty, preserve it and stop
before import/application edits unless the owner has explicitly identified it as
part of this task. Do not automatically stash, reset, clean or stage it. The bundled
importer refuses any dirty tree for this reason. Existing different progress files
are never replaced with fresh `pending` copies.

If the clean selected feature branch already exists, resume it. If it does not,
create it from the reconciled current starting point. Do not create competing
branches blindly. A branch behind or diverged from its corresponding remote is a
reconciliation blocker; do not automatically merge, rebase or force push. A clean
fast-forward update can be performed only after verifying its exact intended source
and that it preserves this project, with no unrelated local work. Do not silently
pull at the beginning of every task.

## Task closure

Implement one task, run its tests and update the durable records. Review the full
diff, `git diff --check` and the exact intended paths before staging. Stage only
those paths, not `git add .` over unknown material. Inspect the staged diff and
confirm no media, secrets, generated outputs or unrelated changes are included.
Use a task-named commit message such as `fix: M1-T03 prevent output name collisions`.

For straightforward exact-revision evidence, use two small checkpoints when useful:

1. Commit implementation/tests as **I**, run tests on that clean code revision,
   and capture actual outputs and environment.
2. Commit sanitized evidence and updated task/handoff records as **E**, referencing
   I. Ensure E changes only evidence/governance, not the tested runtime/tests. Run
   the plan checker and any relevant documentation/package checks on E as well.
3. Push the selected feature branch normally and compare local HEAD (E) with the
   currently advertised remote branch SHA. Record final E in the session response
   and draft PR. Do not amend E merely to put E's own hash inside its contents.

A one-commit alternative is allowed when evidence records exact precommit tested
file hashes plus the base commit and the final committed contents are verified to
match those hashes. Never record `HEAD` as a tested version while uncommitted code
changes were actually under test without those hashes.

Example commands after the actual branch and remote have been verified:

```powershell
$branch = (git branch --show-current).Trim()
if ($LASTEXITCODE -ne 0 -or $branch -ne 'codex/winimg-hardening') {
    throw 'Unexpected branch: inspect before pushing.'
}
git diff --check
if ($LASTEXITCODE -ne 0) { throw 'Diff check failed.' }
# Stage only the individually reviewed paths for THIS task, then inspect and commit.
git push --no-follow-tags --set-upstream origin "HEAD:refs/heads/$branch"
if ($LASTEXITCODE -ne 0) { throw 'Push failed; checkpoint is not synchronized.' }
python tools/codex-winimg/handoff.py status --repo . --check-remote
if ($LASTEXITCODE -ne 0) { throw 'Remote synchronization not verified.' }
```

The explicit push disables automatic tag following; do not use `--all`, `--mirror`, or any force option. Inspect unusual remote push configuration before proceeding.

The helper performs read-only status and `ls-remote` checks. It does not push for
you or certify test/CI success. Existing system aliases/approval constraints take
precedence over the example's path assumptions. [G1,G2]

## What synchronized means

At the recorded observation time: correct repository/remote identity, selected
feature branch, no pending worktree changes, and local HEAD equals the SHA for that
exact remote branch. An absent remote branch, unreachable remote, rejected push,
extra local changes or different SHA is **not synchronized**. Ahead/behind remote-
tracking metadata alone is not sufficient. A local commit made after the comparison
requires a new comparison. CI and owner approval are separate facts.

The final commit cannot contain its own final hash or evidence of a push that has
not happened yet. Tracked STATUS/SESSION records therefore store a past observation
by timestamp and exact SHA, or state `pending_verification`. The session's final
response and PR can carry the post-commit observation. On the next session, verify
live state again and record the previous observed checkpoint if useful. Do not
create an endless chain of evidence commits trying to make a self-referential
current-SHA field true. A registry task `done` denotes accepted implementation;
project readiness additionally requires the separate sync/CI/owner gates.

## Failures and gates

If authentication/network access is unavailable, save the local checkpoint safely
and report `not_verified` or `push_failed`. Do not ask the user to paste tokens into
chat or store credentials in scripts. If GitHub rejects the push, inspect and report;
never append `--force`. CI pending/unavailable/failed remains explicit. When a Windows
case has not run, do not turn a Linux test into that evidence.

An unexpected hard stop leaves STATUS and NEXT_SESSION with the blocker, attempted
commands and next safe action. Do not manufacture successful closure just to get a
clean-looking record. Keep unsynchronized local commits identifiable for recovery.

## Merge and release

Only after explicit owner approval may the named PR be merged and a version tagged
and published. Re-check the actual merged/tagged code and rebuild or verify artifacts
when the revision identity/content changed; never attach an earlier unverified ZIP
under a newer tag. Update the local default branch only from a clean state with a
verified safe fast-forward, otherwise stop. Release/website operations are in M4-T06
and require their own precise authorization.
