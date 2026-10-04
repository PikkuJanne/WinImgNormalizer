# WinImgNormalizer — project instructions

## Product boundary

Improve the existing Windows PowerShell/ImageMagick tool; do not rewrite it. Keep
`WinImgNormalizer.ps1` and `WinImgNormalizer.bat`, folder drag-and-drop, local-only
processing, original-file preservation, mirrored subdirectories, JPEG derivatives,
white alpha background, image metadata stripping after correct colour handling,
and unchanged video bytes. Retain the two existing positional invocation forms and
Windows PowerShell 5.1 plus PowerShell 7 compatibility.

The website presents and distributes releases. No server-side processing, uploads,
telemetry, new accounts, GUI framework, mandatory content hashing, or parallelism.
Do not overwrite sources, modify global ImageMagick policy, relax machine-wide
execution policy, or auto-install software. Use approved synthetic test data only.

## Read only what is needed

Start with `docs/codex-winimg/STATUS.md`, `TASKS.json`, `NEXT_SESSION.md`, and the next
dependency-ready task card. Load its referenced specifications and tests, not every
card. `REQUIREMENTS.md` maps all 18 accepted review items. The review baseline is an
observation, not an instruction to reset newer code. Preserve applicable parent,
root, and nested agent instructions; never replace existing rules blindly.

## One task per thread

Implement only one task, add its regression tests, validate, update handoff records,
and stop. Small test seams are allowed; wholesale restructures are not. Keep new
application dependencies to a minimum; development tools do not become runtime
requirements. Do not import or execute the baseline application through dot-sourcing
until a testable entry-point boundary exists.

## Git and approvals

Use one dedicated feature branch and one draft PR for this hardening series unless
the owner chooses a different strategy. Inspect local changes and fetch/push remote
identity before work. Do not reset, clean, stash, rebase, force-push, silently pull,
overwrite differing files, or stage unrelated work. Normal reviewed commits and
normal pushes to the verified feature branch are authorized by the owner startup
prompt, subject to actual tool permissions. Default-branch pushes, merges, tags,
release publication, settings changes and website deployment need separate approval.
Stop on unexpected dirty state, remote mismatch or divergence; preserve the evidence.

After each task: validate, explicitly stage intended paths, inspect the staged diff,
commit, push, and compare `HEAD` with `git ls-remote` for the exact remote branch.
Clean worktree + equal SHAs means synchronized **at the recorded observation time**.
A push rejection or inaccessible remote never means synchronized. CI passing and
owner acceptance are separate gates, not synonyms for Git sync.

## Durable handoff and honest evidence

Update `STATUS.md`, `TASKS.json`, `NEXT_SESSION.md`, and `SESSION_LOG.md` every task.
Store sanitized task evidence under `docs/codex-winimg/evidence/`. Record code commit
or per-file hashes, environment, commands, test counts, skips, limitations and open
blockers. Keep raw test artifacts and private paths out of tracked files.

A checkpoint cannot contain its own final commit SHA. Follow `GIT_WORKFLOW.md`:
evidence commits may refer to the preceding implementation commit; observe the final
checkpoint SHA after committing and report it externally. Re-verify it next session.
Do not amend repeatedly to chase a self-referential hash or claim a future push worked.

Windows compatibility requires actual Windows results. An unavailable codec is a
skip only for explicitly optional capability tests, never a pass. No screenshot,
benchmark, CI, signature, owner acceptance or runtime reproduction may be fabricated.
The bundle's helper tests do not test the image-normalization application.

## Engineering defaults

Prefer literal-safe paths, no-overwrite finalization, precise process results and
bounded resource use. Do not confuse PowerShell quoting with ImageMagick filename
parsing. Preserve privacy-sensitive colour profiles until conversion is complete,
then apply the documented metadata policy. Stable functionality and readable comments
matter more than broad configuration. Add tests before or with each behavioral fix.
