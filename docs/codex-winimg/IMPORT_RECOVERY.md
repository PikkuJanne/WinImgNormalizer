# Import recovery and manual integration

The external bundle is immutable. The repository progress records are meant to
evolve. Never restore them from the ZIP over newer work.

## Existing AGENTS.md or other conflicts

A conflicting existing file blocks automatic import. Read and preserve it. Do not
delete/rename it just to get a clean importer run. Review the proposed instructions
against applicable parent/root/nested instructions. Manually integrate only the
relevant new guidance without weakening existing constraints. Existing project
state remains authoritative; merge newer task/progress evidence rather than reset
it. Resolve genuine conflicts explicitly and record the decision.

For a manual import, first verify the external manifest, repository identity, actual
clean feature branch and reviewed HEAD. Use BUNDLE_MANIFEST.json as an exact allowlist:
only files with an install_path may enter the checkout. Add absent allowed files;
leave identical files alone. For every differing file, review the differences and
perform a narrow intentional merge. Never copy the whole ZIP root or replace runtime
source, README, license, .git configuration, working media or unrelated tooling.
After integration inspect the actual diff and explicitly stage intended paths only.
The importer itself never commits or pushes; finish the documented Git checkpoint.

Python is optional development tooling. When it is not installed, do not silently
install it or relax machine-wide policy. Codex can use existing local tools to hash
files and perform the same reviewed additive operation. The safety rules do not
become optional because the convenience helper is unavailable.

## Interrupted import

An unexpected filesystem failure can leave some newly created handoff files.
The helper never attempts a broad rollback or deletes unknown user files. Inspect
those files and compare them with the manifest before deciding what remains. Do not
clean or reset the worktree blindly. A repeated apply may refuse the now-dirty tree;
that is intentional. Complete or preserve the partial import with a reviewed manual
operation and a scoped commit, then rerun structural validation.

## Newer code or prior progress

The reviewed commit is historical, not an installation target. Diff the current
source against relevant review anchors, mark already-fixed items only with real
evidence, and adjust bounded tasks while retaining all 18 accepted requirements.
Do not regenerate the registry from the initial bundle to overwrite completed work.

## Verification after import

Run the plan checker, inspect staged paths, ensure runtime source was unchanged in
M0-T01, then commit/push the dedicated branch and verify its current advertised SHA.
Unexpected dirty state, divergence or remote identity mismatch is a blocker, not a
reason for force push or a default-branch write.
