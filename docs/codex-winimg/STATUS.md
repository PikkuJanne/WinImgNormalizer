# Project status

Updated: 2026-10-04 — M0-T01 governance import and workspace reconciliation.

Repository: PikkuJanne/WinImgNormalizer
Review baseline: `4b918da50639d8fc7e3fddc6d6b3678880ea4098`
Actual checkout: `C:/projects/WinImgNormalizer`
Branch: `codex/winimg-hardening`
Reviewed starting HEAD and advertised `main`: `4b918da50639d8fc7e3fddc6d6b3678880ea4098`
Next task: **M0-T02**
Task progress: **1 / 28 accepted**
Application test cases: **0 / 75 executed**
GitHub sync: **pending_verification for this checkpoint**
CI: **no workflows or runs observed at the starting revision**
Owner acceptance: **not requested / not granted**
Merge/release/website deployment: **not authorized by bundle creation**

## Current blocker

No blocker to the M0-T01 import. The original open project was a seven-file source
snapshot without Git metadata. No checkout was found in likely development or
managed-worktree locations; the saved project also reported a non-Git directory.
All seven snapshot blobs matched BASELINE.json. The snapshot was preserved and a
fresh clone was made into the previously absent checkout path above, following the
documented fallback. No existing source, instructions, history or progress was replaced.

The clean clone and live remote `main` matched the baseline; no newer code required
task changes. Both origin fetch and push URLs are the single canonical
`https://github.com/PikkuJanne/WinImgNormalizer.git`; no push mappings, mirror setting,
URL rewrites or custom hooks path were observed. The hardening branch was absent
locally and remotely and was created from that reviewed starting point.

Environment: Windows NT 10.0.26300.0; PowerShell 7.6.5; Windows PowerShell
5.1.26100.9444; Python 3.14.6; Git 2.56.0.windows.1; GitHub CLI 2.97.0.
Pester 3.4.0 is discoverable in both shells. ImageMagick and ExifTool were not found
on PATH; neither was run or installed. Machine-wide policies were not changed.

Bundle integrity passed for 63 files. Preview and apply added only the 56 allowed
files; all matched manifest hashes before the intentional handoff updates. The
task graph passed (28 tasks, 75 cases, all 18 review items). External helper tests:
48 run, 45 passed, 1 failed, 2 skipped. The failed case assumes a case-sensitive
fixture filesystem; an explicit colliding-manifest control passed. Symlink tests
skipped because creation was unavailable. This suite is not application evidence.
See `evidence/M0-T01.json` for commands, exact content hashes and limitations.

## Last observed synchronization

The feature branch did not exist on the remote at initial inspection. The final
checkpoint's push and independent SHA comparison are pending when these files are
written. Report the actual post-commit observation in the session response and draft
PR; re-check it next session. Do not embed this containing commit's own SHA here.

## Next action

Finish the scoped M0-T01 commit/push/remote check, maintain one draft PR, then stop.
Start a fresh thread for M0-T02 using this actual checkout. Verify live sync first;
read NEXT_SESSION.md and the task card. Application/runtime validation remains unrun.
