# Next session handoff

Next task: **M0-T02 — Characterize the current application with safe fixtures**.

Use the actual Git checkout `C:/projects/WinImgNormalizer` and the existing
`codex/winimg-hardening` branch. The original open project directory was a non-Git
source snapshot and was preserved unchanged; do not continue there accidentally.
Read applicable AGENTS instructions, STATUS.md, TASKS.json, GIT_WORKFLOW.md and
`tasks/M0-T02.md`. Implement M0-T02 only in a fresh thread.

M0-T01 imported the 56 manifest-allowed governance/helper files. Its starting commit
and live `main` were `4b918da50639d8fc7e3fddc6d6b3678880ea4098`; all seven original
tree blobs matched. No application file, asset or runtime behavior was changed.
The latest final commit and post-push observation are in the preceding session's
response and the project's draft PR. Inspect clean state, both remote URLs and
live refs again; independently compare HEAD to the advertised feature-branch SHA
before treating it as synchronized. No reset, stash, clean, rebase or force push.

M0-T01 evidence is `evidence/M0-T01.json`. Integrity/import/plan checks passed.
External helper suite: 45 passed, 1 Windows fixture failed, 2 symlink tests skipped;
the explicit colliding-manifest control passed. Preserve the immutable bundle.
Python 3.14.6, PowerShell 7.6.5 and Windows PowerShell 5.1.26100.9444 are available.
Only Pester 3.4.0 was found; ImageMagick and ExifTool were absent from PATH. Establish
M0-T02's needed tooling through the task's authorized process; do not auto-install
software or relax machine-wide policies. No application tests have run. CI was
unconfigured at baseline; sync, CI and owner acceptance remain separate facts.

The website continues to present and distribute the local tool, with no uploaded
image processing. Merging, tags/releases, settings and deployment remain owner gates.
