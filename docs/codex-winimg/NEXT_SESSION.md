# Next session handoff

Next task: **M1-T01 — Validate inputs, root paths and ImageMagick before work**.
M0-T03 is complete. Begin only that next task in a fresh thread.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening. The configured
WinImgNormalizer-main folder is the preserved non-Git snapshot. Read AGENTS.md,
STATUS.md, TASKS.json, GIT_WORKFLOW.md and tasks/M1-T01.md with its referenced specs.
Recheck clean state, canonical fetch/push URLs and the exact advertised feature SHA.
The final evidence checkpoint/sync is reported in the preceding response and draft
PR #3; the tested runtime/tests revision is 3b80ee94b8471a22560ea90a7583283c1613195c.

M0-T03 introduces only a callable orchestration/positional boundary and internal
OutputParent/ProcessRunner injection. Dot-source is now safe after its isolated
import probe. Outer script exits; callable usage/empty-tree paths return. Public
source-only/source-plus-cap forms and batch launcher are retained. Production
preflight, collision, frame, duplicate, weak-output, colour and exit bugs are not
fixed by this seam task. Keep the smallest patch for M1-T01 and add its tests.

Actual desktop PS 5.1/PS 7 each passed 6/6 Pester tests; ten credibility controls reject
false assertions/zero/skip/discovery failures. An import exit 0 mutation also fails
safely. Four independent ordinary parity runs match the M0-T02 baseline hashes,
dimensions/tree/timestamps with sixteen JPEG decodes/twelve videos and preserved
sources. T006 uses only an instrumented outer destination call; real Pictures and
manual drag-and-drop are untested. CI passed the early Windows Server matrix and
deliberate seventh assertion; exact versions/job links are in M0-T03-ci.json.

Development commands from the repo root, separately in each fresh Windows shell:

```powershell
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
.\tests\Invoke-Tests.ps1 -DeliberateFailure # expected exit 1 with one false assertion
```

The explicit bootstrap verifies Pester 5.9.1 and ImageMagick 7.1.2-32 archives, freshly
extracts into ignored scratch and checks manifest/executable hashes. Cached archives
are reused without downloading; normal runtime never bootstraps. GitHub exposed
competing tar executables; bootstrap now explicitly selects Windows System32 tar.
No installer, permanent PATH or persistent policy change. Child shells remove
inherited PSModulePath; never redefine USERPROFILE or test against real Pictures.

Evidence: M0-T03.json plus runner/parity/import-control/CI JSON. Private artifacts
remain ignored. The final record-only checkpoint binds tested files to the code
revision above; containing checkpoint sync requires live verification next session.
Keep one feature branch and draft PR #3. No merges, tags/releases, settings changes,
website deployment or subjective owner approval are authorized by this handoff.
