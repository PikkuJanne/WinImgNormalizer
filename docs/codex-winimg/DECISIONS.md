# Decisions and owner gates

| ID | Decision | State / rationale |
|---|---|---|
| D01 | Keep local PowerShell + batch architecture, existing entry points and positional calls. | Owner constraint; not a rewrite. |
| D02 | Website presents documentation/releases only. | Owner constraint; no server-side processing. |
| D03 | One task per thread, one feature branch and draft PR by default. | Execution plan; keeps a small script's edits sequential. |
| D04 | First hardened version rejects destination-inside-source before writes. | Chosen safe implementation boundary; document user-visible guard. |
| D05 | Preserve simple unique names; deterministic suffixes only for real namespace conflicts. | Accepted safety fix; test secondary collisions. |
| D06 | Register duplicates only after successful finalization and require equal length. | Accepted safeguard; still a heuristic, not content identity. |
| D07 | First displayed frame / first TIFF page; never implicit JPEG numbering. | Explicit implementation of intended single-image behavior. |
| D08 | Convert profiled colour before stripping; white background remains. | Correctness fix; untagged ambiguity must be documented. |
| D09 | Retain 1,048,576-byte cap, scale sequence and best-effort policy initially. | Algorithm tuning beyond corrections needs measured owner review. |
| D10 | Skip/report traversed reparse points; use run-owned temporary paths. | Safety default; no hidden following of links or broad cleanup. |
| D11 | Exit 0 complete, 1 setup/global failure, 2 warning/partial, 130 controlled cancellation. | Proposed normative application contract, subject to real host tests. |
| D12 | Use Python only as an optional development-bundle helper. | No new normalizer runtime dependency. |
| D13 | Tags, merges, releases, settings changes and website deployment are separate approvals. | Publication is not authorized merely by creating this bundle. |
| D14 | Subjective quality/changed defaults require owner evidence. | Agent cannot approve its own aesthetic output for the owner. |
| D15 | Preserve the non-Git source snapshot; use a fresh clone in a new empty directory for M0-T01. | Actual workspace differed from the checkout assumption. No checkout was found; all seven snapshot blobs matched the baseline. The START_HERE fallback preserved existing files and avoided inventing history. Verified clean clone and live remote at the same baseline; no application scope changed. |
| D16 | Record the inherited Windows helper-fixture failure without changing the immutable bundle. | The external suite assumes case-sensitive fixture paths in test 23. An explicit duplicate-manifest control passed on Windows. Integrity, additive import and plan checks pass; the suite remains 45 passed, 1 failed, 2 skipped. No application coverage or owner acceptance is inferred. |

Append a decision with problem, alternatives, choice, compatibility effect, tests and
owner approval when required. Do not use a new decision to silently drop an accepted
improvement or turn a failing requirement into an optional test.
