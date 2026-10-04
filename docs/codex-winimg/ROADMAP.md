# Milestones and thread-sized tasks

Use **one task per Codex thread**, not one entire milestone. The default plan is sequential to avoid overlapping edits to a small script. Task dependencies live in TASKS.json. Keep one dedicated feature branch throughout; a new thread does not require a new branch.

## M0 — Preparation and regression foundation

Characterize safely and establish test seams before behavioral fixes.

- **M0-T01** — Reconcile workspace and import the handoff
- **M0-T02** — Characterize the current application with safe fixtures
- **M0-T03** — Introduce minimal test seams and early Windows checks

## M1 — Safe output

Stop omissions, clobbering, bad success detection and unsafe traversal.

- **M1-T01** — Validate inputs, root paths and ImageMagick before work
- **M1-T02** — Protect traversal and allocate unique run directories
- **M1-T03** — Plan collision-free output names before conversion
- **M1-T04** — Validate temporary outputs and finalize without overwriting
- **M1-T05** — Make heuristic duplicate handling deterministic and success-based
- **M1-T06** — Close the output-safety milestone

## M2 — Correct conversion

Settle frames, profiles, literal filenames, byte caps and native errors.

- **M2-T01** — Implement deliberate first-frame and first-page handling
- **M2-T02** — Convert colour profiles before metadata removal
- **M2-T03** — Make filesystem and native filename handling literal-safe
- **M2-T04** — Clarify byte caps and benchmark quality without silent default changes
- **M2-T05** — Capture useful errors and remove indiscriminate fallback
- **M2-T06** — Close the conversion-correctness milestone

## M3 — Reliable operation

Resilient discovery, budgets, cancellation, reporting and Windows acceptance.

- **M3-T01** — Recover from inaccessible folders and expose ancillary failures
- **M3-T02** — Bound processing and own native process lifetime
- **M3-T03** — Implement and test clean cooperative cancellation
- **M3-T04** — Reconcile counts and add useful safe reports
- **M3-T05** — Propagate exit codes and harden the batch launcher
- **M3-T06** — Complete Windows compatibility, static checks and security CI
- **M3-T07** — Obtain owner acceptance of the familiar workflow **Owner gate.**

## M4 — Public distribution

Versioned packages, accurate docs, website content and approved publication.

- **M4-T01** — Add one version source and release notes
- **M4-T02** — Build a portable release ZIP from reviewed tracked content
- **M4-T03** — Publish-ready documentation for behavior and limitations
- **M4-T04** — Prepare website content and download metadata, not a backend
- **M4-T05** — Audit the complete release candidate
- **M4-T06** — Perform only explicitly approved publication operations **Owner gate.**

## Stop rules

Stop on unclear ownership of local changes, unexpected remotes, divergence, a failing required gate, unavailable approval, or untrusted test data. Preserve state and report a blocker. Do not keep working in an ever-longer thread to finish a milestone. When a task is still too large, split it in the registry with explicit dependencies and keep original requirement/case coverage; do not silently drop acceptance criteria.
