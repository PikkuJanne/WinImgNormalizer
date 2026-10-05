# Next session handoff

Next task: **M3-T07 — Obtain owner acceptance of the familiar workflow**.
M3-T06 is complete; M3-T07 has not started.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening. WinImgNormalizer-main
remains the preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json,
GIT_WORKFLOW.md, tasks/M3-T07.md and templates/OWNER_ACCEPTANCE.md. Recheck clean
state, canonical fetch/push identity, exact live feature SHA and
[draft PR21](https://github.com/PikkuJanne/WinImgNormalizer/pull/21) before editing.
Tested implementation is 17a405bfe0150607486248e845f93ba82e1d0d52; independently verify the
containing record-only checkpoint. Preserve WinImgNormalizer.bat -text and exact
CRLF bytes in the checkout, Git blob and source archive.

M3-T06 requires all 23 suites, 61 corrected case IDs and ten actual codec outcomes;
no zero discovery/skip/unrun/inconclusive result can silently pass. Desktop Windows 11
Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5 passed 468/468 each; both full controls were
469/468/1 with native 1. Pinned Pester 5.9.1/ImageMagick 7.1.2-32 and analyzer 1.25.0
are development dependencies. Scoped normal analyzer results had 7 files/0 findings;
both unsafe-expression controls had one exact finding/native 1.

Final implementation hosted push 37333526329/PR 37333528830 passed all four normal
and assertion/analyzer controls. Actual Server 2025 image win25-vs2026
20260925.250.1, runner 2.337.0, PS 5.1.26100.33438/PS 7.6.6 differ from desktop.
Sanitized desktop/CI evidence binds native exits, hashes, exact commits, 45 source
paths, artifact contents and synthetic merge provenance. Earlier failed and
successful superseded implementations, including I3 lossy timestamp qualification,
remain separate. See evidence/M3-T06.json and evidence/M3-T06-ci.json.

The final exporter rejects nested private values and preserves UTC fractional time.
CI remains contents:read with reviewed full action pins, credentials disabled and
two fixed synthetic JSON uploads. Analyzer excludes Pester DSL/legacy/vendor code.
Live UNC/network-share validation and support are out_of_scope by the owner on
2026-10-05 (D35): no server is available, and no live-share verification is pending.
Keep local long-path, lexical drive/UNC-root and local identity tests mandatory.
Preserve earlier limited localhost results and `not_run` observations; no passing
live network-share result is inferred. See evidence/M3-T06-scope.json. No Linux
substitute or checksum-signature claim exists.

M3-T07 must prepare concrete owner comparisons/checklist using approved disposable
copies, then obtain actual named-commit workflow and quality acceptance. Do not
manufacture owner approval or import private media into the repository. If approval
is absent, prepare evidence and leave that task awaiting_owner per its task card.
No merge/tag/release/settings/deployment approval is inferred. This session stops
at M3-T06; begin M3-T07 only when requested.
