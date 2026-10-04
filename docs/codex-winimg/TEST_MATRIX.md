# Test implementation and acceptance matrix

`TEST_CASES.json` contains 75 specified cases. Initial status is `not_run` for every
case. These are requirements for Codex to implement and execute, not a delivered
passing application test suite. Each task card names its cases. Add tests with fixes.

## Layers

| Layer | What belongs here | What it does not prove |
|---|---|---|
| Unit | Naming plan, containment, argument validation, duplicate bookkeeping, result accounting, retry decisions, CSV neutralization and process-result classification with Pester. | Real native argv, installed codecs, Windows filesystem semantics or actual GUI drag-and-drop. |
| Integration | Real ImageMagick with synthetic images, full JPEG decode, output bytes/metadata/frames, no-overwrite finalization, copied video hashes. | Every codec build, all Windows host interrupt behavior, or subjective owner quality acceptance. |
| Windows shell/launcher | Native Windows PS5.1 and supported PS7, `.bat`, path characters, Unicode, ACLs/junctions, Ctrl+C and exit codes. | A Windows Server CI run alone does not prove a Windows 10/11 desktop experience. |
| Manual owner | Everyday drag-and-drop with disposable approved copies; appearance, clarity, familiar workflow and clean release extraction. | Automated universal correctness or consent to unrelated default changes. |
| Packaging/security | Allowlist ZIP contents, source/version/checksum match, no secrets/private data, workflow permission review. | Cryptographic publisher authenticity merely because a checksum exists. |

## Required execution environments

Unit/integration tests must run under Windows PowerShell 5.1 and a currently
supported PowerShell 7 on Windows with the same documented supported ImageMagick
baseline. Record exact host, OS edition/build, architecture, locale, ImageMagick
version/delegates, Pester and analyzer versions. Pin development dependencies and
verify their provenance rather than installing floating latest versions silently.
PowerShell 7 on Linux may provide additional evidence, not replace either Windows
column. Do not introduce PowerShell 7-only APIs into shared helpers without a 5.1
path. A runner failure is failed infrastructure, not a skipped success. [P1,P2,T1]

Add an early limited Windows workflow in M0-T03 and expand it as cases are fixed.
By M3-T06 all mandatory automated cases must be present and meaningful. A legacy
reproduction suite can be optional/manual; do not use expected failure of old code
to pretend corrected code passed. Pester discovery count must be nonzero and
unexpected skipped mandatory tests must fail the acceptance gate. [T1]

Capability-specific tests carry a precise reason. `optional-codec` means local
absence can be represented explicitly; public supported-format coverage still
needs an actual codec-enabled run. `windows-capability` cases such as UNC may use
a separate controlled target. Missing that environment limits the support claim.
Owner tests require actual owner evidence and cannot be approved by the agent.

## Evidence per execution

Record task/case IDs, UTC time, exact command, tool versions, tested commit or code
file hashes, result counts, log/artifact location, hashes for retained artifacts and
known limitations. Test failures remain failures; do not reduce assertions or mark
an inconvenient test optional to obtain green CI. Sanitize tracked summaries; raw
local paths/media and crash dumps remain outside the repo or in ignored scratch.

A normalizer source-preservation test should verify file content hashes plus source
creation/modified timestamps. OS access-time updates are not a guarantee the tool
can prevent. For copied videos compare byte hashes. For JPEGs validate decoding,
format, dimensions and metadata plus tolerance-based image checks, not only length.

The optional `tools/codex-winimg/handoff.py plan --repo .` checks task/case coverage
and dependencies; it does not execute image-normalization tests. The self-tests
shipped in the external bundle validate the importer/sync helper only.
