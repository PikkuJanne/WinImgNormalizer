# Project status

Updated: 2026-10-05 — M2-T06 completed; conversion-correctness milestone closed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 11ee247154a00d79aaa49a438d1f526457b22fdf
Next task: **M3-T01 — Recover from inaccessible folders and expose ancillary failures**
Task progress: **15 / 28 accepted**
Specified cases exercised: **T001-T046 (46 / 75); T001-T004 are characterization**
Pester suite: **271/271 passed in each desktop shell, zero skipped**
Failure controls: **272 total, 271 passed, one intended failure; native exit 1**
CI: **implementation push/PR Windows matrices passed, including controls**
Containing evidence checkpoint synchronization: **pending_verification**

## Current behavior and comparison

T046 exercises colour handling, frame/page selection, literal names, orientation and
exact byte-cap sizing together through the actual pinned native runner on approved
synthetic inputs. The combined checks preserve source bytes and recorded
creation/modified times, byte-identical video copying, deterministic output mapping and
earlier outputs across repeated runs. Concrete fixture/cap counts and outcomes are
recorded below rather than inferred from the prior task.

Each recorded mixed-tree batch retained 12 fully decoded JPEGs and 1 byte-identical
video, skipped 1 heuristic duplicate, and returned application code 0. The omitted
default cap was 1,048,576 bytes; the explicit exact cap was 65,537 bytes. Source state
and the complete earlier output/log tree stayed unchanged. Focused development and
immutable-I observations remain separately labeled in the comparison evidence.

T046 — passed: Both default1048576-byte and exact65537-byte actual unmocked mixed-tree
cases passed in every fresh PS5.1/PS7 clean-I normal/control invocation, plus focused
development. Eight exact-I comparison runs bind27 gate source files and actual
tool/fixture/output/log/buffer bytes; only T007 deliberately failed in controls.
All12JPEG+1video+1duplicate per normalization and mandatory
frame/ICC/privacy/literal-path/exact-cap/source/prior-run checks passed.

## Verification and handoff

Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1: each normal run passed 271/271 with native exit 0 and zero skipped or incomplete
tests. Each control had 272 total, 271 passed, one exact T007 deliberate failure and
native exit 1. Summary/XML, parent native records and console hashes bind 27 tested
source paths, including the integrated suite and all eight independent colour/reference
assets, to raw checkout and separate Git blobs. Only verified text CRLF-to-LF
normalization is allowed; ICC bytes match exactly. Persistent execution policies were
unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37272249221); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37272253545).
Evidence: evidence/M2-T06.json and evidence/M2-T06-ci.json.

Actual development observations, comparison artifacts and exact-revision acceptance
gates are retained in evidence/M2-T06.json and SESSION_LOG.md. Earlier task histories
remain in their original task evidence. Comparison recipes, native/tool versions and
tolerances belong to this tested revision; automated synthetic comparisons do not
establish subjective owner quality approval.

Started clean at c062b835c4acb2734f42f87d81b8e47164af827a, equal to the live feature
branch; owner-merged PR #14 main e24461bb22155902984784aba7e30d4db0e63147 had the same
tree. Continued without changing checkout/history. Successor draft PR #15 contains this
task. Implementation synchronization is a past exact-SHA observation; this record-only
checkpoint's SHA/live synchronization remains pending_verification until observed
externally after normal commit/push.

JPEG derivatives are lossy and do not replace preserved originals. The comparison
evidence applies to the approved synthetic corpus and pinned decoder/profile behavior;
codec/build and reference-intent qualifications remain explicit. Inaccessible subtrees
and ancillary reporting are next M3-T01; general deadlines, cancellation and native
descendant lifetime remain later M3 gates. Owner acceptance, releases and deployment
remain separate.

This session stops at M2-T06. Begin M3-T01 only when requested.
