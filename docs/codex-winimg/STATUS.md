# Project status

Updated: 2026-10-05 — M3-T05 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 5df7a83ed3fd340479972366f95a59e8e51b1b41
Next task: **M3-T06 — Complete Windows compatibility, static checks and security CI**
Task progress: **20 / 28 accepted**
Specified cases exercised: **T001-T063 (63 / 75); T001-T004 are characterization**
Pester suite: **392/392 passed in each desktop shell, zero skipped**
Failure controls: **393 total, 392 passed, one intended failure; native exit 1**
CI: **implementation push/PR Windows matrices passed, including controls**
Containing evidence checkpoint synchronization: **pending_verification**

## Current behavior and verification

The launcher keeps Windows PowerShell, -NoProfile, process-only Bypass and the
familiar pause. It clears an inherited ERRORLEVEL shadow inside SETLOCAL, saves the
actual child code immediately, prints success/warning/error/cancelled text and returns
that saved code after pause. Zero/extra arguments, including an explicit empty extra,
are rejected before starting the script. Literal path forwarding and trailing-separator
normalization remain. The file-specific WinImgNormalizer.bat -text attribute preserves
the BAT's exact CRLF bytes in the Git blob, checkout and source archive, supporting
reliable native CMD label lookup.

The command catches ordinary unexpected run-level failures, emits bounded escaped
diagnostics and returns 1. Handled cancellation remains 130; PipelineStoppedException
propagates. Completed outputs and existing reports survive a run-level failure.

On actual Windows 11 Pro build 26300, Windows PowerShell 5.1.26100.9444 and PowerShell
7.6.5, Pester 5.9.1, final implementation 5df7a83ed3fd340479972366f95a59e8e51b1b41 normal gates passed
**392/392 in each shell**, zero skipped or unrun. Each control had 393
total, 392 passed, exactly one intended T007
failure and native exit 1. The parent recorded native exits; summary/XML/console hashes
and verified raw/Git source bindings are in evidence/M3-T05.json. Persistent policies,
PATH, profiles and user identity were checked without changing them.

The initial clean implementation 5cff9f22e0149d3acb6f339494094a93c67aeb5c passed
391/391 in each desktop shell; each desktop control had 392 total, 391 passed, one
intended T007 failure and native exit 1. Its hosted Windows PowerShell 5.1 push/PR
normal gates each failed 22 pause assertions (391 total, 369 passed), while hosted
PowerShell 7 normal gates passed. The reason for the hosted/desktop difference is unconfirmed;
local detached-process probes did not reproduce it. These earlier observations and
development failures remain preserved separately from the final implementation.

The repair changes only development test fixtures and their source/compiler evidence
bindings. Production PowerShell and BAT bytes are unchanged between the initial and
final implementations. A shared fixture creates a hidden private console, assigns
itself to a private kill-on-close job before CMD starts, and captures the parent's
stdio pipes before console allocation. Its readiness marker identifies the actual
owned CMD. Tests wait for the real driver to exit, observe CMD blocked at pause, then
supply two key records only to the verified controller/CMD console. Timeout cleanup
kills the exact fixture handle; closing its private job stops its own descendants.
A no-pause command must exit promptly with native 37 while parent stdin stays open,
preventing the fixture's key reader from manufacturing pause evidence. The fixture
uses the existing Windows Framework compiler only for development tests; measured
compiler version/hash and fixture source/executable bindings are recorded locally and
in sanitized evidence. It adds no application runtime requirement.

Final implementation hosted push/PR matrices passed all four jobs and controls: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37322686828); [pull_request CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37322698850).
Evidence: evidence/M3-T05.json and evidence/M3-T05-ci.json.

T061 — passed: Exact BAT bytes with a synthetic child observer prove actual Windows
codes 0/1/2/130/37 and unavailable-host 9009, truthful messages and the blocked pause
after the child exits in an explicitly owned console. Readiness checks exclude fixture
startup from early usage/setup pause observations; the no-pause control verifies that
open parent stdin cannot create a false wait. Separate safe -File drivers exercise the real Command for empty,
ignored, usage/setup, partial and controlled global/cancellation outcomes, retain prior
video hashes and preserve source hashes/times.

T062 — passed: Actual native one-folder argument flow, script resolution and count
rejection under spaces, Unicode, percent/bang expressions, ampersands, parentheses,
brackets and trailing separators. Empty extra arguments reject with the intended
message and pause. Path cases release the key only after observing the blocked pause;
early-key timing failures are preserved. Earlier LF label failures are preserved;
exact CRLF blob/checkout/archive verification passes.

T063 — passed: Real legacy child host/process policy and unchanged same-edition
persistent settings/PATH/profile/user context. Different PowerShell editions expose
different policy views; baselines are compared within the same host edition.

## Handoff qualifications

Started clean at 2ec018decc833d4b4004818f1175058bd61c11a0, equal to the live feature
branch. Owner-merged PR #19 main fa2672f42be3f51cf2bfd0b00cb47ae0308d44b6 has the same
tree. Continued without checkout/history changes. Successor draft [PR](https://github.com/PikkuJanne/WinImgNormalizer/pull/20) contains
M3-T05. This record-only checkpoint's final SHA/live synchronization remains
pending_verification until independently observed after normal commit/push.

Automated Windows launcher argument flow is verified. A physical Explorer gesture,
keyboard interaction and subjective owner workflow approval remain M3-T07. The safe
drivers use internal owned scratch output and dependency paths; no real Pictures or
private media are processed. Controlled cancellation is distinguished from actual
Ctrl+C/forced-exit tests in the existing suite. Stopped pipelines, terminal closure
and force termination may bypass reporting, cleanup and launcher pause, and have no
unconditional exit130 guarantee. Caller-side CMD percent/bang expansion can alter
arguments before the wrapper receives them.

JPEG/colour/frame/size defaults, source/no-overwrite safety, byte-identical videos,
report accounting and native ownership remain covered. No owner quality acceptance,
merge, release, settings change or deployment is inferred.

This session stops at M3-T05. Begin M3-T06 only when requested.
