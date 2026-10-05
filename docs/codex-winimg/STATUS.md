# Project status

Updated: 2026-10-05 — M2-T03 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 79eb1379797cc1bbe4bbcf3cc4083cf83b0e17e8
Next task: **M2-T04 — Clarify byte caps and benchmark quality without silent default changes**
Task progress: **12 / 28 accepted**
Specified cases exercised: **T001-T039 (39 / 75); T001-T004 are characterization; path/codec limits remain explicit**
Pester suite: **227 / 227 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **228 total, 227 passed, one intended assertion failed; exit 1 in each shell**
CI: **implementation push and PR Windows Server PS 5.1 / PS 7 matrices passed, including expected controls**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and evidence

Every native image count, selected colour inspection, source ICC extraction, JPEG
conversion and full final JPEG decode sets registry:filename:literal=true before
the input or output. The process-local registry guard protects percent expressions
in parent directories as well as filenames, and cannot be overridden by a source
filename:literal=false property. Owned neutral snapshots and candidates are
finalized by literal no-overwrite filesystem moves to the deterministic original
identity.

The real selected-input regression places distinct red and blue numeric
counterparts beside a percent-template snapshot. Without the guard, image:frames=0
can return the wrong neighbour with native exit 0; the guarded operations read the
intended source, preserve first-frame/page and ICC policies, and retain literal
final JPEG/video names. Opaque video directory creation uses the literal .NET
Directory API.

The destination boundary retains canonical whole-segment and reparse checks and
adds read-only 128-bit FileIdInfo directory identities for the source and each
existing output ancestor. Local-drive/localhost-share aliases that would place
output inside the source reject before enumeration, probes or allocation. Missing,
unsupported or zero identities fail setup rather than fall back to an unverified
lexical alias boundary.

The BAT localizes DisableDelayedExpansion and conditionally appends a dot after a
trailing slash/backslash, preserving folder/root identity through quoted Windows
PowerShell 5.1 argv. Actual direct positional calls and byte-for-byte copied BAT
routes preserve the tested spaces, apostrophes, ampersands, parentheses, bangs,
percent/brackets and Finnish/German/emoji names. Caller CMD expansion that happens
before BAT entry remains documented; single-quoted direct PowerShell is the
alternative for paired percent expressions.

Seventeen mandatory Paths cases combine real pixels and exact output identity,
tagged ICC, animated first-frame, video byte/time fidelity, external final-name
arrivals, actual launcher argv/application returns, physical directory identities,
lexical roots and supported long paths. Separate owned localhost C$ probes cover
direct UNC input/output, BAT UNC source and both alias rejection directions. Full
suite counts are read from clean-I results, with no unavailable UNC capability
folded into mandatory pass counts.

Public positional forms, standalone PS1/BAT operation, byte cap, six-scale
sequence, JPEG flags, first-frame/page omission reporting,
ICC/orientation/white-alpha/strip order, deterministic naming, transactions,
source preservation and success-only heuristic duplicate links remain the combined
regression baseline. Workflow, bootstrap, dependency pins and trusted profile
assets are byte-unchanged; no host/share/credential/persistent PATH or policy
configuration is introduced.

Decision D26 records the tested literal filename and path-capability policy. Source preservation, deterministic mirrored names, transactional
finalization, verified copied videos and success-only duplicate links retain the
M1 regression baseline. M2-T01 frame/page selection and M2-T02 ICC, orientation,
white-alpha and metadata policy remain combined regression gates. Actual HEIC
evidence remains a two-image still collection; timed animation, thumbnails and
auxiliary-image behavior are unverified. Colour references retain their declared
synthetic patch tolerances and CMYK A2B0 display-mapping limitation.

T037 — passed: Literal percent/%d/%03d/bracket/hash/at-sign/leading-hyphen names,
nested first-GIF-frame selection and opaque copied video retain exact output
identities and full JPEG pixels. Genuine ICC and filename:literal=false property
remain safe in percent ancestors. Real snapshot-delegating observer creates a
differently colored numeric counterpart; all native reads are real. External
final-name arrival is preserved.

T038 — passed: Actual Windows PowerShell 5.1 -File source-only/explicit-cap and
byte-for-byte copies of the frozen BAT under inherited CMD delayed expansion
preserve quoted spaces, apostrophes, ampersands, parentheses, bangs,
percent/brackets and Finnish/German/emoji characters. UTF8 argv receiver records
actual Desktop5.1 host and application return 0; BAT trailing separators retain
canonical folder identity.

T039 — passed: Actual 324-character local source file and longer intended output
succeed in both desktop shells with preserved source hashes/creation/mtime and
full JPEG validation. Native 128-bit directory identities agree for the same
physical folder, differ for siblings and fail closed for missing directories.
Existing owned localhost C$ supports direct UNC input/output and BAT UNC source;
local/UNC descendant aliases reject both directions before native probing/write.
Lexical root tests do not enumerate roots.

Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444
and PS 7.6.5, Pester 5.9.1: 227/227 in each
desktop shell, zero skipped. Controls each have 228 total/227 passed/exactly
one T007 false assertion/exit 1. All four summaries and actual native child exits
bind the same clean committed checkout, with raw file hashes and separate Git blob
hashes for verified CRLF/LF normalization. Persistent execution policies are
unchanged. Actual hosted Windows Server matrices passed all four normal/control
jobs: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37259886557) and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37259892278).
Evidence: evidence/M2-T03.json and M2-T03-ci.json. Literal-name matrices, actual
launcher/native observations, capability limits, retained history and artifact
hashes are explicit. Missing UNC or disabled-environment observations cannot be
inferred from a successful local path or lexical root test.

Eight uncommitted development Pester runs are retained separately from the clean-I
acceptance runs. The existing four-suite PS7 run had 111 total/110 passed/one
failed assertion/exit 1: a Preflight test observer still indexed the owned
snapshot at argument 2 after the explicit literal registry guard shifted it to
argument 4. Correcting this test seam preserved its source, JPEG and video
assertions. Paths PS7 attempt1 had 17 total/15 passed/two failed assertions plus
one failed container/exit 1: case-insensitive $Batch shadowed the batch path
before the BAT launched, and AfterAll generic List serialization failed. Renaming
the test variable and using ToArray corrected those test harness failures. Paths
PS7 attempt2 passed 17/17; PS5.1 attempt2 had 17 total/13 passed/four
failures/exit 1 because Get-Content decoded the UTF-8 receiver JSON as ANSI. The
actual receiver bytes and output preserved the characters; .NET ReadAllText
corrected the test inspection. Both attempt3 and both attempt4 runs passed 17/17.
Attempt4 strengthened real UNC output and alias checks. These summaries' raw
end-of-run source hashes remain observations, not proof of a clean committed
revision or exact execution binding. Genuine native percent-template output and
selected-input red-versus-blue substitution, baseline BAT bang/trailing-separator
corruption, and local/UNC containment gaps were preserved as research/prototype
observations and fixed before implementation acceptance. Caller CMD paired percent
expansion remains a documented limitation. Invalid preliminary probe
argv/driver/pixel inspection and preexecution parser mistakes are qualified
separately, with no application/acceptance pass inferred.

## Handoff and remaining scope

Started clean at ecd0703a4c340741095a5b5b7875c1b3e9eb78ce, independently matching the feature
remote. Owner-merged PR #11 main 10c21366085ad83814bbfda85533cbeead93b571 had the identical tree.
Read-only fetch preserved the checkout and branch. Successor draft PR #12
contains M2-T03; the preserved non-Git snapshot remains separate.

Implementation synchronization is a past exact-SHA observation. This record-only
checkpoint binds that implementation; its own SHA/live remote state is reported
externally after normal commit/push. Raw logs, private media and local tools stay ignored.

The actual 324-character local source and complete longer output succeeded in both
desktop shells in this configuration. A disabled-LongPathsEnabled environment was
not configured or tested and remains not_run. Root-boundary helper cases are
lexical; no entire drive/share root was normalized.

Live UNC evidence covers only the preexisting localhost C$ alias of marked owned
fixtures: direct input/output, BAT source, and both local/UNC containment
directions. It does not establish arbitrary remote SMB servers, ACLs, network
providers or unsupported FileIdInfo implementations. Hosted live-UNC status is
recorded separately from its mandatory test count.

The BAT/application receiver has owned output/dependency/argv/application-return
seams only at the final PS1 invocation. Explorer drag-and-drop, default Pictures
routing and general final BAT exit propagation were not certified. Caller CMD
expansion before entry cannot be reconstructed; use the documented direct
PowerShell literal form for paired percent expressions.

Directory identity and reparse checks provide the documented conservative
preflight boundary, not a hostile concurrent-filesystem sandbox or universal
cross-provider equivalence proof. Source metadata stability and duplicate matching
remain length/time heuristics; same-key same-length differing content can still
skip under the declared heuristic.

Prior codec and colour limitations remain: HEIC/HEIF evidence is a two-image still
collection, not timed animation/thumbnails/auxiliaries; the ICC guard is bounded
structure/model validation, not complete semantic conformance, and the chosen CMYK
A2B0-only display profile uses the declared requested-Relative fallback. Synthetic
patch tolerances do not grant owner photographic/aesthetic approval.

M2-T04 and later byte-cap/quality, process-resource lifecycle, cancellation,
reporting and release work remain explicit future gates. No owner quality/default
approval, merge, tag, release or deployment is implied.

Stop after M2-T03; start M2-T04 only when next requested.
