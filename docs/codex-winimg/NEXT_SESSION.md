# Next session handoff

Next task: **M2-T02 — Convert colour profiles before metadata removal**.
M2-T01 is complete. Begin only M2-T02 when next requested.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; WinImgNormalizer-main
is the preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json,
GIT_WORKFLOW.md, tasks/M2-T02.md and its referenced implementation/fixture/source
requirements for T032-T036. Recheck clean state, canonical fetch/push URLs and exact
advertised feature SHA before editing.
Tested runtime/tests implementation: d64a9034ff0b4324b2515900fdf8b4cb82f80bbb.
Final evidence-checkpoint SHA/live sync is reported externally in the preceding
response and draft PR #10; verify it independently again at session start.

GIF/TIF/TIFF/WebP/HEIC/HEIF inspection and conversion read the same exclusively
created owned snapshot with a neutral source basename and original extension.
Source regular-file/reparse checks and length/modification-time checks surround
the copy, and copied length must match. External snapshot arrivals and unknown
neighboring files are preserved; exact owned snapshots follow existing cleanup and
partial-warning rules.

A separate successful identify -ping probe counts every decoder-exposed image and
rejects malformed, inconsistent or ambiguous count/format/dimension observations.
Conversion selects image:frames=0 before the owned native input. Only actual
decoded GIF/WebP follows FirstDisplayedFrame coalescing onto its logical canvas;
+repage and existing auto-orientation follow. TIFF keeps one first page without
stacking; HEIC/HEIF keeps the decoder's primary/first image. Each finalized output
still must fully decode as one nonempty JPEG at the exact planned path.

SOURCE IMG records source count, selected count 1, omitted count, unit, policy and
actual decoder. Deliberate omission is the normal informational static-output
policy and retains exit 0 on successful runs. Source bytes/times, mirrored naming,
no-overwrite finalization, verified videos, success-based heuristic duplicate
links, size cap/scales/JPEG flags and public BAT/positional forms retain their
established behavior. The bootstrap, dependency pins and workflow were not
changed.

Decision D24 records the scoped frame/page policy. Frame/page counts and deliberate omissions remain visible.
Actual codec/sequence observations are in M2-T01.json; do not infer arbitrary HEIC
sequence support from a codec name or static decode. Covered M1 source preservation,
two-run isolation, deterministic collision plan, validated transactional finalization
and exact owned cleanup remain mandatory combined regressions.

M2-T02 must inspect embedded ICC profiles before stripping, convert known tagged
colour into sRGB through a verified profile-aware path, and only then apply the
privacy metadata policy. Define tagged/untagged RGB and CMYK behavior, alpha order,
malformed profiles and the output profile policy; preserve auto-orientation and the
declared first-frame/page selection. Use licensed, provenance-recorded reference
profiles/fixtures and tolerant colour-managed comparisons. Remove GPS/EXIF/XMP/IPTC
and unwanted profiles under the declared policy; keep uncertain cases visible.
Do not begin later size/native argv/lifecycle/cancellation/report/publication tasks.

Desktop PS 5.1/PS 7 passed 185/185, zero skipped. Controls each 186 total,
185 passed, one T007 false assertion, exit 1. Push/PR Windows Server matrices
passed at the implementation SHA. Exact raw/Git bindings, versions, codec evidence,
saved artifacts, histories and internal review are in the two M2-T01 evidence files.

Commands, separately in each fresh Windows shell:

```powershell
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
.\tests\Invoke-Tests.ps1 -DeliberateFailure # expected exit 1, one false assertion
```

Pinned development bootstrap reuses verified archives or explicitly downloads only
to ignored scratch; no installer/runtime bootstrap/persistent PATH or policy change.
Children clear inherited PSModulePath; never redefine USERPROFILE. Keep media,
tools and raw artifacts ignored, and retain real failures rather than weakening tests.

Actual T031 HEIC/HEIF evidence is a genuine two-image top-level still collection
through two aliases. Timed HEIC animation, thumbnails, auxiliary images and
arbitrary codec builds remain unverified; container counts describe only images
exposed by the pinned decoder. The pinned native build reads HEIC/HEIF but cannot
generate the fixture; one verified official development wheel produced the
synthetic bytes in ignored scratch, with no installation, runtime/CI generator
dependency or distributed generator binary.

Snapshot length/modification-time checks and duplicate keys are stability
heuristics, not proof of content identity. General native argument/input-grammar
safety and source changes retaining identical metadata remain outside this narrow
neutral-snapshot policy. Owned cleanup deliberately preserves unknown entries and
may return partial exit 2 after a valid JPEG has finalized.

Tagged colour/profile conversion before metadata removal is next M2-T02. Broader
alpha/colour/reference fidelity, size-search quality, process timeout/resource
budgets, cancellation, reporting and publication remain later tasks. Automated
internal peer review does not grant owner aesthetic acceptance or authorize merge,
release, deployment, tags or default-branch changes.

Keep the feature branch and draft PR #10; if the owner merged it, inspect exact
remote/tree identity and create one successor draft without rewriting history.
Normal scoped feature commits/pushes remain authorized; merge/default-branch changes,
tags/releases/settings/deployment and owner quality approval are separate gates.
Stop after M2-T01; begin M2-T02 only when next requested.
