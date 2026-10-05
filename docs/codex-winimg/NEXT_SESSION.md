# Next session handoff

Next task: **M2-T03 — Make filesystem and native filename handling literal-safe**.
M2-T02 is complete. Begin only M2-T03 when next requested.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; WinImgNormalizer-main
is the preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json,
GIT_WORKFLOW.md, tasks/M2-T03.md, IMPLEMENTATION_SPEC.md and TEST_MATRIX.md,
including T037-T039. Recheck clean state, canonical fetch/push URLs and exact
advertised feature SHA before editing.
Tested runtime/tests implementation: a30c456594a31368294588cc198444b38bc3e0e3.
Final evidence-checkpoint SHA/live sync is reported externally in the preceding
response and draft PR #11; verify it independently again at session start.

Every eligible image is copied once into a stable owned neutral snapshot;
frame/page inspection, source colour inspection and all conversion attempts use
those same preserved source bytes. Strict native inspection clears conflicting
free-form profile/colorspace properties without stripping actual ICC
characterization.

Tagged sources retain their source ICC through owned exact extraction and bounded
header/tag/model checks, then transform to a hash-verified embedded CC0 sRGB-v4
target. Requested Relative intent explicitly disables black-point compensation.
Decoded untagged sRGB is assumed sRGB; native linear RGB and gray are converted
explicitly. Untagged CMYK without source characterization is rejected with a
source-preserving partial error.

Auto-orientation precedes colour conversion. RGB transforms precede compositing
over white in encoded sRGB; alpha is then disabled and
ICC/EXIF/GPS/XMP/IPTC/comments are removed. Every 100, 90, 80, 70, 60 and 50
percent attempt uses that same order and white-alpha policy. The inherited
alpha-dropping 100-percent retry is removed so a retry cannot change the
background policy.

Malformed/mismatched profiles, ambiguous inspection and nonempty native diagnostic
output cannot claim accurate successful conversion. Existing transactional full
single-JPEG validation and no-overwrite finalization still decide whether output
is retained; later valid files continue after a source error, and exact owned
cleanup preserves unrelated entries.

The established positional/BAT interface, 1,048,576-byte default, scale sequence
and best-effort boundary, first-frame/page selection with informational omissions,
mirrored naming, source preservation, copied videos and success-only duplicate
links remain the regression baseline. No runtime dependency or dependency pin
changes; the runtime target profile is embedded rather than downloaded.

Three CC0 source/target profiles and an optional reproducible independent Pillow
12.3.0/LittleCMS 2.19 recipe provide synthetic references. Mandatory tests use
checked reference data, with distinct 3/255 pre-JPEG and 12/255 JPEG interior
patch tolerances, and require the actual T032-T036
RGB/CMYK/malformed-ICC/orientation/metadata/alpha cases.

Decision D25 records the tested colour/profile and metadata policy. Keep source/profile/metadata preservation and privacy policy honest.
M1 transactional safety and M2-T01 frame/page selection remain mandatory combined
regressions; actual HEIC still-collection evidence does not establish timed animation.

M2-T03 must verify both PowerShell argv and ImageMagick input/output grammar at real
Windows paths. Exercise percent/brackets, quoted spaces, ampersands, apostrophes,
Unicode and supported long/UNC/root boundaries through direct calls and the batch
launcher. Use neutral owned staging or supported literal behavior where needed;
never append unexamined selectors, strip literal identity, truncate paths or
introduce shell evaluation. Document unsupported path environments clearly without
source mutation. Preserve the colour transform, alpha order, metadata policy,
orientation and explicit first-frame/page selection.
Do not begin later size quality, process lifecycle, cancellation, reporting or release work.

Desktop PS 5.1/PS 7 passed 210/210, zero skipped. Controls each 211 total,
210 passed, one T007 false assertion, exit 1. Push/PR Windows Server matrices
passed at the implementation SHA. Exact raw/Git bindings, actual environments,
reference/provenance/tolerance evidence, artifacts and internal review are in the
two M2-T02 evidence files.

Commands, separately in each fresh Windows shell:

```powershell
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
.\tests\Invoke-Tests.ps1 -DeliberateFailure # expected exit 1, one false assertion
```

Pinned development bootstrap reuses verified archives or explicitly downloads only
to ignored scratch; no installer/runtime bootstrap/persistent PATH or policy change.
Children clear inherited PSModulePath; never redefine USERPROFILE. Keep private
media, tools and raw artifacts ignored, and preserve genuine failures for diagnosis.

The selected CGATS001Compat CMYK profile contains only an A2B0 perceptual forward
display mapping. Requested Relative intent uses LittleCMS fallback to that
available mapping; evidence covers this source fixture rather than general
relative-colorimetric printing accuracy, all rendering intents or reverse CMYK
conversion.

The ICC guard checks bounded structure and decoded model consistency, not complete
ICC semantic conformance. Synthetic patch and metadata fixtures do not prove
arbitrary ICC profiles, every codec, production photography or subjective owner
quality.

Length/modification checks remain a source stability heuristic rather than content
identity. Duplicate matching remains the documented same-name/time/length
heuristic; mandatory hashing, content databases and original removal remain
excluded.

Actual HEIC/HEIF coverage retains the two-image still collection from M2-T01.
Timed HEIC animation, thumbnails and auxiliary-image behavior remain unverified.

General literal-safe filenames, supported long/UNC/root boundaries,
process/resource/cancellation controls, later quality/reporting/release work and
owner aesthetic/default acceptance remain later explicit gates. No merge, tag,
release, website deployment or owner quality approval is implied.

Keep the feature branch and draft PR #11; if the owner merged it, inspect exact
remote/tree identity and create one successor draft without rewriting history.
Normal scoped feature commits/pushes remain authorized; merge/default-branch changes,
tags/releases/settings/deployment and owner quality approval remain separate gates.
Stop after M2-T02; begin M2-T03 only when next requested.
