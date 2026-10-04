# Baseline and evidence boundaries

Observed on 4 October 2026: `main` at
`4b918da50639d8fc7e3fddc6d6b3678880ea4098`. Blob IDs are in BASELINE.json.
The source, launcher and repository tree were re-read through GitHub for this
bundle. The initial review also read README, LICENSE and the then-empty release
collection. Re-read mutable GitHub metadata before relying on it. [R1–R6]

No local user checkout was accessible. No end-to-end Windows runtime result was
obtained for the application in this bundle. The earlier review asserted some
Linux command-level reproductions, but their commands, versions and artifacts
were not available to carry forward. Treat those as cases to reproduce, not as
passing evidence. The new bundle-helper tests are separate from application tests.

## Static findings and anchors

| Finding | Source anchor | What inspection establishes |
|---|---|---|
| Output collision | `ChangeExtension($rel, '.jpeg')` | Different same-stem input extensions map to one final name; no collision plan is present. |
| Nested output traversal | `$destRoot` creation before recursive `Get-ChildItem` | No explicit exclusion protects a source that contains the destination. Exact recursion behavior must be reproduced safely. |
| Weak output success | `Test-Path $DestPath` then `$len -le $MaxBytes` | Existence and length can be accepted without a final successful exit and decode validation. |
| Frame mismatch | Raw `$SourcePath` passed to JPEG output | No deliberate first-frame/page selection appears, despite the header claim. |
| Profile order | `-auto-orient -strip -colorspace sRGB` | Source profiles are stripped before the later operation. Quantify actual errors with proper reference fixtures. |
| Failed-first duplicate | `$seen.Add($dupKey)` before processing | A failed first candidate still reserves a key and can suppress a later candidate. |
| Special filename parsing | Source/output strings passed to ImageMagick | PowerShell literal filesystem handling does not settle native filename interpretation. |
| Root-path semantics | `.TrimEnd('\','/')` | A drive root loses its trailing separator; repair with a root-aware canonicalizer. |
| Weak preflight | Direct Int64 cast and late `magick -version` | Friendly bounds, exit-code/version and per-format capability validation are missing. |
| Ambiguous sizing | `Get-ExtentString`, 100..50 scales and fallback | The byte cap is best-effort; fallback results count as converted/OK. Verify ImageMagick units independently. |
| Lost diagnostics | Both native output streams discarded | Failure detail is unavailable and fallback removes alpha flags regardless of cause. |
| Enumeration failure | Recursive file enumeration outside per-file try | One traversal failure can stop the discovery stage; no folder-level recovery policy. |
| Hidden reporting failures | Empty catch blocks | Timestamp/log failures can vanish. |
| Unsupported counter | Extension filter before unsupported branch | Unsupported files cannot reach the later unsupported counter. |
| Launcher result | Unconditional Done and no exit-code propagation | Reported completion is not tied to application outcomes. |

The tree has seven files and no test/workflow directories at the observed commit.
The license is MIT. Do not treat the absence of tests as proof of every suspected
failure mode; reproduce it. Do not treat a future different commit as invalid. [R4,R5]

## Reproduction protocol

Use synthetic directories outside the real archive, bounded execution, and captured
versions. Do not run the current nested-output case against Pictures, a profile or
a drive root. First characterize path planning with fake paths; use a disposable
Windows account/VM or a narrowly injected test seam for end-to-end traversal.
Keep expected legacy failures separated from regression tests for the corrected
behavior. A failing legacy reproduction is useful evidence, not a reason to make
ordinary release CI permanently red.
