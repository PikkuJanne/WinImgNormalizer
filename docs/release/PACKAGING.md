# Portable package preparation

`tools/release/Build-Release.ps1` prepares an **unreleased, unsigned** local ZIP from
an explicitly named clean current commit. It creates no tag or public release,
publishes no downloads, and does not run the packaged application automatically.
Publication and website deployment require the separate owner-approved gates.

## Build a reviewed revision

Use Windows PowerShell 5.1 or PowerShell 7, Git and the repository's ordinary local
filesystem. Review and commit the intended source first. The builder requires the
full 40-character current `HEAD` revision, rejects tracked/staged/untracked changes,
and checks that the selected files are ordinary Git blobs. Ignored scratch outputs
are excluded. It does not repair a dirty checkout or update source files.

Run from the repository root, choosing an absent output directory below the ignored
`.scratch` directory:

```powershell
$releaseRevision = (git rev-parse HEAD).Trim()
.\tools\release\Build-Release.ps1 -Revision $releaseRevision -OutputDirectory '.scratch\release candidate with spaces'
```

Use a new output directory for each build. The builder refuses an existing target
and linked output paths or ancestors. Build artifacts stay in the explicitly owned
scratch directory; no existing destination file is replaced. Git blobs supply the
package bytes, including the launcher's required CRLF bytes, independently of a
checkout's text-conversion settings.

## Exact content boundary

The portable ZIP contains these seven tracked inputs, mapped to root filenames,
plus the generated `package-manifest.json`:

| Tracked source | ZIP filename |
| --- | --- |
| `WinImgNormalizer.ps1` | `WinImgNormalizer.ps1` |
| `WinImgNormalizer.bat` | `WinImgNormalizer.bat` |
| `LICENSE` | `LICENSE` |
| `CHANGELOG.md` | `CHANGELOG.md` |
| `docs/release/GETTING_STARTED.md` | `GETTING_STARTED.md` |
| `docs/release/THIRD_PARTY_NOTICES.md` | `THIRD_PARTY_NOTICES.md` |
| `tests/fixtures/colour/LICENSE-CC0.txt` | `LICENSE-CC0.txt` |

Only the retained third-party license text is selected from the fixture tree; tests,
ICC fixture files, synthetic media and reference artifacts are excluded. The ZIP
also excludes Git internals, governance/handoff records, development tools, raw
media, logs, secrets, branding assets and earlier builds. Runtime entry points stay
self-contained. No ImageMagick binary or installer is bundled or downloaded.

The manifest records the application version from `Get-WinImgVersion`, exact source
commit, each selected source/ZIP mapping, byte lengths and SHA-256 values. The
output directory contains `WinImgNormalizer-<version>-portable.zip`, outer
`build-provenance.json` recording its identity/checksum, and `SHA256SUMS.txt` hashing
the ZIP and provenance file. The checksum file does not hash itself. Public
release/download fields stay unestablished. These records identify source and bytes;
they provide no signature, publisher verification or attestation.

## Reproducibility and verification

The builder sorts entries, uses uncompressed ZIP entries and normalizes entry
timestamps to a fixed ZIP-compatible date. It omits observation time, local paths
and host identity from package bytes. Repeated builds of the same revision with the
same builder/runtime are intended to produce byte-identical ZIPs. Tests must compare
the actual archive hashes before recording that observation. This is not a claim
that every Git, .NET or ZIP implementation on every platform produces identical
bytes. Independently verify content equality through the manifest and extracted
file hashes when comparing different environments.

Packaging regressions cover T069 (exact allowlist and source contents), T070
(fresh extraction and matching entry-point smoke checks with documented
prerequisites), and T071 (recomputed checksums and exact revision provenance).
Run them in both supported PowerShell hosts. Use new extraction directories with
spaces and approved synthetic source data. Real extracted-script image/video checks
use the existing `OutputParent` test seam to keep output in owned scratch, separate
from real Pictures. Exact extracted BAT setup/pause checks exercise the matching
PS1; they do not establish a completed media batch through the BAT. Check source
preservation, real native results, derivatives/video bytes and ordinary-account
observations. Keep raw artifacts and private paths local; persist sanitized evidence
with actual results and limitations. A launcher test using a substitute PS1 driver
is not evidence that the extracted application processed media.

An artifact built before the implementation commit cannot claim that future SHA.
After an evidence-only checkpoint, the package can remain bound to the preceding
implementation commit when all package/build inputs retain their recorded bytes.
For eventual release, rebuild or verify against the actual approved tag revision;
never label an earlier artifact with an unverified newer revision. No application
smoke test, signing check or owner acceptance is implied by a successful build.
