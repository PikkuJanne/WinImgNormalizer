This is the first managed application version, prepared from the existing
unversioned tool. It is **unreleased**: no tag, public release or downloadable
package is established by these notes. The application version is generated from
`Get-WinImgVersion` in `WinImgNormalizer.ps1`; it is separate from development
tooling or task-bundle versions.

## Fixes

- Check arguments, the source/output boundary and the selected ImageMagick
  executable before creating output. Reject linked roots/ancestors and
  source-containing-output paths; skip and report links during traversal.
- Treat supported filename characters literally through PowerShell and
  ImageMagick. Plan names before conversion so same-stem images and generated
  files cannot collide with retained media or mirrored directories.
- Validate a fresh candidate's native result, nonempty bytes, single JPEG frame,
  dimensions and full pixel decode before a no-replacement final move. Failed or
  stale candidates cannot become completed outputs.
- Register heuristic duplicates only after successful finalization, using stable
  source-path order and equal byte lengths. Record the retained source/output and
  warning status so failed earlier files cannot suppress later usable media.
- Pass the requested JPEG extent as exact decimal bytes and compare the actual
  file length with the cap. Retain a valid final above-target JPEG with a visible
  warning instead of reporting a successful size match.
- Select one deliberate animation frame, TIFF page or HEIC/HEIF primary image and
  report decoder-exposed omissions. Preserve source ICC characterization until
  conversion to sRGB; clamp converted sRGB before white-alpha composition on HDRI
  builds, then remove metadata. Reject uncharacterized CMYK and invalid colour
  transforms instead of guessing.
- Drain bounded native diagnostics independently, stop failed items for native
  errors and enforce a shared per-image deadline. Retry only diagnosed temporary
  sharing/lock failures; smaller scales address size, not damaged input or missing
  decoders.
- Continue accessible siblings after scan/item failures and expose incomplete
  scans, log/CSV failures and timestamp-restoration warnings. Reconcile known-file
  outcomes and write CSV text with a visible formula-safety prefix.
- Preserve the launcher's pause and propagate the application exit code. Reject
  extra arguments and preserve literal exclamation marks and root separators.

## Default behavior and safeguards

- Processing stays on the computer, with no uploads or automatic dependency
  installation. Originals remain unchanged; each exclusive run mirrors the
  source's subdirectories under Pictures and never replaces an existing target.
- The familiar folder-to-BAT workflow and both positional PowerShell invocation
  forms remain. Images become lossy JPEG derivatives; videos keep their bytes and
  metadata. The default remains 1,048,576 bytes (1 MiB), the six scales remain
  100/90/80/70/60/50 percent, and no upscaling is introduced.
- Images are oriented, converted to sRGB, composited onto white and stripped of
  private metadata. The target sRGB profile is embedded in the script; no separate
  ICC installation is needed. JPEG sampling/progressive settings remain unchanged.
- Native work runs in an owned Windows process tree with process-local cache,
  thread and time limits. Installed ImageMagick security policy remains effective;
  machine-wide execution policy and other security settings are unchanged.
- Cooperative cancellation keeps completed outputs, stops owned native work and
  removes known incomplete staging files where safe. Application outcomes are
  success (0), setup/run failure (1), warning/partial completion (2) and handled
  cancellation (130).

## Requirements and supported environments

- A 64-bit Windows 10 or Windows 11 host with Windows PowerShell 5.1 or PowerShell
  7. Keep the PS1 and BAT together for the launcher workflow.
- ImageMagick **7.1.2-32 or newer supported 7.x**, with `magick.exe` discoverable
  through PATH. ImageMagick is separately installed and is not bundled by default.
  Version and compiled codec availability are checked at startup. JPEG writing
  and each source format's decoder must be available; installed policy may still
  restrict a format. ImageMagick remains required for video-only and empty runs.
- Development test/analyzer tools are not application runtime dependencies. The
  application does not install or update ImageMagick during a run.

## Actually tested environments

- Desktop: 64-bit Windows 11 Pro build **26300**, Windows PowerShell
  **5.1.26100.9444** and PowerShell **7.6.5**. Maintained full gates used portable
  ImageMagick **7.1.2-32 Q16 x64**; installed **7.1.2-32 Q16-HDRI x64** received
  separate colour regression and owner workflow/appearance checks.
- Hosted CI: 64-bit Windows Server 2025 Datacenter **10.0.26100**, Windows
  PowerShell **5.1.26100.33438** and PowerShell **7.6.6**, with the maintained
  portable ImageMagick **7.1.2-32 Q16 x64** build.
- These are recorded test environments, not coverage of every supported host or
  ImageMagick build. Windows 10 has not been tested. Hosted automation and desktop
  regression tests remain separate from the recorded owner's manual acceptance.

## Limitations

- JPEG output is lossy and is not a replacement backup. Byte targets are
  best-effort; a valid 50-percent candidate may still exceed the cap and returns a
  warning. Duplicate detection is a filename/time/length heuristic and can skip
  different content that shares those values.
- Animation and pages beyond the selected image are omitted. Codec behavior and
  local ImageMagick policy vary by build; compiled capability is not proof that
  every file can be decoded.
- Live UNC/network-share support and validation are outside the owner-approved
  scope. Lexical UNC-root checks and local path/identity tests do not establish
  network-share support. Long local paths depend on the host, filesystem and tools.
- Output inside a cloud-synchronized folder can be uploaded by that sync software.
  The application does not control it. Support logs and reports can contain private
  paths; review them before sharing.
- Resource limits are practical cache/process controls, not a total decoder heap
  cap or exploit-proof sandbox. Blocked filesystem operations can delay stopping;
  concurrent hostile filesystem mutation is outside the isolation guarantee.
  Terminal closure or forced termination may leave recognizable partials and may
  bypass cooperative cleanup, reporting and exit 130.
- Checksums identify bytes; they do not establish a digital signature or verified
  publisher. Packaging, publication and website deployment remain separate gates.

## Authorship and license

WinImgNormalizer remains authored by **Janne Vuorela** and licensed under the
**MIT License**, with the original `Copyright (c) 2025 Janne Vuorela` notice retained
in `LICENSE`. The embedded target ICC profile is CC0-1.0 Compact-ICC-Profiles
sRGB-v4; immutable provenance is recorded beside its bytes in
`WinImgNormalizer.ps1` and in `tests/fixtures/colour/README.md`, with the retained
third-party license in `tests/fixtures/colour/LICENSE-CC0.txt`.
