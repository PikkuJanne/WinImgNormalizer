# Distribution and website handoff

## Deliverable boundary

Keep processing entirely in the local PowerShell application. The website is a
product page, documentation and download channel. Do not create an upload endpoint,
conversion API, server runtime, login requirement, queue, database or telemetry.
Do not choose a site framework/domain/provider without actual website context.

## Version and packaging

Inspect current tags/releases before choosing a first managed semantic version.
Use one version source and include it in help/logs, release notes and package
metadata. Do not label the current baseline v1.0 merely because the bundle is v1.0:
**bundle version and application version are separate**. Preserve the existing MIT
license and author attribution. Any added ICC profile or other redistributed asset
needs its own provenance and license. Do not bundle ImageMagick by default. [R5]

Build from the reviewed clean commit with an explicit allowlist. Suggested release
contents: `WinImgNormalizer.ps1`, matching `.bat`, any genuinely needed small support
file/profile with its license, `README.md` or compact `GETTING_STARTED.md`, `LICENSE`
and release notes. Exclude `.git`, task/governance records, helper Python tools,
tests, build outputs from earlier runs, logs, fixtures, personal media and secrets.
A release builder must not archive the entire developer directory indiscriminately.

Use stable entry ordering and normalized metadata where practical. State whether
reproducibility means identical content or byte-identical ZIPs and test the actual
claim. Produce a SHA-256 for each downloadable asset and a checksum file; test a
clean extraction from that exact ZIP on Windows in a directory with spaces. Document
ImageMagick installation, executable verification and the tested build/capability
range. The application must not download executables during a normalization run.

Do not instruct users to run everything as administrator, disable Defender or
change machine-wide execution policy. Explain source review and, when appropriate,
selective unblocking only after verifying trusted downloaded files. A checksum is
not a digital signature. Signing can be a later separately credentialed enhancement;
never show a verified-publisher claim before actual signing/verification. [P1,G3]

## Website content package

Prepare framework-neutral product copy and metadata under a small `docs/website/`
area or the website's actual existing content structure when supplied. Reuse the
existing PNG/ICO/poster rather than recreating branding without reason. Screenshots
must be genuine captures of the current tested app. Before/after samples must come
from approved synthetic/public fixtures with real dimensions/bytes; do not generate
fake output screens or invented benchmark percentages. [R4]

Recommended metadata fields: product name, short description, application version,
release date, platform/tested environments, ImageMagick prerequisites, license,
repository URL, release notes URL, asset filename, asset size, SHA-256 and download
URL. Fields without real published values remain null/draft and publication checks
must refuse to expose them as live downloads. No invented domain or current release.

The page should explain: drag one folder onto the launcher; find the normalized
copy under Pictures; originals remain unchanged; JPEG size is best-effort at an
explicit byte target; output is a lossy derivative, not a replacement backup; images
lose private metadata after colour handling; videos are copied and keep metadata;
animations/pages beyond the chosen image are omitted and reported; duplicate
skipping is heuristic; codecs vary by ImageMagick build; source-containing-output
paths are rejected by the first hardened release.

Privacy wording: “The tool processes images on your computer and does not upload
them.” Add the limitation that output in a folder managed by cloud-sync software can
be synchronized by that software. Do not claim the application controls OneDrive
or another sync provider. Clarify that support logs can contain private paths.

## Publication gates

M4-T01 through M4-T05 prepare artifacts, not public releases. M4-T06 requires explicit
owner approval naming the actual PR/commit, proposed tag/version, artifact hash,
release visibility and permitted actions. Approval to prepare a bundle does not
satisfy this gate. Website deployment requires separate approval and the actual
website repository/provider context. No hosting settings or credentials are in scope.

Before publishing, verify full required test coverage, CI result, owner acceptance,
license content, release ZIP contents and checksums against the actual tag commit.
After publishing, verify the real download URL and bytes, then update page metadata.
Retain an earlier release instead of force-updating an existing tag or silently
replacing a released asset. Document any corrected replacement with a new version.
