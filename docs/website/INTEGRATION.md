# Website integration handoff

This is preparation for a future product page and documentation/download channel.
Use [the product copy](PRODUCT_COPY.md) and [website metadata](metadata.json).
The actual website repository, domain, framework, hosting provider and deployment
workflow have not been supplied. Integration and deployment require their own
approved context.

## Boundary

Processing stays in the local PowerShell application. This handoff introduces no
upload form or endpoint, conversion API, server-side processing, account requirement,
queue, database or telemetry. The website should explain the tool, link to its
documentation and expose verified release assets when publication is approved.

The current content is a draft. It does not establish a public release, working
download button, website UI, screenshot or deployed service.

## Canonical metadata

[release-metadata.json](../../release-metadata.json) is the authority for version
and release/download fields. Its version derives from `Get-WinImgVersion` in
`WinImgNormalizer.ps1`. The website metadata's `release_metadata_file` points to
`../../release-metadata.json`; it carries no second version/tag/date/asset/download
projection. Keep the existing metadata derivation check before integration.

| Page value | Canonical field or source |
|---|---|
| Product name | `release-metadata.json.product` |
| Application version | `release-metadata.json.version`; use the derived value at rendering time. |
| Release state | `release-metadata.json.release_state` |
| Published tag and release date | `release-metadata.json.tag`, `release_date` |
| Release page | `release-metadata.json.release_url` |
| Asset filename and bytes | `release-metadata.json.asset_filename`, `asset_bytes` |
| Expected SHA-256 | `release-metadata.json.asset_sha256` |
| Download link | `release-metadata.json.download_url` |
| Repository / View source | `release-metadata.json.repository_url` |
| Release notes | Repository file named by `release-metadata.json.release_notes_file` |
| Requirements / quick start | Repository file named by `release-metadata.json.requirements_file` |
| License and author | `release-metadata.json.license`, `author`; retain [the full MIT notice](../../LICENSE). |
| Description, requirements and limitations | [Product copy](PRODUCT_COPY.md), checked against [README](../../README.md) and [behavior](../BEHAVIOR.md). |

The current canonical state is `unreleased`, with published tag/date/release URL,
asset filename/bytes/SHA-256 and download URL all null. `proposed_tag` is a proposal,
not a published tag. No placeholder, constructed GitHub asset URL, repository
source archive or older preparation ZIP substitutes for a verified release asset.

Display **Release download not available** as plain availability text, with
`render_link: false`. Omit the download anchor and any download-related structured
data. Source and documentation links remain separate navigation. They must not be
labeled as a published package download.

## Original artwork

Reference the original repository files below; do not copy or recreate them during
this preparation. Byte counts and hashes bind the inspected files. Dimensions
describe artwork, not processed-media results.

| Asset | Intended use | Original dimensions | Bytes | SHA-256 |
|---|---|---|---:|---|
| [PNG icon](../../WinImgNormalizer_icon_variant.png) | Product branding illustration | 1024 × 1024; 8-bit RGBA PNG | 11141 | `479f819325df22004e38f42183ff1d1e2351f753abed935a59357fed7da750f9` |
| [ICO icon](../../WinImgNormalizer_variant.ico) | Favicon source artwork | Six sizes: 16, 32, 48, 64, 128 and 256 pixels square | 22917 | `78ef6c01d0b466e72e3bad4fe83bb12b57a74c2f245a1274a559857d35374e79` |
| [Original poster](../../WinImgNormalizer_poster.png) | Historical reference only; `display: false` | 1536 × 1024; 8-bit RGB PNG | 2397241 | `99b1cd73450e725a5cce26a00ecdb6e660166ff8b5b9f75896882cfe7825d926` |

The icon depicts a stylized photo frame, overlapping frames and five green bars.
Suggested alternative text: **WinImgNormalizer product icon with a stylized photo
frame.** Use an empty alternative text when it is decorative beside the written
product name. The ICO is branding artwork, with no implication of a new GUI.

The poster is original promotional artwork with a drawn folder/terminal diagram.
It contains legacy `JPEG ≤ 1MB` text, an earlier output-name illustration and
`100 items processed, 37% done`. These are illustrative artwork details. Hold it
from the current hero, screenshot gallery, workflow instructions and result
comparisons: current behavior uses a best-effort 1 MiB target and current run names.
Keep `legacy_poster` marked `historical_reference_only` with display disabled.

Retain repository attribution and license notices. No new asset author, separate
license, benchmark or measured saving is inferred from the artwork.

## Screenshots and comparison material

The website metadata's screenshots and comparisons collections are empty. Existing
owner acceptance records have not been approved for public website reuse. Their
media, paths and raw reports are not website assets.

Before adding future material, obtain approval for the specific genuine capture
and its public use. Capture the named current tested application using approved
synthetic or public fixtures, record the exact source/output dimensions and bytes,
and retain reproducible provenance. Review images and console/report material for
private paths and identifying details. Label artwork and actual captures accurately.
Any before/after or saving statement must use actual measured outputs, with its
fixture, settings and limitations; do not invent percentages or application UI.

## Preparation validation

Run [Test-WebsiteHandoff.ps1](../../tools/website/Test-WebsiteHandoff.ps1) in the
repository's supported PowerShell hosts, together with the release-metadata drift
check. From the repository root:

```powershell
powershell.exe -NoProfile -File .\tools\release\Update-ReleaseMetadata.ps1 -Check
powershell.exe -NoProfile -File .\tools\website\Test-WebsiteHandoff.ps1
pwsh.exe -NoProfile -File .\tools\release\Update-ReleaseMetadata.ps1 -Check
pwsh.exe -NoProfile -File .\tools\website\Test-WebsiteHandoff.ps1
```

The website validator defaults `-RepositoryRoot` to the repository above its own
script. Its optional `-RepositoryRoot` and `-MetadataPath` arguments support owned
local test fixtures and alternate metadata under the selected repository. They do
not select a website host or enable publication.

The local validator is a read-only preparation gate. It validates the
metadata contract, local copy/document/asset references and unchanged original
artwork; it is not a deployment or network-download verification command.

Its current draft guard deliberately refuses any published state or live download
URL, even if a later URL appears plausible. Passing this guard establishes draft
readiness only. External document links are limited to the exact canonical
repository URL and the official ImageMagick Windows download link used in the
product copy. The draft validator accepts inline Markdown links; reference
definitions and HTML/autolink forms are rejected, and bare HTTP(S) destinations
must be on that same allowlist. New external URLs and GitHub release paths,
including latest-release download shortcuts, are rejected. A real published
projection requires a reviewed guard update and actual publication evidence in
the future authorized context.

## Later publication and deployment

M4-T01 through M4-T05 prepare artifacts. M4-T06 requires explicit owner approval
naming the actual PR/commit, proposed tag/version, asset hash, release visibility
and permitted publication actions. Preparation approval does not authorize that
publication. Follow [the release/website specification](../codex-winimg/RELEASE_AND_WEBSITE.md).

When the approved release actually exists, verify its tag commit, canonical version,
notes, exact package allowlist/licenses, clean Windows extraction, required tests
and CI. Read the real GitHub Release asset URL, fetch the actual approved asset in
the authorized release workflow and verify its byte count and SHA-256. Update the
canonical release metadata only from those observations, then review the website
guard and page projection. Display no verified-publisher or signature badge without
actual signing and verification. Checksums alone do not establish publisher identity.

Retain previous released versions; correct released content through a new approved
version instead of moving a tag or silently replacing an asset. After page
integration, validate its real links, rendered labels, accessibility and download
bytes using the actual approved website workflow. Website deployment remains a
separate owner gate requiring the real repository/provider context. No hosting
settings or credentials are requested or changed by this handoff.
