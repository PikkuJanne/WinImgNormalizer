# M3-T07 owner review — awaiting_owner

Owner: Janne Vuorela. Prepared: 2026-10-05; owner review has not occurred.
Exact review commit: **4618a83cb2aa893aea9d6a1587987612624e1cf7**.
No application, launcher, test or default changes are proposed in this task.
The packet's Launcher files are exact copies from this clean checkout. Their hashes
and the verified automated preparation are in [M3-T07 evidence](evidence/M3-T07.json).

## What is ready

A disposable local packet outside the repository contains `Drop this folder`,
`Launcher`, `Owner comparisons.html`, `Comparison overview.png`, `Size warning
folder`, and `OWNER-READ-ME.txt`. The local packet location is supplied in this
session's response and ignored `.scratch/M3-T07-location.json`; private machine
paths and generated media are not committed or uploaded.

The source contains 24 synthetic files: ordinary landscape/portrait/detail/noise
and already-compressed images; genuine RGB/CMYK ICC profiles; transparency;
EXIF orientation 6; GIF/WebP first canvases, TIFF first page and HEIC/HEIF primary
still collections; literal special-character paths; same-stem conflicts; a
heuristic duplicate; opaque MP4 copy-test bytes; two ignored text files and an
empty directory. The MP4 is intentionally unplayable; its byte equality is checked.
Optional owner photos/video must be approved disposable copies kept outside the
repository. No private sample is required for this prepared review.

Fresh automated Command preparation on Windows 11 Pro build 26300, Windows
PowerShell 5.1.26100.9444 and PowerShell 7.6.5 with pinned ImageMagick 7.1.2-32:

| Run, each shell | Result |
|---|---|
| Default 1,048,576-byte target | Native 0; 24 discovered = 20 converted + 1 copied video + 1 duplicate + 2 ignored; zero errors/size warnings; five omitted frames/pages/images |
| Deliberate 1-byte target | Native 2; two valid retained JPEGs, 546 and 1,621 bytes at 50%; two size warnings |

All JPEGs independently fully decode in Pillow, with one frame, matching
dimensions and no retained EXIF/ICC/comment/XMP fields. Video bytes match.
Source file hashes, file/directory times/attributes and persistent execution policy
are unchanged; reports complete and accounting balances. These runs inject a
disposable destination and explicit executable, so they do not establish the
owner's Explorer gesture, actual Pictures destination or appearance approval.

## Run and inspect

1. Open `OWNER-READ-ME.txt` in the local packet. For the everyday Explorer test,
   ImageMagick must be available to the launcher's PATH. This agent's session has
   no `magick.exe` on PATH; the verified portable development copy exists. No
   installation or persistent environment/policy setting was changed.
2. Drag exactly `Drop this folder` onto `Launcher/WinImgNormalizer.bat` in Explorer.
   The unchanged BAT invokes Windows PowerShell, retains pause, and should say
   `Completed successfully`. Record any actual dependency/setup failure.
3. Find the new `Drop this folder_WinImgNormalized_<timestamp>_<runId>` directory in
   Windows' actual Pictures known folder (which may be redirected), with the
   USERPROFILE/Pictures fallback only if no known folder is available. Inspect
   the mirrored and empty directories and `.WinImgNormalizer/reports` log/CSV.
4. Compare the real smoke outputs with the prepared originals/current displays.
   Check orientation, colour, white transparency, fine detail and acceptable JPEG
   appearance. Source previews are native colour-managed display aids, not
   independent colour oracles. Click images for full-resolution inspection.
5. Confirm all three collision outputs exist: `photo__bmp__2.jpeg`,
   `photo__jpg.jpeg`, `photo__bmp.jpeg`. The duplicate row must link its retained
   source/output. Confirm frame/page omissions and the accurate summary are clear.
6. Confirm the source remains intact and the MP4 copy has equal bytes. Review the
   two deliberately above-target examples and their warning text in the prepared
   tiny-cap report. A custom-cap command is separate from the default BAT flow.

`Run owner smoke.cmd` is an optional local convenience that supplies the verified
portable tool via child-process PATH and invokes the unchanged BAT with the fixed
synthetic folder and actual Pictures resolution. It does not prove a physical
folder-to-BAT gesture. Record which method you actually used; never report one
as the other. The helper does not install tools or change persistent settings.

## Appearance comparison and defaults

The five historical baseline outputs are hash-bound accepted M2-T04 artifacts,
baseline commit f9b00d8befac93c4fa1116efcbc2e9d7144f4509. Current columns are fresh
outputs from the review commit above. These are descriptive synthetic comparisons;
they are not new baseline runs or approval supplied by an automated metric.

| Sample | Historical baseline bytes | Current bytes |
|---|---:|---:|
| Smooth landscape | 278,605 | 278,605 |
| Smooth portrait | 264,584 | 264,584 |
| Hard edges/detail | 903,155 | 903,155 |
| Seeded noise | 979,962 | 1,015,369 |
| Already-lossy JPEG | 840,679 | 916,662 |

The first three are byte-identical on this build. The latter two differ following
M2-T04's exact-byte extent correction. The 1,048,576-byte default, six scales,
4:2:0 sampling, orientation, white background, sRGB handling and frame policy are
retained; unchanged settings do not promise identical historical output quality.
No new quality/resize default is proposed, and no later tuning is approved here.

## Owner response

Outcome: **awaiting_owner**. T067: **not_run**. Owner statement: **not supplied**.

After the actual smoke, reply with:

- Commit `4618a83cb2aa893aea9d6a1587987612624e1cf7`; Windows/tool versions and date.
- Actual launch method, completion/pause message and Pictures result; source/video
  preservation; collisions/duplicates/frame/size reports and appearance findings.
- `I accept the workflow and image appearance at this commit`, or the concrete
  requested changes. State any checks you did not perform rather than assuming them.

The agent cannot sign this response for the owner. Live network shares are outside
owner-approved project scope (D35). Merge, release, website deployment and later
default changes retain separate authorization; M4-T01 waits for accepted T067.
