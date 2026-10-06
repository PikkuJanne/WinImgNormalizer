# M3-T07 owner review — accepted

Owner: Janne Vuorela. Prepared: 2026-10-05; guided partial review: 2026-10-06.
Original review commit: **4618a83cb2aa893aea9d6a1587987612624e1cf7**.
Current tested repair for recheck: **5a32fc6533c6f9fae1d36c271b515d1d1dd4d8ad**.
D36 authorizes the bounded colour correctness fix and regression tests; launcher and defaults remain unchanged.
The original packet's launcher copies and automated preparation retain their original
review binding in [preparation evidence](evidence/M3-T07.json). The fresh one-image
packet's repaired copies are bound in [repaired artifact evidence](evidence/M3-T07-repair-artifacts.json);
completed checks are in [repair evidence](evidence/M3-T07-repair.json).

## Original prepared packet — preserved

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

## Original full-packet checklist

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
baseline commit f9b00d8befac93c4fa1116efcbc2e9d7144f4509. Current columns are dated preparation
outputs from original review commit 4618a83cb2aa893aea9d6a1587987612624e1cf7. These are descriptive synthetic comparisons;
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

## Recorded owner smoke and partial review — 2026-10-06

The owner followed the physical `Drop this folder` → copied BAT instruction and
reported `Completed successfully. Press any key to close.` and clean final reporting.
The actual Pictures output was opened. The native BAT exit was not independently
captured. [Actual smoke evidence](evidence/M3-T07-owner-smoke.json) retains the
exact direct statements and separately verified artifact findings.

| Guided visual check | Owner response |
|---|---|
| Smooth landscape | Opens normally and looks smooth gradient |
| Hard edges/detail | Clear and intact |
| Transparent upper-left white area | Yes |

The last response confirms the prompted area only. It is not named-commit acceptance
of all colours, defaults or workflow. Other comparison images were not all visually
reviewed. Original files/dirs, video bytes, mirrors/collisions/duplicate mapping,
output timestamps, 20 JPEG decodes/geometry/metadata and complete balanced reports
were independently checked using only the approved synthetic trees.

The actual run used installed **Q16-HDRI**. The semi-transparent green patch fails
the independent reference: [0,213,169] instead of [127,213,169], error 127 versus
JPEG tolerance 12. The white area is correct. Nineteen outputs match preparation.
[Native diagnosis](evidence/M3-T07-hdri-diagnosis.json) reproduces the owner JPEG;
clamping after sRGB conversion and before white composition fixes this fixture in
isolated experiments. Application/tests were unchanged during that pre-fix diagnosis; those probes did
not substitute for the production repair validation recorded below.

## Completed repaired recheck — 2026-10-06

The first one-image attempt ran the older launcher; live command proof and its exact
old faulty JPEG are retained in [old-launcher evidence](evidence/M3-T07-recheck-old-launcher.json).
The guide's two Launcher window titles were unclear. Its area-check Yes supplies no
repaired appearance consent. A uniquely named **Fixed launcher 5a32fc6** folder then
provided exact tested copies, and the owner repeated the folder-to-BAT launch.

The owner reported **Completed successfully. Press any key to close.** Live CMD
proof binds the repaired BAT, adjacent PS1 and approved Colour recheck source.
[Fixed-run evidence](evidence/M3-T07-recheck-fixed-launcher.json) verifies corrected
JPEG bytes, full decode, reference error 1 within tolerance 12, stripped metadata,
balanced 1/1/no warnings, source preservation against its prior snapshot and output
timestamps. Native conversion 0 is logged; native BAT exit was not captured.

Final prompt: **Do you accept fixed version 5a32fc6 for the drag-and-drop workflow
and image appearance?** It followed the instruction to open selected Transparency.jpeg.

Direct owner response: **Accepted**.

Outcome: **accepted**. Approved implementation:
**5a32fc6533c6f9fae1d36c271b515d1d1dd4d8ad**. **T067 passed; M3-T07 done.**
[Acceptance evidence](evidence/M3-T07-acceptance.json) records the exact prompt/reply,
agent observation time and scope. Earlier observations remain dated; not every
original comparison image was individually visually reviewed, and no independent
screenshot/image-opening gesture is claimed.

This acceptance covers the named familiar workflow and image appearance. Existing
defaults remain; future tuning and publication operations retain their own approvals.
Live network shares remain owner-excluded D35. The next task remains unstarted.
