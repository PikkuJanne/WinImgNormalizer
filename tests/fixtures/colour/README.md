# Synthetic colour references

The three ICC files are exact bytes from the CC0 Compact-ICC-Profiles repository
at commit `bdd84663061bc4ae95ca70decff54f581e27f702`. `MANIFEST.json` records immutable
URLs, byte counts, SHA256, Git blob identities and header observations. The original
CC0 license is retained as `LICENSE-CC0.txt`.

`REFERENCE.json` records synthetic patch values and independent Pillow 12.3.0 /
LittleCMS 2.19 outputs. `Generate-Reference.py` is an optional reproducible recipe
that prints the colour tuples; it installs nothing and writes no media. The
PowerShell suite generates owned RGB and true four-channel CMYK fixtures and uses
precomputed references, so neither Python nor Pillow is a test/runtime requirement.
Its EXIF template contains only invented camera strings, dates and GPS coordinates.

The reference requests Relative colorimetric intent (1), with black point
compensation off. CGATS001Compat-v2-micro provides only an A2B0 perceptual forward
mapping; LittleCMS falls back to that available mapping for this request. Passing
this fixture does not certify general relative-colorimetric printer accuracy,
other intents or reverse CMYK transforms.

Pre-JPEG channel tolerance is 3/255; final JPEG patch-center tolerance is 12/255.
Actual pinned native lossless transforms differed by at most 1/255 from the
independent reference. All assets here are hashed by the Windows test runner.
Generated fixtures, outputs, native probes and raw logs stay in ignored scratch.
