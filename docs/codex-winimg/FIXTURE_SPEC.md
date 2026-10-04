# Synthetic fixture plan

Implement a deterministic fixture builder in M0-T02; this document is a recipe,
not a claim that binaries or their tests have already been generated. Store build
recipes and small redistributable fixtures under the repository's chosen `tests/`
layout. Keep generated outputs in an explicitly ignored disposable directory.
The test harness must inject a destination or run in a disposable Windows account;
never rely on temporarily redefining USERPROFILE to override Windows Known Folders.

## Required corpus

| Group | Recipe and independent checks |
|---|---|
| Ordinary images | Small RGB colour patches, gradients, lines, portrait/landscape dimensions and seeded noise. Produce JPEG, PNG, BMP and TIFF with a current known ImageMagick build. Inspect format, dimensions and basic pixel values independently. |
| Collisions | Put different source formats under the same stem; include a pre-existing suffixed stem, a `photo.jpeg` directory and Windows case-only collision planning mocks. |
| Transparency | Transparent and semitransparent PNG pixels with known foreground colours and white-composited reference values. |
| Orientation | Deliberately tagged EXIF orientations 1 through 8 with asymmetric corner markers. Verify metadata is really embedded, not just an in-memory ImageMagick property. |
| Frames/pages | Two-frame GIF with an offset canvas, animated WebP when available, two clearly distinct TIFF pages and verified HEIC/HEIF primary/sequence behavior. |
| Profiles | Trusted, redistributable sRGB, wide-gamut RGB and CMYK fixtures with genuine embedded profiles, known patch values and provenance/license records. Untagged and benign malformed-profile cases. Do not label an RGB image as CMYK without actually converting it. |
| Video copy | Seeded opaque bytes under supported video suffixes; the tool copies bytes and does not decode video. A real short video is optional for manual acceptance, with provenance. |
| Duplicate timing | Set equal LastWriteTimeUtc deliberately, with same/different lengths. Simulate a first failure using a process/filesystem test seam rather than depending on fragile filesystem timing. |
| Invalid outputs | Benign invalid bytes, zero-byte files, header-only/truncated JPEGs and controlled process fakes that emit exit/stream combinations. No exploit payload required. |
| Filename parsing | Legal platform filenames with spaces, brackets, percent/%d, apostrophes, ampersands, parentheses, exclamation marks, accented letters, emoji and CSV-leading characters. Windows-illegal cases remain mocks with an explicit reason. |
| Environment failures | Inject denied writes, log failures, stalled processes, source changes and low-resource budgets. Real ACL/junction cases only inside an owned disposable Windows tree. |

Fixture manifest fields: ID, generation recipe, seed, generating tool/version,
license/source when external, expected format/dimensions/frame count/profile,
SHA-256, and privacy status (`synthetic` or `owner-approved-public`). Never add
private photos, source logs revealing private paths or copyrighted sample libraries
without permission. Real owner acceptance copies stay outside the repository.

For colour tests, compare against a trusted profile-aware reference. Separate
pre-JPEG colour-transform tolerance from JPEG encode/decode tolerance; document
threshold rationale. A golden output created by the same buggy code path is not an
independent oracle. Do not demand byte-identical JPEGs across codec versions.

For HEIC/WebP positives, detect actual decoder availability. Optional local jobs may
skip with a reason; at least one codec-enabled required job or recorded target run
must execute each format before advertising it as tested. Absent codecs also need
mandatory negative-path tests. `Get-ChildItem` zero-results, a skipped suite or a
missing Pester module must not produce a fake green test report.
