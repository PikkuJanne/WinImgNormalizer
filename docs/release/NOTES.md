This patch repairs JPEG colour-profile handling in the existing Windows launcher.
An ordinary readable colour JPEG is retained in the viewing copy when only its
embedded ICC metadata is unusable. Original files remain untouched.

## Colour repair

- Primary JPEG ICC APP2 pieces are reconstructed by their sequence numbers with
  bounded parsing. Valid split and reordered profiles preserve their wide-gamut
  characterization. The exact validated source bytes are applied before the sRGB
  transform. A narrowly recognized three-zero-byte suffix outside an otherwise
  valid declared profile can be removed on an owned copy.
- When ICC metadata is unusable, an otherwise sound ordinary RGB-family JPEG can
  be decoded once with that profile excluded. A trustworthy supported alternative
  characterization takes precedence; otherwise the viewing copy assumes sRGB and
  records: “JPEG saved; embedded colour profile unusable; sRGB assumed, colours
  may differ.” This is an approximate preview, counted as converted with a warning
  and exit code 2. It is not an accurate ICC conversion.
- Pixel damage, unidentified decoder errors, policy/resource failures and
  uncharacterized CMYK/YCCK remain failures. Clean decode and final JPEG validation
  still precede the no-overwrite final move. A failed item cannot suppress a later
  usable sibling.
- Recurring colour-warning reasons receive one short console notice and final
  counts; full per-file colour paths, reasons and outcomes stay in the local log/CSV.

## Preserved behavior

The private offline ImageMagick 7.1.2-35 Beta Q16-HDRI runtime, dependency selection,
minimum version and launcher remain unchanged. No new setting, prompt, application
dependency or user action is required. Windows PowerShell 5.1 and PowerShell 7
remain supported.

Orientation, sRGB conversion before metadata stripping, white transparency
flattening, first-frame/page selection, the 1,048,576-byte best-effort target,
progressive sizing, duplicate rules, deadlines, cancellation and exit 0/1/2/130
contracts remain. Videos are still copied with unchanged bytes and metadata.
Processing stays local; this repair does not deploy the website.

## Verification and release identity

The colour regression corpus distinguishes colour-managed wide-gamut outputs,
explicit approximate previews and genuine failures using independent references.
Observed Windows, launcher and exact public-download results are recorded in the
sanitized JPEG-ICC-1.0.3 evidence after each gate completes. Checksums and provenance
bind the published bytes to their source revision. The application remains unsigned.

Earlier releases and their historical evidence remain at their original tags and
download bytes. Version 1.0.2 repaired cold clone/source-ZIP/cache distribution;
this patch changes only colour handling and its tests, descriptions and records.

## Authorship and license

WinImgNormalizer is authored by Janne Vuorela under the MIT License, retaining
Copyright (c) 2025 Janne Vuorela. Existing third-party licenses and runtime source
sidecars are retained. The supported EXIF Adobe RGB alternative embeds the same
CC0 project's standard profile, with its exact provenance in THIRD_PARTY_NOTICES.md.
