# ImageMagick source and rebuilding

The bundled engine is the unchanged upstream ImageMagick 7.1.2-35 **Beta**,
Q16-HDRI x64 portable build from the official
[31 October 2026 workflow](https://github.com/ImageMagick/ImageMagick/actions/runs/41175952230),
source revision `39ad88fc5637b0ed9d2adbf37c3c5f6f174c63a3`. It was selected because
the latest tagged 7.1.2-34 portable build carries libheif 1.23.6; the chosen build
contains libheif 1.23.7 with the subsequently published HEIF security fixes.
The WinImgNormalizer application release is stable; this upstream dependency is
accurately identified as Beta. No claim of universal security is made.

ImageMagick is copyright ImageMagick Studio LLC. `imagemagick/LICENSE.txt` and
`imagemagick/NOTICE.txt` retain the complete upstream license and delegate notices;
`imagemagick/upstream-sbom.cdx.json` identifies the exact dependency inventory.
These files are unchanged. No other ImageMagick executables are included.

The exact complete upstream Windows source/build archive, including delegate
sources and build scripts, is offered beside the binary downloads at the
[same release](https://github.com/PikkuJanne/WinImgNormalizer/releases/tag/v1.0.1):
[ImageMagick-7.1.2-35-Windows.7z](https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.1/ImageMagick-7.1.2-35-Windows.7z).
Its SHA-256 and size are recorded in `runtime-manifest.json`, the outer build
provenance, and `SHA256SUMS.txt`. It includes the preferred form of the ImageMagick
and LGPL/MPL delegate source and the Windows configuration/build system. Follow
the archive's Windows build instructions to rebuild/relink ImageMagick against
modified library sources. A suitable developer compiler is needed only for that
optional rebuild, never for using WinImgNormalizer.

You may modify those library components and reverse engineer the engine for
debugging such modifications, as permitted by their included licenses. Replacing
the engine requires regenerating the development integrity manifest with the
reviewed staging tool; a modified portable package is not an official release.
WinImgNormalizer does not restrict rights granted by the component licenses.

The exact original portable archive and SBOM are also mirrored as separate release
assets for reproducible developer staging after upstream Actions artifacts expire.
No dependency is downloaded while media are processed.
