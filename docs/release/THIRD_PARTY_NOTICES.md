# Third-party notices

WinImgNormalizer is authored by Janne Vuorela and distributed under the MIT license
in `LICENSE`, retaining `Copyright (c) 2025 Janne Vuorela`.

The application embeds the **sRGB-v4 ICC profile** from
[saucecontrol/Compact-ICC-Profiles](https://github.com/saucecontrol/Compact-ICC-Profiles),
dedicated under **CC0 1.0 Universal (CC0-1.0)**. The full retained license text is
[LICENSE-CC0.txt](LICENSE-CC0.txt) beside this notice in the portable package.

- Immutable upstream commit: `bdd84663061bc4ae95ca70decff54f581e27f702`.
- Original path: `profiles/sRGB-v4.icc`.
- [Exact upstream profile](https://raw.githubusercontent.com/saucecontrol/Compact-ICC-Profiles/bdd84663061bc4ae95ca70decff54f581e27f702/profiles/sRGB-v4.icc).
- Profile length: 480 bytes.
- SHA-256: `c56e1685d888f5edb92fe07f2750f387f8fe8e91b32ff8fb0b56bfbbb9458353`.

These are the decoded profile bytes embedded in `WinImgNormalizer.ps1`, used as the
target for conversion to sRGB before image metadata removal. No separate profile
installation is required. The other colour profiles and synthetic test references
in the development repository are excluded from the portable package.

The portable package includes the unchanged upstream **ImageMagick 7.1.2-35
(Beta), Q16-HDRI x64** executable and its runtime configuration, ICC profile,
license and complete delegate notices. ImageMagick is copyright ImageMagick
Studio LLC and licensed under the ImageMagick License. Full license and delegate
texts are retained in `runtime/imagemagick/LICENSE.txt` and
`runtime/imagemagick/NOTICE.txt`; the exact upstream CycloneDX inventory is
`runtime/imagemagick/upstream-sbom.cdx.json`.

The static executable includes LGPL/MPL and permissive components. Complete
corresponding Windows/delegate source and build scripts are available beside the
binary download as `ImageMagick-7.1.2-35-Windows.7z`; see
`runtime/SOURCE_INFORMATION.md` for source/relink instructions and rights. Source
asset identities are in `runtime/runtime-manifest.json` and outer build provenance.
No development test dependencies or installer are included in the portable ZIP.
