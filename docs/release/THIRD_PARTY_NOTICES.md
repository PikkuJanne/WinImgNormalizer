# Third-party notices

WinImgNormalizer is authored by Janne Vuorela and distributed under the MIT license
in `LICENSE`, retaining `Copyright (c) 2025 Janne Vuorela`.

The application embeds the **sRGB-v4** and **AdobeCompat-v2 ICC profiles** from
[saucecontrol/Compact-ICC-Profiles](https://github.com/saucecontrol/Compact-ICC-Profiles),
dedicated under **CC0 1.0 Universal (CC0-1.0)**. The full retained license text is
[LICENSE-CC0.txt](LICENSE-CC0.txt) beside this notice in the portable package.

- Immutable upstream commit: `bdd84663061bc4ae95ca70decff54f581e27f702`.
- Original path: `profiles/sRGB-v4.icc`.
- [Exact upstream profile](https://raw.githubusercontent.com/saucecontrol/Compact-ICC-Profiles/bdd84663061bc4ae95ca70decff54f581e27f702/profiles/sRGB-v4.icc).
- Profile length: 480 bytes.
- SHA-256: `c56e1685d888f5edb92fe07f2750f387f8fe8e91b32ff8fb0b56bfbbb9458353`.

The supported explicit EXIF Adobe RGB characterization uses the same upstream
project's standard Adobe RGB compatible profile:

- Original path: `profiles/AdobeCompat-v2.icc`.
- [Exact upstream profile](https://raw.githubusercontent.com/saucecontrol/Compact-ICC-Profiles/bdd84663061bc4ae95ca70decff54f581e27f702/profiles/AdobeCompat-v2.icc).
- Profile length: 374 bytes.
- SHA-256: `60fb2adecacf82132db0b1c09b303316f3bbd9e2823e7ba096d01627d12d57c9`.

These decoded bytes are embedded in `WinImgNormalizer.ps1`. sRGB-v4 is the target
for conversion before image metadata removal; AdobeCompat-v2 is used only when
supported explicit EXIF characterization supplies Adobe RGB after unusable JPEG
ICC metadata. No separate profile installation is required. Other development
profiles and synthetic test references are excluded from the portable package.

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

The engine also requires **Microsoft OpenMP VCOMP140.DLL 14.51.36247.0 x64**,
copyright Microsoft Corporation. The unchanged, signed DLL is shipped beside
`magick.exe`; a machine-wide Visual C++ installation is not required. It is
extracted at build time from Microsoft's authenticated Visual C++ v14
Redistributable and covered by the applicable Microsoft Visual Studio
redistribution rights and runtime license, separately from the open-source
ImageMagick/delegate licenses. The authoritative original license document and a
complete plain-text extraction are retained in
`runtime/imagemagick/MICROSOFT-VISUAL-C-RUNTIME-LICENSE.docx` and `.txt`.
`runtime/runtime-manifest.json` and build provenance pin the official source URL,
independently published Microsoft WinGet digest, signatures, exact DLL and license
bytes. No redistributable installer is executed or included in the portable ZIP.
