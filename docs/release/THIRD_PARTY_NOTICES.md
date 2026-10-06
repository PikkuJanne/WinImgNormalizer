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

ImageMagick is a separately obtained prerequisite. Its installer, executable,
delegates and development test dependencies are excluded from this package.
