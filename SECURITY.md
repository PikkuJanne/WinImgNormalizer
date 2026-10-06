# Security and reporting

WinImgNormalizer processes files locally. It has no upload, telemetry, account or
automatic dependency-update service. The current application version is
**unreleased**; prepared packages are unsigned. There is no published release
support schedule or guaranteed response time.

## Report a security concern

Check this repository's [Security page](https://github.com/PikkuJanne/WinImgNormalizer/security).
If GitHub offers **Report a vulnerability**, use that private reporting route.
Private vulnerability reporting was disabled when checked on 6 October 2026.

If a private route is unavailable, open a
[minimal public issue](https://github.com/PikkuJanne/WinImgNormalizer/issues/new)
asking the maintainer for a private security contact. State that you have a security
concern, without publishing exploit details, sensitive files or personal logs.
An ordinary GitHub issue is public. See GitHub's
[reporting guidance](https://docs.github.com/en/code-security/how-tos/report-and-fix-vulnerabilities/report-privately).

Once a suitable route is agreed, include the application version or exact commit,
Windows and PowerShell versions, `magick.exe -version`, expected and observed
behavior, possible security impact, and a minimal reproduction using synthetic
data where possible. Ordinary bugs can use public issues after removing private
information.

## Protect information when seeking help

Logs and CSV reports can expose source paths, filenames and native diagnostics.
Review and sanitize them before sharing. Do not post personal media, credentials
or complete raw reports in public. JPEG derivatives lose source image metadata
after orientation and colour handling; copied videos retain their original bytes
and metadata. Output in cloud-managed Pictures can be synchronized by the user's
sync software. Local processing does not prevent that synchronization.

## Dependency and download boundaries

Obtain ImageMagick separately from its
[official site](https://imagemagick.org/download/#windows-binary-release), keep it
patched, and respect the installed
[security policy](https://imagemagick.org/security-policy/). The application does
not install dependencies or change machine-wide policy. Run with ordinary source,
destination and temporary-storage permissions; do not disable protection software
or weaken global execution/ImageMagick policy to make a file process.

Native image parsing carries residual risk. Per-image resource/time limits and
owned-process cleanup are practical safeguards, not an exploit-proof sandbox.
Checksums identify bytes relative to a trusted expected value; they do not prove
publisher identity. Review the source and verify trusted package hashes before
selectively unblocking downloaded files. No verified-signature claim is made.
