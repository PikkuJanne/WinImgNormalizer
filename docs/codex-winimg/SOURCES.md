# Primary-source reference index

Prepared 4 October 2026. Pinned repository URLs identify the audited source; external docs and release collections can change. Re-check exact syntax and versions during implementation. This plan is mostly original engineering requirements, not copied external documentation. Initial runtime results are deliberately not supplied.

## R1 — Pinned PowerShell source

https://github.com/PikkuJanne/WinImgNormalizer/blob/4b918da50639d8fc7e3fddc6d6b3678880ea4098/WinImgNormalizer.ps1

Static source anchors and existing behavior.

## R2 — Pinned batch launcher

https://github.com/PikkuJanne/WinImgNormalizer/blob/4b918da50639d8fc7e3fddc6d6b3678880ea4098/WinImgNormalizer.bat

Launcher invocation, pause and exit handling.

## R3 — Pinned README

https://github.com/PikkuJanne/WinImgNormalizer/blob/4b918da50639d8fc7e3fddc6d6b3678880ea4098/README.md

Existing public contract, prerequisites and limitations.

## R4 — Pinned repository tree

https://api.github.com/repos/PikkuJanne/WinImgNormalizer/git/trees/4b918da50639d8fc7e3fddc6d6b3678880ea4098?recursive=1

Seven-file baseline and existing branding assets.

## R5 — Pinned MIT license

https://github.com/PikkuJanne/WinImgNormalizer/blob/4b918da50639d8fc7e3fddc6d6b3678880ea4098/LICENSE

Retain authorship and license text.

## R6 — Release collection

https://api.github.com/repos/PikkuJanne/WinImgNormalizer/releases

Mutable distribution metadata; recheck before release planning.

## I1 — ImageMagick command-line options

https://imagemagick.org/command-line-options/

Read/profile/strip/frame/output semantics. Revalidate exact commands against installed version.

## I2 — ImageMagick defines

https://imagemagick.org/defines/

jpeg:extent and literal-filename-related behavior; verify byte units.

## I3 — ImageMagick security policy

https://imagemagick.org/security-policy/

Resource policies, codecs and installed-policy boundaries.

## I4 — Upstream jpeg:extent advisory GHSA-gwr3-x37h-h84v

https://github.com/ImageMagick/ImageMagick/security/advisories/GHSA-gwr3-x37h-h84v

Relevant 2026 hang risk. Read upstream release/fix information rather than trusting an inconsistent version-range string.

## P1 — Microsoft about_PowerShell_exe

https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_powershell_exe?view=powershell-5.1

Native -File, arguments, process policy and exit/interrupt semantics.

## P2 — Microsoft about_Automatic_Variables

https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables?view=powershell-7.5

LASTEXITCODE and host state; consult corresponding target version.

## T1 — Pester quick start

https://pester.dev/docs/quick-start

PowerShell tests, assertions, discovery and runner setup.

## G1 — Git push documentation

https://git-scm.com/docs/git-push

Explicit normal push, remote refspec and rejection semantics.

## G2 — Git ls-remote documentation

https://git-scm.com/docs/git-ls-remote

Independently read advertised branch SHA.

## G3 — GitHub secure use reference

https://docs.github.com/en/actions/reference/security/secure-use

Least privileges, untrusted code and reviewed action pins.

## O1 — OpenAI AGENTS.md guidance

https://developers.openai.com/codex/guides/agents-md

Scoped repository instructions; current documentation may redirect to ChatGPT Learn.
