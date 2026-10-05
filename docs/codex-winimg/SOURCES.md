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

## I5 — M1-T01 dependency review, 4 October 2026

https://github.com/ImageMagick/ImageMagick/releases/tag/7.1.2-32

Current stable release 7.1.2-32, published 27 September 2026. The product minimum
is this reviewed release; the existing test pin already matches it. Recheck upstream
notices for future maintenance rather than treating this floor as permanent safety.

https://github.com/ImageMagick/ImageMagick/security/advisories/GHSA-xxrx-rjp3-27rj

https://github.com/ImageMagick/ImageMagick/security/advisories/GHSA-jph3-37cg-79wg

JPEG decoder disclosure and XMP profile over-read fixes are included in 7.1.2-32.
The earlier jpeg:extent hang advisory names patch 7.1.2-15 but has an inconsistent
affected-range string. The actual fix is included in 32:

https://github.com/ImageMagick/ImageMagick/commit/c448c6920a985872072fc7be6034f678c087de9b

https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/magick.c

ListMagickInfo emits format rows with an optional module column and r/w/adjoin
mode. A star denotes native blob support, not read capability. Compiled formats
do not override security policy or certify arbitrary input files.

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


## W1 — M1-T02 Windows directory APIs, 4 October 2026

Reviewed primary Microsoft documentation: [CreateDirectoryW](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createdirectoryw) fails for an existing directory and creates only the last component; [GetFinalPathNameByHandleW](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-getfinalpathnamebyhandlew) returns normalized final paths and documents DOS/UNC prefix handling and SMB permission limitations; [CreateFileW](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew) permits directory handles with FILE_FLAG_BACKUP_SEMANTICS. The implementation uses no-access directory handles with shared read/write/delete, checks native errors and closes handles. Native calls add the documented extended drive/UNC prefix internally for generated descendants; this avoids truncating names or depending on process long-path opt-in. An inaccessible canonical query fails setup; it does not substitute an unverified lexical alias. These references support API use, not unexecuted UNC or adversarial race claims.

## W2 — M1-T03 no-overwrite file operations, 4 October 2026

Reviewed primary Microsoft documentation: [File.Move(String,String)](https://learn.microsoft.com/en-us/dotnet/api/system.io.file.move?view=netframework-4.8.1) throws IOException when the destination exists; [File.Copy(String,String,Boolean)](https://learn.microsoft.com/en-us/dotnet/api/system.io.file.copy?view=netframework-4.8.1) also refuses an existing destination when overwrite is false. Both selected overloads are available to Windows PowerShell 5.1. The newer three-argument Move overload is unnecessary. Image candidates and final paths share the owned run volume. Checks immediately before finalization improve diagnostics; the no-overwrite operations protect an arrival after the check. These APIs do not establish batch atomicity, crash durability, validated image content, complete interrupted video-copy recovery or isolation from hostile ancestor mutation.

## W3 — M1-T04 candidate validation and staged streams, 4 October 2026

[ImageMagick identify](https://imagemagick.org/identify/) documents `+ping` for reading pixel characteristics instead of the default metadata-only probe. The validator uses it with stored-format/frame/dimension checks. [regard-warnings](https://imagemagick.org/command-line-options/#regard-warnings) turns some format warnings into errors; the supported JPEG build must also pass the real truncated-output regression. These references do not promise rejection of every possible corrupt file.

[FileStream constructors](https://learn.microsoft.com/en-us/dotnet/api/system.io.filestream.-ctor?view=netframework-4.8.1) document that CreateNew fails if the path already exists. [FileShare](https://learn.microsoft.com/en-us/dotnet/api/system.io.fileshare?view=netframework-4.8.1) permits concurrent readers with Read while refusing cooperative writes, and None refuses other opens while the stream is held. Videos are copied into an exclusively reserved partial, flushed and closed, then checked against source length and modification time. [File.Move(String,String)](https://learn.microsoft.com/en-us/dotnet/api/system.io.file.move?view=netframework-4.8.1) refuses an existing final name; the application checks equal volume roots because cross-volume Move can copy/delete. These checks do not establish content identity for a same-length/same-timestamp mutation, batch atomicity, crash durability or hostile-filesystem isolation.

## I6 — M2-T01 frame selection, canvas placement and HEIC order, 4 October 2026

[Input frame selection](https://imagemagick.org/command-line-processing/#input)
documents `-define image:frames=list` as an alternative to filename bracket syntax.
The [pinned 7.1.2-32 SetImageInfo source](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/image.c)
reads `image:frames` when a filename has no subimage component. The application
selects index 0 on an owned neutral snapshot, rather than appending a selector to an
unexamined user filename. Native argument and general input-grammar work remains
separate from this limited selection policy.

[Coalesce documentation](https://imagemagick.org/command-line-options/#coalesce)
describes displayed animation frames using page offsets and disposal metadata.
The [pinned CoalesceImages implementation](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/layer.c)
allocates the first virtual canvas and places its image at the recorded offset.
The application coalesces only the selected actual GIF/WebP image; TIFF/HEIC pages
are not overlaid. Actual offset, pixel and orientation fixtures establish the tested
behavior; these references do not certify every animation or colour case.

The [pinned HEIC reader](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/coders/heic.c)
uses the sequence track reader where available; otherwise it reads the primary
image before other top-level images. The policy follows that decoder order.
Counts describe decoder-exposed images, not every container item or thumbnail.
T031 uses an actual two-image still collection plus animated WebP; timed HEIC
animation, auxiliary images and arbitrary codec builds remain unverified.

## I7 — M2-T02 colour profiles, orientation and alpha, 4 October 2026

[Profile operations](https://imagemagick.org/command-line-options/#profile) run in
command order. The [pinned 7.1.2-32 ProfileImage implementation](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/profile.c)
distinguishes assigning a first profile from transforming pixels between existing
and target profiles, using LittleCMS for the latter. Removing a source profile
before that transform loses its characterization.

The [pinned property implementation](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/property.c)
can return free-form image properties before computed metadata. Separate inspection
therefore clears conflicting profile/colorspace text properties while preserving
the actual ICC data. The [ICC description of profile tag tables](https://www.color.org/security/malformed/added-bytes/)
explains the 128-byte header followed by signature/offset/length entries. Bounded
structural and decoded-model checks support a fail-closed policy; they do not
establish complete ICC conformance or semantic accuracy. Actual native diagnostic
and malformed/model-mismatch behavior is retained as synthetic test/probe evidence.

[Colour management](https://imagemagick.org/color-management/) distinguishes
declaring a colorspace from converting pixel values, and describes untagged sRGB
assumptions. [Auto-orient](https://imagemagick.org/command-line-options/#auto-orient)
uses orientation metadata while present;
[strip](https://imagemagick.org/command-line-options/#strip) removes profiles and
comments. [Alpha remove](https://imagemagick.org/command-line-options/#alpha)
composites against the selected background. These semantics support a tested
ordering; actual profile licenses, reference values and separate lossy-JPEG
tolerances remain evidence requirements.

The [pinned Compact ICC profile source](https://raw.githubusercontent.com/saucecontrol/Compact-ICC-Profiles/bdd84663061bc4ae95ca70decff54f581e27f702/readme.md)
publishes the selected sRGB-v4, AdobeCompat-v2 and CGATS001Compat-v2-micro profiles
under [CC0](https://raw.githubusercontent.com/saucecontrol/Compact-ICC-Profiles/bdd84663061bc4ae95ca70decff54f581e27f702/license).
The compact CMYK profile contains only an A2B0 perceptual display mapping. A requested
rendering intent does not add absent characterization; reference evidence must
state the engine's actual supported/fallback mapping and avoid promising general
printing accuracy or reverse CMYK conversion.

[Pillow ImageCms](https://pillow.readthedocs.io/en/stable/reference/ImageCms.html)
provides source/target profile transforms with explicit intent and flags. The
independent synthetic patch recipe records its Pillow/LittleCMS versions and keeps
pre-JPEG transform tolerance separate from lossy JPEG tolerance; mandatory tests
consume checked reference data without a Python/Pillow runtime dependency.

## I8 — M2-T03 literal filenames, native argv and directory identity, 5 October 2026

[ImageMagick command-line processing](https://imagemagick.org/command-line-processing/)
describes filename globs, frame selectors and embedded scene/property formatting,
and documents `registry:filename:literal=true` for literal formatting. The
[defines reference](https://imagemagick.org/defines/) also documents the shorter
`filename:literal` output define. PowerShell literal paths do not by themselves
disable this separate native grammar.

The [pinned 7.1.2-32 InterpretImageFilename implementation](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/image.c)
returns the unexpanded filename when its literal registry setting is true.
[ReadImages/PingImages](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/constitute.c)
also apply filename interpretation while selecting scenes. The retained actual
red-versus-blue input probe shows that `image:frames=0` can select a numeric
counterpart in a percent ancestor when the guard is absent; placing it before
input restores the intended pixels. This is a measured pinned-build result, not
an inference that neutral basenames remove all ancestor syntax.
The [pinned filename expansion code](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/utility.c)
and [registry define handling](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickWand/operation.c)
support the explicit process-local choice; checked synthetic files establish its
actual behavior. Version-tagged source bytes and native probe hashes stay in the
ignored provenance manifest.

[CMD](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/cmd)
documents quoting and delayed exclamation expansion.
[SETLOCAL](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/setlocal)
localizes `DisableDelayedExpansion` to a batch invocation.
[PowerShell parsing](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_parsing)
describes native argument boundaries and the legacy handling of CMD/BAT.
Actual receiver/application tests support the conditional trailing-separator
suffix and tested argv identities. Caller CMD substitution that occurs before
entry is a separate limitation; use the documented direct PowerShell literal form
for paired percent expressions. These observations do not certify Explorer use,
default Pictures routing or general final BAT exit propagation.

[Windows path limits](https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation)
distinguish extended paths, ordinary MAX_PATH, process opt-in and shell support.
The actual supported 324-character source and complete output tests do not imply
universal provider or disabled-policy support. No host policy is changed.

[GetFileInformationByHandleEx](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getfileinformationbyhandleex)
documents `FileIdInfo` and its supported filesystem/network technologies.
The [information-class enumeration](https://learn.microsoft.com/en-us/windows/win32/api/minwinbase/ne-minwinbase-file_info_by_handle_class)
places that class at 18 and states its Windows 8/Server 2012 availability boundary.
[FILE_ID_INFO](https://learn.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-file_id_info)
combines a volume serial with a 128-bit identifier for comparing handles on one
computer. The runtime rejects unavailable or unusable identities. Actual local
drive/localhost C$ ancestor rejection supports this conservative preflight policy;
it does not certify arbitrary remote SMB equivalence or a hostile concurrent
filesystem sandbox. Live UNC capability remains separately observed from the
mandatory assertion count.
