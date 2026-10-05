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

## I9 — M2-T04 exact extent bytes and best-effort targeting, 5 October 2026

[ImageMagick's defines reference](https://imagemagick.org/defines/) describes
`jpeg:extent` as an encoder-quality search toward a file-size budget and notes the
interaction with an explicit quality option. This task adds no quality option or
floor and retains the existing six resize attempts and JPEG settings.

The [pinned 7.1.2-32 JPEG encoder](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/coders/jpeg.c)
parses the extent with `SiPrefixToDoubleInterval`, compares encoded trial blob
lengths and selects a quality. The
[pinned string helper](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/string-private.h)
delegates that operand to
[InterpretSiPrefixValue](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/string.c).
Its K/M prefixes use decimal factors; an `i` selects the binary factor. Thus native
`1KB` means 1,000 bytes and `1MB` means 1,000,000, while `1KiB` and `1MiB` mean
1,024 and 1,048,576. PowerShell's binary `1KB`/`1MB` values must not be formatted
back into these decimal suffixes. An invariant decimal byte operand such as
`1048576B` removes that unit mismatch.

Retained real probes of the pinned Windows executable confirm byte-identical JPEG
pairs for `1KB`/`1000B`, `1KiB`/`1024B`, `1MB`/`1000000B` and
`1MiB`/`1048576B`. The larger seeded-noise fixture produced distinct 964,543-byte
decimal-MB and 1,011,926-byte binary-MiB outputs. A `1B` target still produced a
fully decoded 527-byte JPEG with native exit 0 and no diagnostics. Those results
support checking actual final length and preserving the existing valid-above-target
fallback with an explicit warning; native success alone cannot establish size
compliance.

The parser returns a `double`, so not every integer above 2^53 can be represented
exactly by the native search budget. The application keeps every accepted Int64
digit in its byte operand and uses the validated file's exact Int64 length for the
final at-or-below/above comparison. No huge-file experiment or universal native
precision guarantee is claimed.

Version-tagged source bytes, executable identity, synthetic recipes, native argv,
full-decode results and hashes remain in the ignored extent-research artifact.
Those unit probes and incidental timings are separate from T042's representative
benchmark and from owner photographic/aesthetic approval.


The reproducible development harness `tests/Measure-SizeQuality.ps1` compares
the exact frozen baseline `f9b00d8befac93c4fa1116efcbc2e9d7144f4509` with corrected
implementation `0e490cc5f972b5f22374d2599eed929e477b40ca`, using pinned ImageMagick
7.1.2-32. The measured synthetic corpus contains 1600x1200 landscape gradients,
1200x1600 portrait gradients, hard coloured edges with white/black line detail,
seeded RGB noise and an existing lossy JPEG reference; a separate opaque video
copy checks bytes and creation/modification times. Fresh PS5.1 and PS7 hosts ran
sequentially after local correctness tests. At 1,048,576, 262,144, 65,537 and 1,024
bytes, each host recorded 40 JPEG rows plus two video rows, with actual native
child exit 0, unchanged conversion flags and scale prefixes except extent, and
preserved source state. All five 65,537-byte control JPEGs are byte-identical
between baseline and candidate in each shell.

In the PS7 default-cap measurement, the seeded-noise JPEG changed from 979,962
to 1,015,369 bytes at 100 percent; decoded source-grid PSNR changed from 12.39
to 12.44 dB. The existing-lossy-JPEG case changed from 840,679 to 916,662 bytes
at 100 percent and from 25.01 to 25.04 dB. These observations qualify the
effective-budget correction rather than promise a universal quality improvement.
At 1,024 bytes, all five valid JPEGs remained above target at the final 50-percent
attempt: baseline returned 0 while the candidate returned 2 with exact warning
bytes/dimensions/scale and SizeWarnings=5. Every opaque video copy remained
byte-identical with its creation/modification times preserved.

RGB8 MAE and PSNR compare output with the original resized to output dimensions;
the separate source-grid metric compares upscaled output with the original
decoded source. The already-lossy reference is its existing decoded JPEG.
These encoded-sRGB metrics are not perceptual or photographic acceptance scores.
Native timing sums actual conversion attempts; complete application timing
includes validation, hash tracing and logging, and is reported per five-image
batch. One repeat per host is descriptive, with no confidence interval,
statistically controlled speed claim or subjective owner approval. The numeric
default remains 1,048,576 bytes; no quality floor, scale, chroma or encoder tuning
is introduced.

Sanitized benchmark notes `.scratch/M2-T04-benchmark-notes.json` have SHA256
`105d863414e201e6191b3dfb38bf39dfcf2cf994ffe1640c1ea70f277082c204`;
the compact comparison table is `.scratch/M2-T04-benchmark-owner-review.md`.
The notes retain raw/Git source bindings, fixture recipes/hashes, native parent
exits, artifact hashes, all metrics/timings and separate exploratory tooling
history. Raw native argv, transcripts and media remain ignored. Their scoped
summary is recorded in [M2-T04 evidence](evidence/M2-T04.json), with the mandatory
test/hosted gates and independent metric audit kept separate from owner approval.

The independent NumPy audit `.scratch/M2-T04-benchmark-metric-audit.json`
recomputed all 160 output-grid/source-grid metric pairs across the 84 rows and
checked 188 actual native conversions plus the five raw/Git source bindings per
host. This is an independent verification of the saved buffers and observations,
not an additional timing repeat or subjective quality approval.

## I10 — M2-T05 native diagnostics and justified retries

Microsoft's [Process.StandardOutput documentation](https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.process.standardoutput?view=netframework-4.8.1) describes full-pipe dependencies and separate reader threads for stdout/stderr. The Framework-compatible [StreamReader.Read(char[], int, int)](https://learn.microsoft.com/en-us/dotnet/api/system.io.streamreader.read?view=netframework-4.8.1) reads up to a specified character count. The reviewed [Framework AsyncStreamReader source](https://github.com/microsoft/referencesource/blob/main/System/services/monitoring/system/diagnosticts/AsyncStreamReader.cs) accumulates decoded characters before delivering complete lines, so limiting callback text alone would not bound a newline-free stream. The runtime instead uses fixed reads and prefix/tail storage, counts all decoded UTF-16 characters and rejects any truncation. These APIs establish capture mechanics, not a complete process-lifetime guarantee.

[Microsoft Win32 error definitions](https://learn.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-) distinguish sharing violation 32 and byte-range lock violation 33. The implementation's two retries with 100/200 ms backoff are a project decision. Generic Permission denied is not evidence of either code: the pinned native lock probe emitted that generic reason despite a separate CreateFileW query establishing code 32. Thus generic denial stays nonretryable.

Pinned ImageMagick 7.1.2-32 [exception.h](https://github.com/ImageMagick/ImageMagick/blob/7.1.2-32/MagickCore/exception.h) and [exception.c](https://github.com/ImageMagick/ImageMagick/blob/7.1.2-32/MagickCore/exception.c) define warning/error/fatal severities and emitted reason/location text. [operation.c](https://github.com/ImageMagick/ImageMagick/blob/7.1.2-32/MagickWand/operation.c) handles quiet, while [magick-cli.c](https://github.com/ImageMagick/ImageMagick/blob/7.1.2-32/MagickWand/magick-cli.c) handles regard-warnings status. [jpeg.c](https://github.com/ImageMagick/ImageMagick/blob/7.1.2-32/coders/jpeg.c) emits LosslessToLossyJPEGConversion for the intended lossy JPEG conversion. The allowance is that exact whole warning line at zero exit with complete capture and full JPEG validation; it does not accept every JPEG warning.

Genuine native provenance separates visible benign warning+zero without regard-warnings, the same visible warning+nonzero with regard-warnings, and production quiet suppression. An unsupported ICC version yielded native zero, an error diagnostic and a decodable JPEG; all must still reject. Resource probes use process-local 1-byte cache limits and make no machine-pressure claim. Downloaded tagged source hashes, native/probe bindings, actual wrapper/retry cases and honest development failures are recorded in [M2-T05 evidence](evidence/M2-T05.json) and [Windows CI evidence](evidence/M2-T05-ci.json). General deadlines/cancellation and owner acceptance remain later gates.

Actual wrapper cases include simultaneous newline-free streams, bounded prefix/tail retention and rejection of a valid JPEG after stream overflow. The eight-call boundary combines two controlled sharing failures with six real above-cap conversions; synthetic child diagnostics and COM padding remain identified as controls. Initial committed gates exposed an unchanged T005 test that expected an unexplained exit 9 to retry. That test was narrowed to an explicit transient result; the failed gates and exact source snapshots remain evidence. These results preserve the encoding/default policy and do not grant owner acceptance.

A further actual native result-shape probe found the existing development measurement trace read an absent legacy DiagnosticOutput field and recorded null despite a visible native warning. A narrow adapter now records the bounded separate streams with legacy fallback while returning the same result object to the application. Saved-AST checks passed five modern/legacy warning, quiet and stdout cases per host with complete JPEG decode. The genuine warning command deliberately omits quiet/regard. The edited-source probes, superseded passing I2 matrices and exact source snapshots are retained separately from fresh final I3 gates; no representative benchmark, encoder/default change or new quality/performance claim is made.

## I11 — M3-T01 scan and ancillary reporting

Microsoft's [FileSystemInfo.LastWriteTimeUtc documentation](https://learn.microsoft.com/en-us/dotnet/api/system.io.filesysteminfo.lastwritetimeutc?view=netframework-4.8.1) explains that enumerated metadata may be cached. The application refreshes each source entry and captures immutable creation/modified values before processing. Creation and modified timestamp setters are separate operations; this task catches and identifies each failed field while retaining already verified final bytes. The one-warning-per-output accounting and warning exit 2 are project policy, not an atomic filesystem guarantee.

[Console.Error](https://learn.microsoft.com/en-us/dotnet/api/system.console.error?view=netframework-4.8.1) provides the standard error writer. It offers an emergency route independent of PowerShell reporting cmdlets, but an absent/closed sink remains a best-effort limitation. [PipelineStoppedException](https://learn.microsoft.com/en-us/dotnet/api/system.management.automation.pipelinestoppedexception?view=powershellsdk-7.6.0) denotes an already stopped pipeline; explicit propagation is kept separate from ordinary log/timestamp I/O degradation. These APIs do not establish a cancellation deadline, native descendant cleanup or a guaranteed host interrupt exit.

The bounded fallback capacities, one-time disk-sink disablement, final reporting status, scan-completeness fields and skip/report link policy are the D29 implementation decisions. Actual Windows ACL, junction and controlled log/timestamp results must be recorded against their tested revision in M3-T01 evidence; this source note makes no execution or owner-acceptance claim.

## I12 — M3-T02 native lifetime, cache budgets and upstream review, 5 October 2026

Microsoft's [UpdateProcThreadAttribute reference](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-updateprocthreadattribute) documents JOB_LIST assignment during creation on Windows 10/Server 2016 or newer and an explicit HANDLE_LIST for inherited handles. The official [atomic job creation explanation](https://devblogs.microsoft.com/oldnewthing/20230209-00/?p=107812) explains the gap left by assigning a separately started process. The implementation creates directly in its private job, limits inheritance to the intended pipe handles and starts suspended until both readers are ready. Unsupported attributes or incompatible parent-job restrictions fail closed; there is no fallback to an uncontained process or process-name/PID-tree scan.

The [job object documentation](https://learn.microsoft.com/en-us/windows/win32/procthread/job-objects) describes descendant membership and [kill-on-close limits](https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-jobobject_basic_limit_information). [TerminateJobObject](https://learn.microsoft.com/en-us/windows/win32/api/jobapi2/nf-jobapi2-terminatejobobject) targets processes associated with that job. D30 requires a private noninheritable job without breakaway, verifies membership before resume, and checks that its active count reaches zero even if the root process exits while a descendant remains. Unrelated processes are outside that ownership boundary. Bounded raw reads and UTF-8 decoding retain the existing per-stream UTF-16 capture caps. [CancelSynchronousIo](https://learn.microsoft.com/en-us/windows/win32/api/ioapiset/nf-ioapiset-cancelsynchronousio) requests cancellation for a specified reader thread; it does not itself wait for completion. Termination, active-count checks and reader joins/cancellation share a fixed 3-second grace. A failed/incomplete drain cannot authorize a candidate. If the tree is not confirmed empty, the item keeps all cache and staging scratch rather than deleting data a remaining process may use. These mechanisms do not promise completion when a kernel driver refuses I/O cancellation or establish cooperative Ctrl+C behavior.

[ImageMagick resource documentation](https://imagemagick.org/resources/) describes pixel-cache memory/map/disk limits, thread limits and temporary-path environment variables. Its [security policy reference](https://imagemagick.org/security-policy/) describes installed resource ceilings. The [pinned 7.1.2-32 resource implementation](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/MagickCore/resource.c) uses the smaller requested/policy memory, map, disk, thread and time values. D30's process-local requests are 512 MiB memory, 1 GiB map, 2 GiB disk, two threads and the ceiling of the remaining shared image budget in seconds. No installed policy, persistent environment or machine settings are edited. These are pixel-cache budgets, not a total decoder/delegate heap or general file-write cap. Native elapsed-time enforcement is separate from ImageMagick's cooperative time checks, which can be disabled by SOURCE_DATE_EPOCH; synchronous filesystem copying also cannot be preempted by this wrapper.

The project allocates each native cache directory exclusively under the existing process Windows TEMP root. Before output writes the root must be canonical and regular, have no linked ancestors, fit within 160 characters and remain outside the source's physical directory identity. Only the child receives the plain WinImgNormalizer-cache-GUID path through TEMP, TMP and MAGICK_TEMPORARY_PATH. This addresses the observed pinned cache filename handling of deep/extended paths; source snapshots, ICC data and JPEG candidates retain same-output-volume staging. After confirmed job completion, cleanup is nonrecursive and accepts only regular magick-* members. Foreign names, directories and reparse entries remain with a warning. A concurrent same-pattern arrival is indistinguishable from native cache ownership; this policy is not an adversarial filesystem sandbox.

The upstream [jpeg:extent hang advisory](https://github.com/ImageMagick/ImageMagick/security/advisories/GHSA-gwr3-x37h-h84v), published 23 February 2026, declares 7.1.2-15 patched. Its affected 7.x field contains an inconsistent spelling; the declared patched version and [pinned 7.1.2-32 JPEG source](https://raw.githubusercontent.com/ImageMagick/ImageMagick/7.1.2-32/coders/jpeg.c) are the reviewed basis. The first extent search exits on a failed trial write; the second loop decrements its bound even when a write fails. No crafted exploit sample is run. The bounded benign sleeper/tree/resource tests are separate application defenses, not proof against every vulnerability.

The current required-codec notices for [PNG initialization leakage](https://github.com/ImageMagick/ImageMagick/security/advisories/GHSA-hvwf-wmm6-hvc3), [GIF quantization leakage](https://github.com/ImageMagick/ImageMagick/security/advisories/GHSA-w68r-6fmr-cvg4) and [XMP heap over-read](https://github.com/ImageMagick/ImageMagick/security/advisories/GHSA-jph3-37cg-79wg), published 27 September 2026, each declare 7.1.2-32 patched. This supports the retained D17 minimum and matching development pin at this dated review; later maintenance must recheck [upstream advisories](https://github.com/ImageMagick/ImageMagick/security/advisories). Saved primary-source hashes and qualified design/probe provenance remain ignored. Exact-I tests and desktop/hosted results must be recorded separately; this source note grants no future gate or owner acceptance.

## I13 — M3-T03 cooperative requests and console/force-exit boundaries

Microsoft's [SetConsoleCtrlHandler reference](https://learn.microsoft.com/en-us/windows/console/setconsolectrlhandler) describes per-process registration/removal and last-registered-first dispatch; another process's handler list is unaffected. [HandlerRoutine](https://learn.microsoft.com/en-us/windows/console/handlerroutine) runs on a new system-created thread. D31's callback therefore changes synchronized C# request state only, with no PowerShell invocation, reporting, cleanup or wait on that thread. Command execution enables scoped capture; callable execution defaults to no capture, and completion removes the handler. These are implementation choices, not identical event delivery or exits in every PowerShell host.

[GenerateConsoleCtrlEvent](https://learn.microsoft.com/en-us/windows/console/generateconsolectrlevent) cannot limit CTRL_C_EVENT to a nonzero process group; group zero broadcasts to the caller's shared console. Test delivery must therefore use a separate owned console after verifying its members. Redirected stdin ETX is not that Windows event. Actual event-driven tests with controlled pacing/routing remain distinct from unchanged public -File/BAT commands and directly injected request state. A successful event-send return alone proves no application handling, output cleanup or interrupted summary.

The Framework-compatible [Stream.Read(byte[], int, int)](https://learn.microsoft.com/en-us/dotnet/api/system.io.stream.read?view=netframework-4.8.1) may block and return fewer requested bytes. D31's 256-KiB copy boundaries and 50-ms native-wait polling cannot bound an individual filesystem call. Guarded finalization retains [File.Move(source, destination)](https://learn.microsoft.com/en-us/dotnet/api/system.io.file.move?view=netframework-4.8.1), which refuses an existing destination. Request-versus-move ordering belongs to shared application state, not to a cancellation guarantee in that API. A completed validated final remains; incomplete candidates stay nonfinal.

[Console control-handler documentation](https://learn.microsoft.com/en-us/windows/console/console-control-handlers) describes default exit and attachment boundaries. [TerminateProcess](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-terminateprocess) stops threads and initiates asynchronous termination; pending I/O may delay actual exit. Forced host/process exit or terminal closure can bypass cleanup/reporting. Recognizable owned scratch is not validated final media. Cooperative application code 130, bounded INTERRUPTED reporting and exact cleanup require actual-host evidence; they are not universal force-exit or unchanged-BAT guarantees.

Observed official-page bytes and hashes remain in ignored M3-T03 documentation scratch. This note supplies mechanics and chosen policy, with no future focused/full/CI or owner-acceptance claim.
