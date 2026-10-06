# Getting started with WinImgNormalizer

This portable application package is **unreleased and unsigned**. Its application
version and exact source revision are recorded in `package-manifest.json`; release
notes are in [CHANGELOG.md](CHANGELOG.md). Preparation of a ZIP does not establish a
public download, signed publisher or released tag.

## Verify and extract

Obtain the ZIP and its expected SHA-256 through a source you trust. Compare the
actual ZIP hash with the exact filename's entry in `SHA256SUMS.txt`, and with the
hash provided through the trusted channel. Replace the example path with your
downloaded filename:

```powershell
Get-FileHash -LiteralPath 'C:\Downloads\WinImgNormalizer.zip' -Algorithm SHA256
```

A matching checksum identifies the bytes relative to that trusted value; it does
not authenticate the publisher. Review the source before running it. If you trust
the verified package and Windows marks it as downloaded, selectively unblock that
ZIP before extracting it:

```powershell
Unblock-File -LiteralPath 'C:\Downloads\WinImgNormalizer.zip'
```

Extract into a new ordinary folder in Downloads, such as `WinImg Normalizer`. Keep
`WinImgNormalizer.bat` and `WinImgNormalizer.ps1` together with their original names.
If the extracted PS1 remains blocked, review it and selectively unblock that exact
file with `Unblock-File -LiteralPath`; do not unblock an entire Downloads tree.
The command removes a downloaded-file marker; it does not change execution policy
or make an unsigned script satisfy an `AllSigned` policy.
See Microsoft's [hashing documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-filehash)
and [selective unblocking documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/unblock-file).

## Prepare the dependency

Use 64-bit Windows 10 or 11 with Windows PowerShell 5.1 or PowerShell 7. Separately
obtain ImageMagick from its [official Windows downloads](https://imagemagick.org/download/#windows-binary-release).
Use **7.1.2-32 or a newer supported 7.x build**, keeping it updated for upstream
security fixes. The maintained test baseline is 7.1.2-32 Q16 x64; the same version's
Q16-HDRI x64 build has separate colour checks. Windows 11 was tested; Windows 10
has not been tested. See the changelog for the exact recorded hosts.

Follow the dependency's installation instructions, or use its portable x64 build,
and make the chosen `magick.exe` available through PATH in the launching session.
In PowerShell, verify the actual application path and its capabilities:

```powershell
Get-Command magick.exe -CommandType Application
magick.exe -version
magick.exe -list format
```

JPEG writing and the decoders for your source formats must be available. HEIC/HEIF
and other codecs vary by build and installed policy. ImageMagick is required even
for video-only or empty folders. This package contains no ImageMagick installer or
binary, and the application does not download or install dependencies during a run.

## Run on an ordinary account

Drag exactly one source folder onto `WinImgNormalizer.bat`. The launcher uses
Windows PowerShell with `-ExecutionPolicy Bypass` for that launched process only;
it does not persist a policy change or override Group Policy. See Microsoft's
[execution-policy scope documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies).
Normalization needs ordinary read access to the source and write access to Pictures
and temporary storage. Do not change machine-wide execution policy, disable
protection software, or weaken ImageMagick policy to run it. If organizational
controls block execution, follow your organization's process.

For the direct positional forms, open PowerShell in the extracted package folder:

```powershell
.\WinImgNormalizer.ps1 'D:\Photos\2024'
.\WinImgNormalizer.ps1 'D:\Photos\2024' 2097152
```

The optional second argument is a positive whole-number byte target. The default is
1,048,576 bytes (1 MiB); a valid JPEG may exceed its target and produce a warning.
Choose a source outside Pictures and outside any parent that contains the output.
Linked source paths are rejected. Live UNC/network-share support is outside scope.

Find the new `<source>_WinImgNormalized_<time>_<runId>` folder under your Windows
Pictures location. It contains mirrored subdirectories, JPEG derivatives, unchanged
video copies and local reports. Originals remain unchanged. Images lose private
metadata after orientation, colour handling and white transparency composition;
videos retain their metadata. JPEGs are lossy derivatives, so keep the originals.
Only the selected image/frame/page is retained; omissions and heuristic duplicate
skips are reported. Review the log and CSV, especially when completion has warnings.

The application processes files on your computer and does not upload them. A
cloud-sync application can synchronize output in Pictures. Logs and reports can
contain private paths; review them before sharing. Exit codes are 0 for success,
1 for setup/run failure, 2 for warnings/partial completion, and 130 for handled
cancellation. Completed results survive cooperative cancellation.

WinImgNormalizer is authored by Janne Vuorela under [LICENSE](LICENSE).
The embedded colour profile's provenance and license are in
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) and [LICENSE-CC0.txt](LICENSE-CC0.txt).
