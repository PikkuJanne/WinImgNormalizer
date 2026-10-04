# WinImgNormalizer — Non-destructive image normalizer for Windows (PowerShell + ImageMagick)
Minimal, no-frills normalizer I use to prep mixed photo/video folders for archiving. It mirrors the tree into a safe copy, converts images to JPEG ≤ 1 MB, copies videos as-is, and skips simple duplicates. Purpose-built for my workflow, I don’t expect most people to need this. It trades options for speed and repeatability.

**Synopsis**
Recursively mirror a source folder under Pictures, convert images to JPEG (≤ 1 MB, auto-orient, strip metadata, flatten alpha), copy videos unchanged, skip duplicates by (filename + LastWriteTimeUtc), show progress, write a detailed log.
Drag-&-drop workflow: I drop a folder onto the .bat and find the normalized copy in Pictures.

**Requirements**
Windows 10/11
PowerShell (Windows PowerShell 5.1+ or PowerShell 7)
ImageMagick 7.1.2-32 or newer supported 7.x build (magick.exe) in PATH
(Optional) The included .bat wrapper for drag-and-drop

**Installation**
Install a supported ImageMagick build for Windows and ensure magick.exe is in PATH
(`magick -version` should work). The minimum reviewed build is **7.1.2-32**, as of
4 October 2026; update it as upstream publishes security fixes. The tool selects
an application executable rather than a PowerShell function or alias, checks its
version and compiled format capabilities, and records its exact path and identity
in the local log. It never installs or updates dependencies automatically.
Place these two files together (same base name), e.g. in Downloads:
WinImgNormalizer.ps1
WinImgNormalizer.bat (wrapper for double-click + drag-and-drop)
No config files, it just runs.

**Usage**
1) My everyday flow (drag & drop onto .bat)
Drag a folder (with photos/videos) onto WinImgNormalizer.bat.
The normalized copy appears in: %USERPROFILE%\Pictures\<Source>_WinImgNormalized_<yyyyMMdd_HHmmss>_<runId>
A log file is saved inside the output folder.
2) Command line (positional args, avoids PS 5.1 param-set quirks)
#Whole folder (recursive), default 1 MB cap
.\WinImgNormalizer.ps1 "D:\Photos\2024"
#Whole folder with custom size cap (bytes), for example 2 MB
.\WinImgNormalizer.ps1 "D:\Photos\2024" 2097152

Supply exactly one source directory and, optionally, a positive whole-number byte
cap from 1 through 9223372036854775807. Small caps remain best-effort and may be
unachievable. Missing folders, files used as folders, non-filesystem sources,
drive-relative paths such as `C:`, and linked source roots or ancestors are rejected
with a setup message. Ordinary relative directory paths are resolved to absolute filesystem
paths; drive and UNC share roots retain their root separator.

The destination must be outside the source tree. Choosing Pictures itself, its
parent, or a drive root that contains Pictures is rejected before enumeration or
any destination write, including probes. Choose a source subfolder instead; the
tool does not silently relocate output. Path comparisons use canonical directories
and case-insensitive whole path segments, so `PhotosBackup` is not inside `Photos`.
Linked destination paths or ancestors are also rejected. Device paths and directory
components with trailing dots/spaces are unsupported.

Traversal skips and logs directory/file reparse points (links and junctions) without
following them. Inaccessible subtrees are reported as an incomplete scan; other
readable media continues. These omissions return exit code 2, including when no
eligible files were readable. Concurrent hostile filesystem changes are outside the
tool's isolation guarantee, though paths are rechecked before reading/writing.

Arguments and ImageMagick are checked before creating Pictures or a run folder.
The destination is checked with a temporary exclusive write probe that is removed,
and a best-effort free-space estimate: copied video bytes, per-image output budgets,
twice the largest readable image input for scratch, and 64 MiB reserve. Each image
budget is the smaller of its byte cap and the larger of 1 MiB or four times its
input size. Missing-decoder files need no conversion space. Insufficient measured
space fails setup; unavailable free-space information (for example, some UNC shares)
is logged as a warning. These checks cannot guarantee later writes succeed.

ImageMagick remains required for **video-only and empty batches**. JPEG write support
is required when there are readable images. An unavailable image decoder is reported
for each affected file without conversion retries; supported images and videos
continue. Such a partial batch returns exit code 2; setup failures return 1.
Compiled capabilities do not override local ImageMagick policy or prove every file
can be decoded. Keep the installed security policy in place.

**What it does**
Non-destructive mirror: exact subfolder structure; image files become .jpeg (same base names).
Images: JPG/JPEG/PNG/BMP/TIF/TIFF/GIF/HEIC/HEIF/WebP → JPEG ≤ 1 MB.
Auto-orient via EXIF
Strip metadata
Convert to sRGB
Flatten transparency to white
Progressive attempt: scale 100->50 % while enforcing jpeg:extent
Videos: mp4/mov/mkv/avi/m4v/wmv/webm/mts/m2ts/3gp/3g2 are copied as-is.
Duplicates: any later file whose (filename lowercase + LastWriteTimeUtc ticks) matches a previously seen one is skipped and logged.
Progress + logs: console progress bar and a timestamped log in the destination.

**Output location**
Default target is the Windows Pictures folder:
%USERPROFILE%\Pictures\<Source>_WinImgNormalized_<timestamp>_<runId>\
Each run has a random 32-character ID and is created exclusively; an existing file
or folder is never reused as a run. Inside you'll find the mirrored tree, converted
.jpeg images and copied videos. Generated files occupy `.WinImgNormalizer`, or the
first free suffix such as `.WinImgNormalizer__2` when a source entry uses that name.
Its `reports` subfolder contains `WinImgNormalizer_<timestamp>.log`; `work` is
reserved for temporary processing. Source directories with these names are mirrored
normally. A log creation failure stops setup and retains the run for diagnosis.

**Batch wrapper (included)**
The repo includes a minimal wrapper so you can drag a folder onto the .bat.
It runs the .ps1 positionally (no named params), which is the safest path on PowerShell 5.1.
Keep the .bat and .ps1 in the same folder and with the same base name.

**Technical details**
ImageMagick is invoked with -auto-orient -strip -colorspace sRGB -sampling-factor 4:2:0 -interlace Line -define jpeg:extent=<bytes>.
Alpha is flattened to white, a retry path omits alpha flags for formats that don’t need them.
JPEG warnings like “Invalid SOS parameters for sequential JPEG” are suppressed (-quiet) and do not abort processing.
Output file timestamps are set to the source file’s timestamps.

**Tweaks (optional)**
Different size cap: pass a second positional argument in bytes (for example 2097152 for 2 MB).
Background color for alpha: change -background white in the script (for example to black).
Dimension ceiling: add an explicit -resize rule (for example -resize "1920x1920>") before jpeg:extent if you want a max edge.

**Troubleshooting**
“magick not found” -> Install ImageMagick 7.1.2-32 or newer supported 7.x and ensure magick.exe is in PATH. Test with magick -version.
PS 5.1 “Parameter set cannot be resolved” → Always launch via the provided .bat (positional args) or call the .ps1 with positional arguments only.
Lots of JPEG warnings -> Expected for some camera apps; the script runs ImageMagick with -quiet and continues. Check the log for per-file results.
No outputs -> See the log file in the destination for errors (permissions, unreadable files, ...).
Very noisy images won’t reach ≤ 1 MB → They’ll be saved best-effort and flagged in the log.

**Intent & License**
This is a personal tool for a specific workflow (archiving mixed photo folders for interviews/projects). It’s provided as-is, without warranty. Use at your own risk. Feel free to adapt it, just note it intentionally avoids extra features to keep the workflow fast and predictable.
