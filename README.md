# WinImgNormalizer — Non-destructive image normalizer for Windows (PowerShell + ImageMagick)
Minimal, no-frills normalizer I use to prep mixed photo/video folders for archiving. It mirrors the tree into a safe copy, converts images to JPEG ≤ 1 MB, copies videos as-is, and skips simple duplicates. Purpose-built for my workflow, I don’t expect most people to need this. It trades options for speed and repeatability.

**Synopsis**
Recursively mirror a source folder under Pictures, convert images to JPEG (≤ 1 MB, auto-orient, strip metadata, flatten alpha), copy videos unchanged, skip duplicates by (filename + LastWriteTimeUtc), show progress, write a detailed log.
Drag-&-drop workflow: I drop a folder onto the .bat and find the normalized copy in Pictures.

**Requirements**
Windows 10/11
PowerShell (Windows PowerShell 5.1+ or PowerShell 7)
ImageMagick 7+ (magick.exe) in PATH or callable by name
(Optional) The included .bat wrapper for drag-and-drop

**Installation**
Install ImageMagick 7+ for Windows and ensure magick.exe is in PATH (magick -version should work).
Place these two files together (same base name), e.g. in Downloads:
WinImgNormalizer.ps1
WinImgNormalizer.bat (wrapper for double-click + drag-and-drop)
No config files, it just runs.

**Usage**
1) My everyday flow (drag & drop onto .bat)
Drag a folder (with photos/videos) onto WinImgNormalizer.bat.
The normalized copy appears in: %USERPROFILE%\Pictures\<Source>_WinImgNormalized_<yyyyMMdd_HHmmss>
A log file is saved inside the output folder.
2) Command line (positional args, avoids PS 5.1 param-set quirks)
#Whole folder (recursive), default 1 MB cap
.\WinImgNormalizer.ps1 "D:\Photos\2024"
#Whole folder with custom size cap (bytes), for example 2 MB
.\WinImgNormalizer.ps1 "D:\Photos\2024" 2097152

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
%USERPROFILE%\Pictures\<Source>_WinImgNormalized_<timestamp>\
Inside you’ll find the mirrored tree, converted .jpeg images, copied videos, and WinImgNormalizer_<timestamp>.log.

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
“magick not found” -> Install ImageMagick 7+ and ensure magick.exe is in PATH. Test with magick -version.
PS 5.1 “Parameter set cannot be resolved” → Always launch via the provided .bat (positional args) or call the .ps1 with positional arguments only.
Lots of JPEG warnings -> Expected for some camera apps; the script runs ImageMagick with -quiet and continues. Check the log for per-file results.
No outputs -> See the log file in the destination for errors (permissions, unreadable files, ...).
Very noisy images won’t reach ≤ 1 MB → They’ll be saved best-effort and flagged in the log.

**Intent & License**
This is a personal tool for a specific workflow (archiving mixed photo folders for interviews/projects). It’s provided as-is, without warranty. Use at your own risk. Feel free to adapt it, just note it intentionally avoids extra features to keep the workflow fast and predictable.