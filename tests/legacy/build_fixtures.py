#!/usr/bin/env python3
"""Build synthetic M0-T02 inputs in a new, explicitly selected scratch directory.

ImageMagick creates the encoded images; Pillow independently decodes every frame.
No binaries are downloaded, application entry point is run, or Known Folder changed.
"""

from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
from pathlib import Path
import random
import re
import shutil
import struct
import subprocess
import sys
from datetime import datetime, timezone


SCHEMA_VERSION = 1
FIXED_UTC = "2020-02-03T04:05:06Z"
FIXED_SECONDS = int(datetime(2020, 2, 3, 4, 5, 6, tzinfo=timezone.utc).timestamp())
SEED = 4418
MARKERS = [(240, 16, 16), (16, 240, 16), (16, 16, 240), (240, 240, 16)]
LIMIT_SECONDS = 30


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def fixed_times(path: Path) -> None:
    """Set actual Windows creation/write times, then independently read them back."""
    if os.name == "nt":
        from ctypes import wintypes

        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                      ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                      wintypes.HANDLE]
        kernel.CreateFileW.restype = wintypes.HANDLE
        kernel.SetFileTime.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.FILETIME),
                                      ctypes.POINTER(wintypes.FILETIME),
                                      ctypes.POINTER(wintypes.FILETIME)]
        kernel.SetFileTime.restype = wintypes.BOOL
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        kernel.CloseHandle.restype = wintypes.BOOL
        handle = kernel.CreateFileW(str(path), 0x100, 0x7, None, 3, 0x02000000, None)
        if handle == ctypes.c_void_p(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        ticks = (FIXED_SECONDS + 11644473600) * 10000000
        filetime = wintypes.FILETIME(ticks & 0xFFFFFFFF, ticks >> 32)
        try:
            if not kernel.SetFileTime(handle, ctypes.byref(filetime), None,
                                      ctypes.byref(filetime)):
                raise ctypes.WinError(ctypes.get_last_error())
        finally:
            kernel.CloseHandle(handle)
    os.utime(path, (FIXED_SECONDS, FIXED_SECONDS))
    stat = path.stat()
    if int(stat.st_mtime) != FIXED_SECONDS:
        raise RuntimeError("Last-write timestamp was not set exactly")
    if os.name == "nt" and int(stat.st_birthtime) != FIXED_SECONDS:
        raise RuntimeError("Creation timestamp was not set exactly")


def ppm(path: Path, width: int, height: int, mode: str,
        color: tuple[int, int, int] | None = None) -> None:
    rng = random.Random(SEED)
    body = bytearray()
    for y in range(height):
        for x in range(width):
            if color is not None:
                pixel = color
            elif x < 8 and y < 8:
                pixel = MARKERS[0]
            elif x >= width - 8 and y < 8:
                pixel = MARKERS[1]
            elif x < 8 and y >= height - 8:
                pixel = MARKERS[2]
            elif x >= width - 8 and y >= height - 8:
                pixel = MARKERS[3]
            elif mode == "noise":
                pixel = tuple(rng.randrange(256) for _ in range(3))
            elif mode == "lines":
                pixel = (192, 192, 192) if (x + y) % 7 else (32, 32, 32)
            else:
                pixel = (x * 255 // max(1, width - 1),
                         y * 255 // max(1, height - 1), 96)
            body.extend(pixel)
    path.write_bytes(f"P6\n{width} {height}\n255\n".encode("ascii") + body)


def rgba_pam(path: Path) -> None:
    pixels = bytes([255, 0, 0, 0, 0, 255, 0, 128, 0, 0, 255, 255, 255, 0, 0, 255])
    path.write_bytes(b"P7\nWIDTH 4\nHEIGHT 1\nDEPTH 4\nMAXVAL 255\nTUPLTYPE RGB_ALPHA\nENDHDR\n" + pixels)


class Builder:
    def __init__(self, root: Path, magick: Path, Image, pillow_version: str) -> None:
        self.root, self.magick, self.Image = root, magick, Image
        self.pillow_version = pillow_version
        self.recipes = root / ".recipes"
        self.commands: list[dict] = []
        self.files: list[dict] = []
        self.counter = 0
        self.version = self.run(["-version"]).stdout.strip()
        if not re.match(r"Version: ImageMagick 7\.", self.version):
            raise RuntimeError("A real ImageMagick 7 executable is required")
        self.formats = self.run(["-list", "format"]).stdout

    def sanitized(self, text: str) -> str:
        return text.replace(str(self.root), "<fixture-root>").replace(str(self.magick), "<magick>")

    def run(self, args: list[str]) -> subprocess.CompletedProcess:
        result = subprocess.run([str(self.magick), *args], capture_output=True,
                                text=True, encoding="utf-8", errors="replace",
                                timeout=LIMIT_SECONDS, check=False)
        self.commands.append({"argv": ["<magick>", *[self.sanitized(arg) for arg in args]],
                              "exit_code": result.returncode,
                              "stdout": self.sanitized(result.stdout),
                              "stderr": self.sanitized(result.stderr)})
        if result.returncode:
            raise RuntimeError(f"ImageMagick failed (exit {result.returncode}): "
                               + self.sanitized(result.stderr))
        return result

    def source(self, width: int, height: int, mode: str = "gradient",
               color: tuple[int, int, int] | None = None) -> Path:
        self.counter += 1
        path = self.recipes / f"source-{self.counter:03}.ppm"
        ppm(path, width, height, mode, color)
        return path

    def inspect(self, path: Path) -> dict:
        with self.Image.open(path) as image:
            frame_count = getattr(image, "n_frames", 1)
            result = {"format": image.format, "dimensions": list(image.size),
                      "frame_count": frame_count, "mode": image.mode,
                      "exif_orientation": image.getexif().get(274),
                      "icc_profile_present": bool(image.info.get("icc_profile")),
                      "frames": []}
            if frame_count < 1:
                raise RuntimeError("Independent decoder found no frames")
            for index in range(frame_count):
                image.seek(index)
                bounds = list(image.tile[0][1]) if image.tile else None
                image.load()  # Force full decode, including every TIFF/GIF page.
                rgb = image.convert("RGB")
                points = [(min(4, rgb.width - 1), min(4, rgb.height - 1)),
                          (max(0, rgb.width - 5), min(4, rgb.height - 1)),
                          (min(4, rgb.width - 1), max(0, rgb.height - 5)),
                          (max(0, rgb.width - 5), max(0, rgb.height - 5))]
                result["frames"].append({"index": index, "dimensions": list(image.size),
                                         "encoded_tile_bounds": bounds,
                                         "corner_rgb": [list(rgb.getpixel(p)) for p in points]})
            return result

    def record(self, identifier: str, relative: str, recipe: str, expected: dict,
               *, extra: dict | None = None, invalid: bool = False) -> dict:
        path = self.root / relative
        if invalid:
            try:
                self.inspect(path)
            except (OSError, ValueError, SyntaxError) as error:
                observed = {"decode_succeeded": False, "decoder_error_type": type(error).__name__}
            else:
                raise RuntimeError("Intentionally invalid fixture unexpectedly decoded")
        else:
            observed = self.inspect(path)
            for key in ("format", "dimensions", "frame_count", "exif_orientation"):
                if key in expected and observed[key] != expected[key]:
                    raise RuntimeError(f"Independent check failed for {identifier}: {key}")
            if "corner_rgb" in expected:
                actual = observed["frames"][0]["corner_rgb"]
                tolerance = expected.get("pixel_tolerance", 0)
                if any(abs(a - b) > tolerance for pixel_a, pixel_b in
                       zip(actual, expected["corner_rgb"]) for a, b in zip(pixel_a, pixel_b)):
                    raise RuntimeError(f"Independent marker check failed for {identifier}")
        fixed_times(path)
        stat = path.stat()
        entry = {"id": identifier, "path": relative, "recipe": recipe, "seed": SEED,
                 "generating_tool": "ImageMagick 7 / Python standard library",
                 "generating_tool_version": self.version.splitlines()[0],
                 "source": "synthetic recipe in tests/legacy/build_fixtures.py",
                 "license": "repository MIT", "privacy_status": "synthetic",
                 "expected": expected, "observed": observed, "sha256": sha256(path),
                 "length": stat.st_size, "last_write_utc": FIXED_UTC,
                 "creation_utc": FIXED_UTC if os.name == "nt" else None}
        if extra:
            entry.update(extra)
        self.files.append(entry)
        return entry

    def encode(self, identifier: str, relative: str, inputs: list[str], expected: dict,
               recipe: str, options: list[str] | None = None) -> dict:
        """Two fresh encodes must match within this exact tool/codec version."""
        destination = self.root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        suffix = destination.suffix
        first = self.recipes / f"encode-{len(self.files):03}-a{suffix}"
        second = self.recipes / f"encode-{len(self.files):03}-b{suffix}"
        settings = ["+set", "date:create", "+set", "date:modify"]
        if suffix == ".png":
            settings.extend(["-define", "png:exclude-chunk=date,time"])
        if suffix in (".jpg", ".jpeg"):
            settings.extend(["-quality", "95"])
        for target in (first, second):
            self.run([*inputs, *(options or []), *settings, str(target)])
            if not target.is_file() or not target.stat().st_size:
                raise RuntimeError("Required encoding did not produce a nonempty file")
        if sha256(first) != sha256(second):
            raise RuntimeError(f"Encoding {identifier} is nondeterministic in this environment")
        shutil.copyfile(first, destination)
        return self.record(identifier, relative, recipe, expected,
                           extra={"determinism": "two fresh encodes have equal SHA-256"})

    def ordinary(self) -> None:
        for identifier, relative, fmt, size, mode in [
            ("ordinary-jpeg", "corpus/ordinary/landscape.jpg", "JPEG", (64, 48), "gradient"),
            ("ordinary-png", "corpus/ordinary/sub/portrait.png", "PNG", (48, 64), "gradient"),
            ("ordinary-bmp", "corpus/ordinary/sub/lines.bmp", "BMP", (64, 48), "lines"),
            ("ordinary-tiff", "corpus/ordinary/seeded-noise.tiff", "TIFF", (80, 56), "noise"),
        ]:
            source = self.source(*size, mode)
            self.encode(identifier, relative, [str(source)],
                        {"format": fmt, "dimensions": list(size), "frame_count": 1,
                         "corner_rgb": [list(c) for c in MARKERS],
                         "pixel_tolerance": 35 if fmt == "JPEG" else 0},
                        f"P6 PPM {size[0]}x{size[1]} {mode}; four 8x8 RGB corner markers")
        (self.root / "corpus/ordinary/empty-directory").mkdir()
        rng = random.Random(SEED)
        for extension in ("mp4", "mov", "mkv"):
            relative = f"corpus/ordinary/sub/opaque-video.{extension}"
            path = self.root / relative
            path.write_bytes(bytes(rng.randrange(256) for _ in range(4096)))
            fixed_times(path)
            self.files.append({"id": f"video-{extension}", "path": relative,
                               "recipe": "4096 seeded opaque bytes; suffix tests copy only",
                               "seed": SEED, "generating_tool": "Python standard library random.Random",
                               "generating_tool_version": sys.version.split()[0],
                               "source": "synthetic", "license": "repository MIT",
                               "privacy_status": "synthetic", "expected": {"decoded_video": False},
                               "observed": {"length": 4096}, "sha256": sha256(path),
                               "length": 4096, "last_write_utc": FIXED_UTC,
                               "creation_utc": FIXED_UTC if os.name == "nt" else None})
        path = self.root / "corpus/ordinary/ignored.txt"
        path.write_bytes(b"Synthetic unsupported-file control.\n")
        fixed_times(path)
        self.files.append({"id": "unsupported-text", "path": "corpus/ordinary/ignored.txt",
                           "recipe": "literal UTF-8 control text", "seed": None,
                           "generating_tool": "Python standard library",
                           "generating_tool_version": sys.version.split()[0],
                           "source": "synthetic", "license": "repository MIT",
                           "privacy_status": "synthetic", "expected": {"unsupported": True},
                           "observed": {"length": path.stat().st_size}, "sha256": sha256(path),
                           "length": path.stat().st_size, "last_write_utc": FIXED_UTC,
                           "creation_utc": FIXED_UTC if os.name == "nt" else None})

    def collisions_and_duplicates(self) -> None:
        for stem, extension, color in [("photo", "jpg", (224, 32, 32)),
                                       ("photo", "png", (32, 224, 32)),
                                       ("photo", "tiff", (32, 32, 224)),
                                       ("photo_001", "png", (224, 224, 32))]:
            source = self.source(32, 24, color=color)
            fmt = {"jpg": "JPEG", "png": "PNG", "tiff": "TIFF"}[extension]
            self.encode(f"collision-{stem}-{extension}", f"corpus/collision/same-stem/{stem}.{extension}",
                        [str(source)], {"format": fmt, "dimensions": [32, 24], "frame_count": 1,
                                        "corner_rgb": [list(color)] * 4,
                                        "pixel_tolerance": 20 if fmt == "JPEG" else 0},
                        f"P6 PPM 32x24 uniform RGB {color}")
        source = self.source(32, 24, color=(160, 32, 224))
        self.encode("collision-directory-conflict", "corpus/collision/directory-conflict/photo.png",
                    [str(source)], {"format": "PNG", "dimensions": [32, 24], "frame_count": 1,
                                    "corner_rgb": [[160, 32, 224]] * 4},
                    "P6 uniform RGB (160,32,224); intended .jpeg destination already a directory")
        (self.root / "corpus/collision/directory-conflict/photo.jpeg").mkdir()
        for index, color in enumerate(((224, 32, 32), (32, 224, 32))):
            source = self.source(32, 24, color=color)
            self.encode(f"duplicate-{index}", f"corpus/duplicate/{'ab'[index]}/same.png", [str(source)],
                        {"format": "PNG", "dimensions": [32, 24], "frame_count": 1,
                         "corner_rgb": [list(color)] * 4},
                        f"P6 PPM uniform RGB {color}; identical filename and fixed UTC timestamp")
        duplicate_a = self.root / "corpus/duplicate/a/same.png"
        duplicate_b = self.root / "corpus/duplicate/b/same.png"
        if sha256(duplicate_a) == sha256(duplicate_b):
            raise RuntimeError("Duplicate-key controls need different source bytes")

    def frames(self) -> dict:
        first = self.source(64, 48, color=(224, 32, 32))
        second = self.source(64, 48, color=(32, 32, 224))
        entry = self.encode("two-tiff-pages", "corpus/frames/pages.tiff", [str(first), str(second)],
                            {"format": "TIFF", "dimensions": [64, 48], "frame_count": 2},
                            "Two uniform P6 pages: red then blue")
        colors = [frame["corner_rgb"][0] for frame in entry["observed"]["frames"]]
        if colors != [[224, 32, 32], [32, 32, 224]]:
            raise RuntimeError("Independent TIFF pages were not red then blue")
        small = self.source(16, 16, color=(32, 224, 32))
        entry = self.encode("offset-two-frame-gif", "corpus/frames/offset.gif",
                            ["-delay", "12", "(", str(first), "-set", "page", "64x48+0+0", ")",
                             "(", str(small), "-set", "page", "64x48+8+10", ")"],
                            {"format": "GIF", "dimensions": [64, 48], "frame_count": 2},
                            "Red 64x48 canvas then green 16x16 tile at offset +8+10",
                            ["-loop", "0"])
        if entry["observed"]["frames"][1]["encoded_tile_bounds"] != [8, 10, 24, 26]:
            raise RuntimeError("Independent GIF decoder did not confirm the offset tile")
        with self.Image.open(self.root / "corpus/frames/offset.gif") as image:
            image.seek(1)
            image.load()
            if list(image.convert("RGB").getpixel((12, 14))) != [32, 224, 32]:
                raise RuntimeError("Offset GIF tile has unexpected decoded pixels")
        # ImageMagick format listings may include a separate module column.
        webp = re.search(r"(?m)^\s*WEBP\*?\s+(?:\S+\s+)?rw", self.formats) is not None
        from PIL import features
        # Recent Pillow versions expose animation through the WebP module rather
        # than the removed "webp_anim" feature name. Full frame decode follows.
        if features.check("webp"):
            from PIL import _webp
            webp_decoder = hasattr(_webp, "WebPAnimDecoder")
        else:
            webp_decoder = False
        if webp and webp_decoder:
            self.encode("animated-webp", "corpus/frames/animated.webp", ["-delay", "12", str(first), str(second)],
                        {"format": "WEBP", "dimensions": [64, 48], "frame_count": 2},
                        "Two uniform P6 frames red/blue, lossless animated WebP",
                        ["-loop", "0", "-define", "webp:lossless=true"])
            return {"animated_webp": "generated and independently decoded"}
        return {"animated_webp": "deferred: ImageMagick write or Pillow animation capability absent"}

    def alpha_and_orientation(self) -> None:
        source = self.recipes / "alpha.pam"
        rgba_pam(source)
        entry = self.encode("alpha-rgba", "corpus/advanced/alpha/transparency.png", [str(source)],
                            {"format": "PNG", "dimensions": [4, 1], "frame_count": 1},
                            "P7 PAM four RGBA pixels: transparent red, half green, opaque blue/red",
                            ["-define", "png:color-type=6"])
        with self.Image.open(self.root / "corpus/advanced/alpha/transparency.png") as image:
            rgba = image.convert("RGBA")
            actual = [list(rgba.getpixel((x, 0))) for x in range(4)]
            expected = [[255, 0, 0, 0], [0, 255, 0, 128], [0, 0, 255, 255], [255, 0, 0, 255]]
            if actual != expected:
                raise RuntimeError("RGBA fixture lost declared pixel/alpha values")
            entry["expected"]["rgba"] = expected
            entry["observed"]["rgba"] = actual
            entry["expected"]["white_composited_rgb"] = [[255, 255, 255], [127, 255, 127],
                                                         [0, 0, 255], [255, 0, 0]]
        source = self.source(64, 48)
        base = self.encode("orientation-source", "corpus/advanced/orientation/source.jpg", [str(source)],
                           {"format": "JPEG", "dimensions": [64, 48], "frame_count": 1,
                            "corner_rgb": [list(c) for c in MARKERS], "pixel_tolerance": 35},
                           "P6 asymmetric corner-marked JPEG before EXIF orientation tagging")
        original = (self.root / base["path"]).read_bytes()
        for orientation in range(1, 9):
            tiff = b"II" + struct.pack("<HIH", 42, 8, 1)
            tiff += struct.pack("<HHIHHI", 274, 3, 1, orientation, 0, 0)
            payload = b"Exif\x00\x00" + tiff
            app1 = b"\xff\xe1" + struct.pack(">H", len(payload) + 2) + payload
            relative = f"corpus/advanced/orientation/orientation-{orientation}.jpg"
            (self.root / relative).write_bytes(original[:2] + app1 + original[2:])
            self.record(f"exif-orientation-{orientation}", relative,
                        "ImageMagick marker JPEG; standard-library APP1 Exif TIFF IFD orientation tag",
                        {"format": "JPEG", "dimensions": [64, 48], "frame_count": 1,
                         "exif_orientation": orientation, "corner_rgb": [list(c) for c in MARKERS],
                         "pixel_tolerance": 35})

    def invalid_and_names(self) -> None:
        (self.root / "corpus/advanced/invalid").mkdir()
        for identifier, relative, body in [("zero-jpeg", "corpus/advanced/invalid/zero.jpg", b""),
                                           ("header-jpeg", "corpus/advanced/invalid/header.jpg", b"\xff\xd8\xff\xe0\x00\x10JFIF\x00"),
                                           ("invalid-jpeg", "corpus/advanced/invalid/benign.jpg", b"Benign synthetic non-image bytes.\n")]:
            (self.root / relative).write_bytes(body)
            self.record(identifier, relative, "literal benign invalid bytes", {"decode_succeeded": False}, invalid=True)
        source = self.source(32, 24, color=(96, 160, 224))
        for index, name in enumerate(["space name", "[brackets]", "percent%d", "quote's",
                                      "amp&ersand", "(parentheses)", "bang!", "ää-ö-é", "emoji-😀",
                                      "=formula", "+formula", "@formula", "-formula"]):
            self.encode(f"filename-{index}", f"corpus/advanced/filenames/{name}.png", [str(source)],
                        {"format": "PNG", "dimensions": [32, 24], "frame_count": 1,
                         "corner_rgb": [[96, 160, 224]] * 4},
                        "P6 uniform source; encoded under neutral path then copied to legal literal filename")

    def build(self) -> None:
        self.root.mkdir()
        write_json(self.root / ".winimg-fixture-root", {"schema_version": SCHEMA_VERSION,
                                                        "purpose": "synthetic M0-T02 scratch inputs"})
        self.recipes.mkdir()
        try:
            self.ordinary()
            self.collisions_and_duplicates()
            capabilities = self.frames()
            self.alpha_and_orientation()
            self.invalid_and_names()
            source = self.source(64, 48)
            self.encode("weak-source-jpeg", "corpus/weak/source.jpg", [str(source)],
                        {"format": "JPEG", "dimensions": [64, 48], "frame_count": 1,
                         "corner_rgb": [list(c) for c in MARKERS], "pixel_tolerance": 35},
                        "P6 asymmetric corner-marked JPEG; controlled runner supplies weak output")
            manifest = {"schema_version": SCHEMA_VERSION,
                        "generator": "tests/legacy/build_fixtures.py",
                        "generator_sha256": sha256(Path(__file__)), "seed": SEED,
                        "timestamp_utc": FIXED_UTC, "python_version": sys.version.split()[0],
                        "independent_decoder": {"name": "Pillow", "version": self.pillow_version},
                        "imagemagick_version": self.version, "imagemagick_formats": self.formats,
                        "capabilities": {**capabilities,
                                         "heic_heif": "deferred: verified primary/sequence corpus not supplied",
                                         "external_profiles": "deferred: no trusted redistributable profile provenance supplied"},
                        "directories": ["corpus/ordinary/empty-directory",
                                        "corpus/collision/directory-conflict/photo.jpeg"],
                        "groups": {"same_stem_collision": "corpus/collision/same-stem",
                                   "directory_collision": "corpus/collision/directory-conflict"},
                        "planning_mocks": [{"names": ["case.png", "CASE.PNG"],
                                             "reason": "Windows default filesystem cannot store case-only sibling files"},
                                            {"names": ["bad:name.png", "bad?name.png"],
                                             "reason": "Windows-illegal filename characters; mocks only"}],
                        "limitations": ["Fixture checks do not execute or validate WinImgNormalizer.",
                                        "JPEG byte determinism is asserted only within the recorded tool/codec environment.",
                                        "Video suffixes contain opaque bytes; no real video codec claim.",
                                        "No color-profile accuracy or HEIC/HEIF coverage is claimed."],
                        "files": self.files, "commands": self.commands}
            write_json(self.root / "manifest.json", manifest)
            for directory in sorted((p for p in self.root.rglob("*") if p.is_dir()),
                                    key=lambda p: len(p.parts), reverse=True):
                fixed_times(directory)
            fixed_times(self.root)
            print(json.dumps({"result": "generated", "fixtures": len(self.files),
                              "manifest_sha256": sha256(self.root / "manifest.json"),
                              "independent_decoder": f"Pillow {self.pillow_version}"}))
        except Exception:
            write_json(self.root / "generation-failure.json", {"result": "failed",
                                                                "commands": self.commands})
            raise


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True, help="New absent scratch directory")
    parser.add_argument("--magick", type=Path, required=True, help="Existing real ImageMagick 7 executable")
    args = parser.parse_args()
    candidate = args.root.absolute()
    root, magick = candidate.resolve(), args.magick.resolve()
    repository = Path(__file__).resolve().parents[2]
    scratch = repository / ".scratch"
    if root == scratch or scratch not in root.parents:
        parser.error("--root must be a new child of the repository's ignored .scratch directory")
    if root.exists() or root.is_symlink():
        parser.error("--root must be absent; existing files are never overwritten")
    if any(p.is_symlink() or getattr(p, "is_junction", lambda: False)()
           for p in [candidate, *candidate.parents] if p.exists()):
        parser.error("Scratch ancestors must not be reparse points")
    if not root.parent.is_dir():
        parser.error("--root parent must already exist")
    ignored = subprocess.run(["git", "check-ignore", "--quiet", str(root)], cwd=repository,
                             capture_output=True, timeout=LIMIT_SECONDS, check=False)
    if ignored.returncode:
        parser.error("Scratch root must be ignored before generation")
    if not magick.is_file():
        parser.error("--magick must identify an existing executable; no auto-installation")
    try:
        from PIL import Image, __version__ as pillow_version
        Builder(root, magick, Image, pillow_version).build()
    except (ImportError, OSError, RuntimeError, subprocess.TimeoutExpired) as error:
        print(f"Fixture generation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
