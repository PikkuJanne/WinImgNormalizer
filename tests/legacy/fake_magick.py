"""Controlled native-process outcomes for instrumented legacy observations.

This is deliberately not an ImageMagick implementation. Its -version output,
call trace and observation labels identify it as a fake.
"""
import json
import os
from pathlib import Path
import sys

from PIL import Image


def main():
    args = sys.argv[1:]
    trace = Path(os.environ["WINIMG_FAKE_TRACE"])
    mode = os.environ["WINIMG_FAKE_MODE"]
    with trace.open("a", encoding="utf-8") as stream:
        stream.write(json.dumps({"mode": mode, "argv": args}) + "\n")
    if args == ["-version"]:
        print("WinImg test fake 1; no real codec behavior")
        return 0
    source = Path(args[1])
    destination = Path(args[-1].removeprefix("JPEG:"))
    destination.parent.mkdir(parents=True, exist_ok=True)
    if mode == "first_failure":
        if source.parent.name == "a":
            print("Controlled first candidate failure", file=sys.stderr)
            return 9
        with Image.open(source) as image:
            image.convert("RGB").save(destination, format="JPEG")
        return 0
    if mode == "weak_zero":
        destination.write_bytes(b"")
    elif mode == "weak_invalid":
        destination.write_bytes(b"synthetic invalid jpeg\n")
    else:
        raise ValueError("Unknown fake mode")
    print("Controlled failure after creating invalid output", file=sys.stderr)
    return 9


if __name__ == "__main__":
    sys.exit(main())
