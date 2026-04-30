#!/usr/bin/env python3
"""Ensure a log file starts with a UTF-8 BOM for Windows viewers."""

import sys
from pathlib import Path


BOM = b"\xef\xbb\xbf"


def main():
    path = Path(sys.argv[1] if len(sys.argv) > 1 else "update_log.txt")
    data = path.read_bytes() if path.exists() else b""
    if not data.startswith(BOM):
        path.write_bytes(BOM + data)


if __name__ == "__main__":
    main()
