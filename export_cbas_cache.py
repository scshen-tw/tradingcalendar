#!/usr/bin/env python3
"""Export latest CBAS Outlook mail HTML to local cache files."""

import sys
from pathlib import Path

import pythoncom
import win32com.client

BASE = Path(__file__).resolve().parent
TARGET_FOLDER = "cbas"
SUBJECT_KW = "cb案件整理表"
HTML_OUT = BASE / "cbas_latest_email.html"
META_OUT = BASE / "cbas_latest_email_meta.txt"


def find_folder(folders, target):
    for folder in folders:
        if folder.Name == target:
            return folder
        try:
            found = find_folder(folder.Folders, target)
            if found:
                return found
        except Exception:
            pass
    return None


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    pythoncom.CoInitialize()
    outlook = win32com.client.GetActiveObject("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    folder = find_folder(namespace.Folders, TARGET_FOLDER)
    if not folder:
        raise RuntimeError(f"找不到 Outlook 資料夾: {TARGET_FOLDER}")

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    for msg in items:
        try:
            if SUBJECT_KW.lower() not in (msg.Subject or "").lower():
                continue
            HTML_OUT.write_text(msg.HTMLBody or "", encoding="utf-8")
            META_OUT.write_text(
                f"{msg.Subject or ''}\n{str(msg.ReceivedTime)[:19]}\n",
                encoding="utf-8",
            )
            print(f"已匯出: {msg.Subject}")
            print(f"HTML: {HTML_OUT}")
            print(f"META: {META_OUT}")
            return
        except Exception:
            pass

    raise RuntimeError(f"找不到主旨含「{SUBJECT_KW}」的郵件")


if __name__ == "__main__":
    main()
