#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import os
import subprocess
from datetime import datetime


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EVENTS_JSON = os.path.join(BASE_DIR, "events.json")
STATUS_LOG = os.path.join(BASE_DIR, "update_status.log")


def event_counts():
    if not os.path.exists(EVENTS_JSON):
        return "events=missing"

    try:
        with open(EVENTS_JSON, encoding="utf-8") as f:
            events = json.load(f)
    except Exception as exc:
        return f"events=unreadable:{exc}"

    cb_count = sum(1 for e in events if str(e.get("type", "")).startswith("CB"))
    stock_count = sum(1 for e in events if str(e.get("type", "")).startswith("股票"))
    return f"events={len(events)} cb={cb_count} stocks={stock_count}"


def git_head():
    try:
        return subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=BASE_DIR,
            text=True,
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
        ).strip()
    except Exception:
        return "unknown"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("task", help="Task name, for example CB or STOCK")
    parser.add_argument("result", help="START, SUCCESS, NO_CHANGE, or ERROR")
    parser.add_argument("message", nargs="*", help="Optional status detail")
    parser.add_argument("--counts", action="store_true", help="Append event counts")
    parser.add_argument("--commit", action="store_true", help="Append current git commit")
    args = parser.parse_args()

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    parts = [f"[{timestamp}]", f"[{args.task}]", args.result]

    if args.message:
        parts.append("- " + " ".join(args.message))
    if args.counts:
        parts.append("| " + event_counts())
    if args.commit:
        parts.append("| commit=" + git_head())

    line = " ".join(parts)
    with open(STATUS_LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")

    print(line)


if __name__ == "__main__":
    main()
