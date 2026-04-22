#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
股票競拍行事曆獨立更新器
不依賴 Outlook，直接從 auction_stocks.json 更新 calendar.html 和 events.json
保留現有 CB 事件，只重新計算股票競拍部分。

用法：
  python update_stocks.py
"""

import json
import re
import os
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
EVENTS_JSON  = os.path.join(SCRIPT_DIR, 'events.json')
CALENDAR_HTML = os.path.join(SCRIPT_DIR, 'calendar.html')

STOCK_TYPES = {'股票競拍截止', '股票競拍掛牌'}


def load_existing_cb_events():
    """讀取現有 events.json，只保留 CB 類型事件"""
    if not os.path.exists(EVENTS_JSON):
        return []
    with open(EVENTS_JSON, encoding='utf-8') as f:
        events = json.load(f)
    cb = [e for e in events if e.get('type') not in STOCK_TYPES]
    print(f"  保留既有 CB 事件：{len(cb)} 筆")
    return cb


def update_html(events):
    if not os.path.exists(CALENDAR_HTML):
        print(f"  ⚠️  找不到 {CALENDAR_HTML}")
        return
    with open(CALENDAR_HTML, 'r', encoding='utf-8') as f:
        html = f.read()
    json_str = json.dumps(events, ensure_ascii=False)
    html = re.sub(
        r'// __EVENTS_START__.*?// __EVENTS_END__',
        f'// __EVENTS_START__\nconst EVENTS_DATA = {json_str};\n// __EVENTS_END__',
        html, flags=re.DOTALL
    )
    with open(CALENDAR_HTML, 'w', encoding='utf-8') as f:
        f.write(html)


def main():
    print("=" * 50)
    print("  股票競拍行事曆獨立更新器")
    print("=" * 50)

    cb_events = load_existing_cb_events()

    print("\n📈 讀取股票競拍資料...")
    try:
        from extract_stocks import extract_stock_events
        stock_events = extract_stock_events()
    except Exception as ex:
        print(f"  ❌ 股票資料讀取失敗: {ex}")
        sys.exit(1)

    all_events = cb_events + stock_events
    print(f"\n合計 {len(all_events)} 個行事曆事件（CB: {len(cb_events)}, 股票: {len(stock_events)}）")

    with open(EVENTS_JSON, 'w', encoding='utf-8') as f:
        json.dump(all_events, f, ensure_ascii=False, indent=2)
    print(f"💾 已儲存 {EVENTS_JSON}")

    update_html(all_events)
    print(f"🗓️  已更新 {CALENDAR_HTML}")
    print("\n✅ 完成！")


if __name__ == '__main__':
    main()
