#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
股票競拍截止事件提取器
從 d:/VS Code/Auction/auction_stocks.json 讀取資料
篩選條件：截止日 >= 今天 - 14天（即截止日未超過兩週前）
"""

import json
import os
import urllib.request
from datetime import datetime, timedelta

STOCKS_JSON   = r"d:\VS Code\Auction\auction_stocks.json"
STOCKS_JSON_URL = "https://raw.githubusercontent.com/scshen-tw/auction-viewer/main/auction_stocks.json"


def parse_date_slash(s):
    """解析 YYYY/MM/DD 格式，回傳 date 物件或 None"""
    if not s:
        return None
    try:
        s = str(s).strip()
        parts = s.replace('-', '/').split('/')
        if len(parts) == 3:
            return datetime(int(parts[0]), int(parts[1]), int(parts[2])).date()
    except Exception:
        pass
    return None


def to_iso(d):
    return d.strftime('%Y-%m-%d') if d else None


def extract_stock_events():
    """提取股票競拍截止事件，回傳 list of event dict（與 CB 行事曆格式一致）"""
    if os.path.exists(STOCKS_JSON):
        with open(STOCKS_JSON, encoding='utf-8') as f:
            stocks = json.load(f)
        print(f"  來源：本機 {STOCKS_JSON}")
    else:
        print(f"  本機檔案不存在，從 GitHub 下載...")
        try:
            with urllib.request.urlopen(STOCKS_JSON_URL, timeout=30) as resp:
                stocks = json.loads(resp.read().decode('utf-8'))
            print(f"  來源：{STOCKS_JSON_URL}（共 {len(stocks)} 筆）")
        except Exception as e:
            print(f"  ⚠️  下載失敗: {e}")
            return []

    today = datetime.now().date()
    cutoff = today - timedelta(days=14)   # 兩週前

    events = []
    for r in stocks:
        deadline_raw = r.get('投標結束日') or ''
        deadline = parse_date_slash(deadline_raw)
        if not deadline:
            continue

        # 篩選：截止日 >= 兩週前（含尚未截止 + 截止日在過去14天內的）
        if deadline < cutoff:
            continue

        code = str(r.get('證券代號') or '').strip()
        name = str(r.get('證券名稱') or '').strip()
        if not code or not name:
            continue

        # 取消競拍的跳過
        cancel = str(r.get('取消競價拍賣(流標或取消)') or '').strip()
        if cancel:
            continue

        # 承銷價：0 或空白視為尚未確定
        price_raw = r.get('實際承銷價格(元)') or ''
        try:
            price_val = float(str(price_raw).replace(',', ''))
            price = f"{price_val:.2f}" if price_val > 0 else ''
        except Exception:
            price = str(price_raw).strip()

        # 掛牌日
        listing_raw = r.get('撥券日期(上市、上櫃日期)') or ''
        listing = parse_date_slash(listing_raw)

        # ── 事件1：競拍截止日 ──
        display_deadline = f"股票競拍截止 {code} {name}"
        events.append({
            'date':       to_iso(deadline),
            'type':       '股票競拍截止',
            'code':       code,
            'name':       name,
            'tcri':       '',
            'amount':     '',
            'conv_price': price,
            'display':    display_deadline,
        })

        # ── 事件2：掛牌日（有掛牌日才加） ──
        if listing:
            price_label = f" 承銷價:{price}" if price else ''
            display_listing = f"股票競拍掛牌 {code} {name}{price_label}"
            events.append({
                'date':       to_iso(listing),
                'type':       '股票競拍掛牌',
                'code':       code,
                'name':       name,
                'tcri':       '',
                'amount':     '',
                'conv_price': price,
                'display':    display_listing,
            })

    cnt_deadline = sum(1 for e in events if e['type'] == '股票競拍截止')
    cnt_listing  = sum(1 for e in events if e['type'] == '股票競拍掛牌')
    print(f"  股票競拍截止：{cnt_deadline} 筆 / 股票競拍掛牌：{cnt_listing} 筆（截止日 >= {cutoff}）")
    return events


if __name__ == '__main__':
    events = extract_stock_events()
    for e in sorted(events, key=lambda x: x['date']):
        print(f"  {e['date']}  {e['display']}")
