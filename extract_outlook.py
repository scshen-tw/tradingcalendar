#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CB行事曆資料提取器
從 Outlook cbas 資料夾中提取「cb案件整理表」郵件的表格
依據掛牌日、CB競拍截止生成互動式 HTML 行事曆

安裝依賴: pip install pywin32 beautifulsoup4 lxml
"""

import sys
import json
import re
import os
from datetime import datetime

# ========== 依賴檢查 ==========
try:
    import win32com.client
except ImportError:
    print("❌ 請先安裝 pywin32：pip install pywin32")
    sys.exit(1)

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("❌ 請先安裝 beautifulsoup4：pip install beautifulsoup4")
    sys.exit(1)

# ========== 設定區（可依需求修改） ==========
CONFIG = {
    'outlook_folder':  'cbas',          # Outlook 目標資料夾名稱
    'email_subject':   'cb案件整理表',   # 郵件主旨關鍵字
    'output_json':     'events.json',    # 輸出 JSON 檔案
    'output_html':     'calendar.html',  # 輸出行事曆 HTML
}

# 欄位識別關鍵字（按優先順序）
COLUMN_PATTERNS = {
    'method':        ['承銷方式', '方式', '銷售方式'],
    'code':          ['CB代碼', '代號', '債代', '標的代號', '編號', '債券代號'],
    'name':          ['發行標的', '名稱', '債名', '標的名稱', '公司名', '債券名稱'],
    'tcri':          ['tcri', 'TCRI', '評等', 'TCRI評等', '信評'],
    'listing_date':  ['掛牌日', '掛牌', '上市日', '掛牌日期'],
    'auction_end':   ['競拍期間', '詢圈期間', '競拍截止', '拍止', '截止日', '競拍日', '投標截止'],
    'amount':        ['發行金額', '金額', '總額', '發行規模'],
    'conv_price':    ['轉換價', '轉換價格', '換股價', '轉換'],
}

# ========== 工具函數 ==========

def normalize(text):
    """清理文字空白與特殊字元"""
    if text is None:
        return ''
    s = str(text)
    s = re.sub(r'[\r\n\t]+', ' ', s)
    s = s.replace('\u200b', '').replace('\xa0', ' ').replace('\u3000', ' ')
    return s.strip()


def parse_tcri(raw):
    """TCRI 評等解析：有數字→ tcriN，否則→「有擔」"""
    s = normalize(raw)
    if not s or s in ('-', '/', 'N/A', 'NA', '—', ''):
        return ''
    m = re.search(r'\d+', s)
    return f'tcri{m.group()}' if m else '有擔'


def parse_amount(raw):
    """發行金額統一為 NE 格式（N億）"""
    s = normalize(raw)
    if not s:
        return ''
    s = s.replace('億', 'E').replace('亿', 'E')
    # 純數字補 E
    if re.match(r'^\d+(\.\d+)?$', s):
        s += 'E'
    # 確保 E 大寫
    s = re.sub(r'e$', 'E', s)
    return s


def parse_date(raw, ref_year=None):
    """解析日期字串，回傳 YYYY-MM-DD 或 None。
    支援範圍格式（如 4/16~4/20、4/16-4/20），取後段結束日。"""
    s = normalize(raw)
    if not s or s in ('-', '/', 'N/A', 'NA', '—', ''):
        return None
    y = ref_year or datetime.now().year

    # 範圍格式：X/X~X/X 或 X/X-X/X，取後半段
    range_m = re.search(r'[~～—\-]\s*(\d{1,2}[/\.]\d{1,2}|\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})\s*$', s)
    if range_m:
        s = range_m.group(1)

    # YYYY/MM/DD 或 YYYY-MM-DD
    m = re.match(r'^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$', s)
    if m:
        return f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'

    # MM/DD 或 M/D
    m = re.match(r'^(\d{1,2})[/\-.](\d{1,2})$', s)
    if m:
        mo, da = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12 and 1 <= da <= 31:
            return f'{y}-{mo:02d}-{da:02d}'

    # N月N日
    m = re.match(r'^(\d{1,2})月(\d{1,2})日?$', s)
    if m:
        return f'{y}-{int(m.group(1)):02d}-{int(m.group(2)):02d}'

    return None


def build_cell_matrix(table):
    """建立考慮 rowspan/colspan 的二維儲存格矩陣"""
    rows = table.find_all('tr')
    n_rows = len(rows)
    if n_rows == 0:
        return []

    # 估算最大欄數
    n_cols = 0
    for row in rows[:5]:
        cnt = sum(int(c.get('colspan', 1)) for c in row.find_all(['td', 'th']))
        n_cols = max(n_cols, cnt)
    if n_cols == 0:
        n_cols = 30

    matrix = [[None] * n_cols for _ in range(n_rows)]

    for i, row in enumerate(rows):
        cells = row.find_all(['td', 'th'])
        j = 0
        for cell in cells:
            while j < n_cols and matrix[i][j] is not None:
                j += 1
            if j >= n_cols:
                break
            text = normalize(cell.get_text())
            rs = int(cell.get('rowspan', 1))
            cs = int(cell.get('colspan', 1))
            for ri in range(i, min(i + rs, n_rows)):
                for ci in range(j, min(j + cs, n_cols)):
                    matrix[ri][ci] = text
            j += cs

    return matrix


def find_col_index(headers_combined, patterns):
    """在欄位組合標題中尋找匹配欄位的索引"""
    for i, h in enumerate(headers_combined):
        hl = h.lower()
        for p in patterns:
            if p.lower() in hl:
                return i
    return None


def find_cbas_folder(namespace):
    """遞迴搜尋 Outlook 中的目標資料夾"""
    target = CONFIG['outlook_folder']

    def _search(folders):
        for folder in folders:
            if folder.Name == target:
                return folder
            try:
                result = _search(folder.Folders)
                if result:
                    return result
            except Exception:
                pass
        return None

    for store in namespace.Stores:
        try:
            result = _search(store.GetRootFolder().Folders)
            if result:
                return result
        except Exception:
            pass
    return None


def extract_events_from_email(email, ref_year=None):
    """從郵件 HTML 中提取 CB 事件列表"""
    soup = BeautifulSoup(email.HTMLBody, 'lxml')
    tables = soup.find_all('table')
    if not tables:
        print("  ⚠️  郵件中找不到表格")
        return []

    # 選取最多資料列的表格
    best = max(tables, key=lambda t: len(t.find_all('tr')))
    matrix = build_cell_matrix(best)

    if len(matrix) < 2:
        print("  ⚠️  表格列數不足")
        return []

    # 找標題列（含最多欄位關鍵字的列）
    all_kws = [kw for kws in COLUMN_PATTERNS.values() for kw in kws]
    header_end = 0
    max_score = 0
    for i, row in enumerate(matrix[:6]):
        score = sum(1 for cell in row if cell for kw in all_kws if kw.lower() in (cell or '').lower())
        if score > max_score:
            max_score = score
            header_end = i

    # 合併多列標題（每欄取所有標題列的文字）
    n_cols = len(matrix[0])
    combined_headers = []
    for ci in range(n_cols):
        parts = []
        seen = set()
        for ri in range(header_end + 1):
            v = matrix[ri][ci] or ''
            if v and v not in seen:
                parts.append(v)
                seen.add(v)
        combined_headers.append(' '.join(parts))

    print(f"  標題列 (0~{header_end}行)，共 {n_cols} 欄")
    print(f"  全部欄位標題：")
    for ci, h in enumerate(combined_headers):
        if h.strip():
            print(f"    第{ci+1}欄: 「{h}」")

    # 欄位索引對應
    col = {}
    for field, patterns in COLUMN_PATTERNS.items():
        idx = find_col_index(combined_headers, patterns)
        col[field] = idx
        label = combined_headers[idx] if idx is not None else '未找到'
        print(f"  [{field}] → 第{(idx+1) if idx is not None else '?'}欄「{label}」")

    if col.get('listing_date') is None:
        print("  ❌ 找不到掛牌日欄位，請確認欄位關鍵字設定")
        return []

    # 提取資料列
    events = []
    for row in matrix[header_end + 1:]:
        def g(field):
            i = col.get(field)
            return normalize(row[i]) if (i is not None and i < len(row)) else ''

        code = g('code')
        name = g('name')
        if not code or not name:
            continue
        # 略過疑似標題的列
        if any(kw in code for kw in ['代號', '名稱', '項目']):
            continue

        method       = g('method')   # 承銷方式：競拍 / 詢圈
        tcri         = parse_tcri(g('tcri'))
        listing_date = parse_date(g('listing_date'), ref_year)
        auction_end  = parse_date(g('auction_end'),  ref_year)
        amount       = parse_amount(g('amount'))
        conv_price   = normalize(g('conv_price'))

        if not listing_date and not auction_end:
            continue

        # 判斷承銷方式：優先用 method 欄，否則用有無 auction_end 推斷
        is_auction = '競拍' in method if method else bool(auction_end)

        def make_event(date, etype):
            return {
                'date':       date,
                'type':       etype,
                'code':       code,
                'name':       name,
                'tcri':       tcri,
                'amount':     amount,
                'conv_price': conv_price,
            }

        if is_auction:
            # 競拍：截止日 + 掛牌日各一筆
            if auction_end:
                events.append(make_event(auction_end, 'CB競拍截止'))
            if listing_date:
                events.append(make_event(listing_date, 'CB競拍掛牌'))
        else:
            # 詢圈：只有掛牌日
            if listing_date:
                events.append(make_event(listing_date, 'CB詢圈掛牌'))

    return events


def format_display(e):
    """產生行事曆格子中的簡短顯示文字"""
    parts = [e['type'], e['code'], e['name']]
    if e.get('tcri'):
        parts.append(e['tcri'])
    if e.get('amount'):
        parts.append(e['amount'])
    if e.get('conv_price'):
        parts.append(f"轉:{e['conv_price']}")
    return ' '.join(parts)


# ========== HTML 生成 ==========

def generate_html(events, path):
    """將事件資料嵌入 HTML 行事曆並儲存"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, 'calendar.html')

    if not os.path.exists(template_path):
        print(f"  ⚠️  找不到 calendar.html 範本，請確認檔案存在於 {script_dir}")
        return

    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    json_str = json.dumps(events, ensure_ascii=False)
    # 替換嵌入式資料區塊
    html = re.sub(
        r'// __EVENTS_START__.*?// __EVENTS_END__',
        f'// __EVENTS_START__\nconst EVENTS_DATA = {json_str};\n// __EVENTS_END__',
        html,
        flags=re.DOTALL
    )

    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)


# ========== 主程式 ==========

def main():
    print("=" * 50)
    print("  CB行事曆資料提取器")
    print("=" * 50)
    print(f"目標資料夾: {CONFIG['outlook_folder']}")
    print(f"郵件主旨:   {CONFIG['email_subject']}\n")

    # 連線 Outlook
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"❌ 無法連線 Outlook: {e}")
        sys.exit(1)

    # 尋找目標資料夾
    print(f"🔍 搜尋「{CONFIG['outlook_folder']}」資料夾...")
    folder = find_cbas_folder(namespace)
    if not folder:
        print(f"❌ 找不到資料夾「{CONFIG['outlook_folder']}」")
        sys.exit(1)
    print(f"✅ 找到資料夾: {folder.Name}")

    # 尋找最新符合主旨的郵件
    print(f"🔍 搜尋主旨含「{CONFIG['email_subject']}」的郵件...")
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    email = None
    for msg in items:
        try:
            if CONFIG['email_subject'].lower() in (msg.Subject or '').lower():
                email = msg
                break
        except Exception:
            pass

    if not email:
        print(f"❌ 找不到符合主旨的郵件")
        print(f"\n📋 cbas 資料夾內所有郵件主旨（最新20封）：")
        items2 = folder.Items
        items2.Sort("[ReceivedTime]", True)
        count = 0
        for msg in items2:
            try:
                subj = msg.Subject or '（無主旨）'
                recv = str(msg.ReceivedTime)[:10]
                print(f"  [{recv}] {subj}")
                count += 1
                if count >= 20:
                    break
            except Exception:
                pass
        if count == 0:
            print("  （資料夾是空的）")
        sys.exit(1)

    print(f"✅ 找到郵件: {email.Subject}")
    print(f"   收信時間: {email.ReceivedTime}\n")

    try:
        ref_year = email.ReceivedTime.year
    except Exception:
        ref_year = datetime.now().year

    # 提取表格
    print(f"📊 解析表格資料 (參考年份: {ref_year})...")
    events = extract_events_from_email(email, ref_year)

    # 加入顯示文字
    for e in events:
        e['display'] = format_display(e)

    print(f"\nCB 事件共 {len(events)} 筆:")
    for e in sorted(events, key=lambda x: x['date']):
        print(f"  {e['date']}  {e['display']}")

    # 合併股票競拍事件
    print(f"\n📈 讀取股票競拍資料...")
    try:
        from extract_stocks import extract_stock_events
        stock_events = extract_stock_events()
        events = events + stock_events
    except Exception as ex:
        print(f"  ⚠️  股票競拍資料讀取失敗: {ex}")

    print(f"\n合計 {len(events)} 個行事曆事件")

    # 儲存 JSON
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_path  = os.path.join(script_dir, CONFIG['output_json'])
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(events, f, ensure_ascii=False, indent=2)
    print(f"\n💾 已儲存 JSON: {json_path}")

    # 更新 HTML
    html_path = os.path.join(script_dir, CONFIG['output_html'])
    generate_html(events, html_path)
    print(f"🗓️  已更新行事曆: {html_path}")
    print("\n✅ 完成！請用瀏覽器開啟 calendar.html")


if __name__ == '__main__':
    main()
