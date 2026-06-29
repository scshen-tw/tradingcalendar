# TradingCalendar

CB 可轉債與股票競拍行事曆自動更新工具。

本專案會從 Outlook 的 CBAS 郵件取得可轉債資料，合併 `auction-viewer` 產生的股票競拍 JSON，輸出 `events.json` 與內嵌資料的 `calendar.html`，再推送到 GitHub Pages。

## 目前更新流程

### CB 可轉債資料

目前的主要方式不是讓排程直接控制 Outlook COM，而是先由 Outlook VBA 在使用者登入的 Outlook session 內匯出郵件快取：

1. Outlook 內的 `outlook_trigger.vba` 安裝在 `ThisOutlookSession`。
2. Outlook 啟動時執行 `InitCbasWatcher`，監看名稱為 `cbas` 的資料夾。
3. 當新郵件進入且主旨包含 `cb案件整理表` 時，VBA 會寫出：
   - `cbas_latest_email.html`
   - `cbas_latest_email_meta.txt`
4. VBA 接著以背景方式執行 `auto_update.bat`。
5. `extract_outlook.py` 優先讀取上述本機快取並解析 HTML 表格。
6. 如果快取不存在或主旨不符，才 fallback 到 Outlook COM；COM 仍失敗時，沿用既有 `events.json` 裡的 CB 事件。

主旨只比對 `cb案件整理表`，後面的日期例如 `0430`、`0501` 不需要寫死。

### 股票競拍資料

股票資料由另一個專案產生：

- 本機優先讀取：`D:\vscode\Auction\auction_stocks.json`
- GitHub Actions 環境改讀：`https://raw.githubusercontent.com/scshen-tw/auction-viewer/main/auction_stocks.json`

`extract_stocks.py` 會略過已取消競拍資料，並只保留投標結束日不早於今天前 14 天的項目。

## 主要檔案

- `outlook_trigger.vba`：Outlook 巨集，負責監看新信、匯出 HTML 快取、觸發更新。
- `extract_outlook.py`：CB 主程式，優先讀本機郵件快取，fallback 到 Outlook COM，最後合併股票事件。
- `extract_stocks.py`：股票競拍事件擷取模組。
- `update_stocks.py`：只更新股票事件，不碰 Outlook。
- `auto_update.bat`：本機完整更新，會更新 CB 與股票、commit、push。
- `update_stocks.bat`：本機股票獨立更新，會 commit、push。
- `export_cbas_cache.py`：手動備援工具，可在 Outlook 開啟時匯出最新 CBAS 郵件快取。
- `ensure_utf8_log.py`：確保 log 是 UTF-8，避免中文亂碼。
- `log_update_status.py`：寫入 `update_status.log`，方便確認排程結果。

## 重要快取與 Log

以下檔案是本機運作狀態，不進 git：

- `cbas_latest_email.html`
- `cbas_latest_email_meta.txt`
- `update_log.txt`
- `update_status.log`
- `export_cbas_cache.log`

其中 `cbas_latest_email.html` 與 `cbas_latest_email_meta.txt` 是排程讀取 Outlook 郵件表格的關鍵快取，不是測試垃圾檔。除非要重建快取，不要手動刪除。

## 手動操作

手動匯出最新 CBAS 郵件快取：

```powershell
python export_cbas_cache.py
```

手動完整更新：

```powershell
.\auto_update.bat
```

只更新股票：

```powershell
.\update_stocks.bat
```

## Git 與 Push

本機排程使用 SSH key push，避免排程卡在 GitHub 帳密或互動式登入。

- TradingCalendar remote：`git@github.com:scshen-tw/tradingcalendar.git`
- SSH key：`C:\Users\User\.ssh\id_ed25519_tradingcalendar`

`auto_update.bat` 與 `update_stocks.bat` 都會設定：

```bat
GIT_TERMINAL_PROMPT=0
GCM_INTERACTIVE=never
GIT_SSH_COMMAND=ssh -i C:/Users/User/.ssh/id_ed25519_tradingcalendar -o IdentitiesOnly=yes -o StrictHostKeyChecking=accept-new
```

## 排程注意事項

- Outlook 必須開著且巨集可執行，VBA 才能在新信進來時更新快取。
- 排程可以在鎖定畫面下執行，但不要在排程時間附近登出、切換使用者，或讓 Outlook 被關掉。
- 如果當天沒有新信進來，但要強制刷新，可在 Outlook VBA 手動執行 `ExportLatestCbasMailCache`，或在 Outlook 開啟時執行 `python export_cbas_cache.py`。
- 手動執行 VBA 匯出會跳成功提示窗；自動新信觸發不會每天跳成功提示窗。

## GitHub Actions

`.github/workflows/update.yml` 只做股票資料更新，不讀 Outlook。

Outlook 郵件資料只能在本機透過 Outlook/VBA 快取流程更新。
