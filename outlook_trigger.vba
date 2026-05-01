' ============================================================
' Outlook VBA 巨集 - 信件進來自動觸發 CB 行事曆更新
'
' 安裝方式：
'   1. Outlook → Alt+F11 開啟 VBA 編輯器
'   2. 左側點兩下「ThisOutlookSession」
'   3. 貼上下方程式碼（全部取代）
'   4. 修改 CALENDAR_DIR 為你的實際路徑
'   5. 儲存並關閉 VBA 編輯器
'   6. 重新啟動 Outlook（第一次需要允許執行巨集）
' ============================================================

Private cbas_Folder As Outlook.MAPIFolder
Private WithEvents cbas_Items As Outlook.Items

' ===== 設定區 =====
Const TARGET_FOLDER  As String = "cbas"
Const SUBJECT_KW     As String = "cb案件整理表"
Const CALENDAR_DIR   As String = "D:\VS Code\TradingCalendar"   ' ← 修改為你的路徑
Const CACHE_HTML     As String = "cbas_latest_email.html"
Const CACHE_META     As String = "cbas_latest_email_meta.txt"
Const COOLDOWN_SECS  As Integer = 30   ' 同一封信觸發後，冷卻秒數（避免重複執行）

' ===== 內部變數 =====
Dim lastRunTime As Date

' ===== Outlook 啟動時自動執行 =====
Private Sub Application_Startup()
    Call InitCbasWatcher
End Sub

Sub InitCbasWatcher()
    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    ' 遞迴尋找 cbas 資料夾
    Dim folder As Outlook.MAPIFolder
    Set folder = FindFolder(ns.Folders, TARGET_FOLDER)

    If folder Is Nothing Then
        ' 靜默失敗，不打擾使用者
        Exit Sub
    End If

    Set cbas_Folder = folder
    Set cbas_Items  = folder.Items

    lastRunTime = DateAdd("s", -COOLDOWN_SECS - 1, Now)  ' 確保第一次可立即執行
    Exit Sub

ErrHandler:
    ' 靜默處理，不影響 Outlook 正常使用
End Sub

' ===== 手動匯出最新一封符合主旨的 CBAS 郵件 =====
Sub ExportLatestCbasMailCache()
    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Dim folder As Outlook.MAPIFolder
    Set folder = FindFolder(ns.Folders, TARGET_FOLDER)
    If folder Is Nothing Then Exit Sub

    Dim items As Outlook.Items
    Set items = folder.Items
    items.Sort "[ReceivedTime]", True

    Dim msg As Object
    For Each msg In items
        If msg.Class = olMail Then
            If InStr(1, msg.Subject, SUBJECT_KW, vbTextCompare) > 0 Then
                SaveCbasMailCache msg
                MsgBox "已匯出最新 CBAS 郵件：" & vbCrLf & msg.Subject, vbInformation
                Exit Sub
            End If
        End If
    Next msg

    MsgBox "找不到主旨含「" & SUBJECT_KW & "」的郵件。", vbExclamation
    Exit Sub
ErrHandler:
    MsgBox "匯出 CBAS 郵件失敗：" & Err.Description, vbExclamation
End Sub

' ===== 新信件進來時觸發 =====
Private Sub cbas_Items_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrHandler

    ' 只處理郵件
    If Item.Class <> olMail Then Exit Sub

    ' 確認主旨
    If InStr(1, Item.Subject, SUBJECT_KW, vbTextCompare) = 0 Then Exit Sub

    ' 冷卻檢查（避免同批多封信重複執行）
    If DateDiff("s", lastRunTime, Now) < COOLDOWN_SECS Then Exit Sub
    lastRunTime = Now

    ' 先把信件 HTML 匯出到本機，排程更新時可直接讀檔，不必再連 Outlook COM
    SaveCbasMailCache Item

    ' 延遲 3 秒再執行（讓 Outlook 完成信件處理）
    Application.OnTime Now + TimeValue("00:00:03"), "RunCalendarUpdate"

    Exit Sub
ErrHandler:
End Sub

' ===== 匯出最新 CBAS 郵件 HTML =====
Sub SaveCbasMailCache(ByVal mail As Outlook.MailItem)
    On Error GoTo ErrHandler

    Dim htmlPath As String
    Dim metaPath As String
    htmlPath = CALENDAR_DIR & "\" & CACHE_HTML
    metaPath = CALENDAR_DIR & "\" & CACHE_META

    WriteUtf8Text htmlPath, mail.HTMLBody
    WriteUtf8Text metaPath, mail.Subject & vbCrLf & Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn:ss")

    Exit Sub
ErrHandler:
End Sub

' ===== 用 UTF-8 寫文字檔 =====
Sub WriteUtf8Text(ByVal filePath As String, ByVal text As String)
    On Error GoTo ErrHandler

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close

    Exit Sub
ErrHandler:
    MsgBox "寫入檔案失敗：" & filePath & vbCrLf & Err.Description, vbExclamation
End Sub

' ===== 執行更新腳本 =====
Sub RunCalendarUpdate()
    On Error GoTo ErrHandler

    Dim batPath As String
    batPath = CALENDAR_DIR & "\auto_update.bat"

    ' 確認檔案存在
    If Dir(batPath) = "" Then Exit Sub

    ' 背景執行，不顯示視窗
    Shell "cmd.exe /c """ & batPath & """", vbHide

    Exit Sub
ErrHandler:
End Sub

' ===== 遞迴搜尋資料夾 =====
Function FindFolder(folders As Outlook.Folders, targetName As String) As Outlook.MAPIFolder
    Dim f As Outlook.MAPIFolder
    For Each f In folders
        If f.Name = targetName Then
            Set FindFolder = f
            Exit Function
        End If
        Dim result As Outlook.MAPIFolder
        Set result = FindFolder(f.Folders, targetName)
        If Not result Is Nothing Then
            Set FindFolder = result
            Exit Function
        End If
    Next f
End Function
