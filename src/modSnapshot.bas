Attribute VB_Name = "modSnapshot"
Option Explicit

'======================================================
' modSnapshot
' Watchlist全銘柄の市況スナップショットを1回で取得する
'
' 前提：マーケットスピードIIが起動・ログイン済み
'
' 手順：
'   1) Call SnapshotWatchlist   … RSS数式を仕込んで取得完了まで待つ
'   2) Call FreezeSnapshot      … 数式を値に固定してRSS登録を解放する
'      ※ 2をやらずに保存すると、次回オープン時に再登録が走る
'
' 取得項目はすべて楽天証券公式ヘルプ（取得項目・銘柄コード一覧）で
' 確認済みの正式名称：
'   https://marketspeed.jp/ms2_rss/onlinehelp/ohm_002/ohm_002_07.html
'======================================================

Private Const SHEET_SNAP As String = "スナップショット"
Private Const RSS_SAFE_LIMIT As Long = 480   ' 500銘柄上限に対する安全値
Private Const WAIT_TIMEOUT_SEC As Long = 300

'------------------------------------------------------
' ① Watchlist全銘柄にRSS数式を仕込み、取得完了まで待機する
'------------------------------------------------------
Sub SnapshotWatchlist()
    Dim wsW As Worksheet, ws As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long
    Dim n As Long
    Dim res As String

    Set wsW = ThisWorkbook.Sheets("Watchlist")
    lastRow = wsW.Cells(wsW.Rows.Count, 1).End(xlUp).row
    n = lastRow - 1

    If n < 1 Then
        MsgBox "Watchlistに銘柄がありません。", vbExclamation
        Exit Sub
    End If

    If n > RSS_SAFE_LIMIT Then
        MsgBox "Watchlistが " & n & " 銘柄あります。" & vbCrLf & _
               "RSSの同時取得上限（500銘柄）に対して安全値 " & RSS_SAFE_LIMIT & _
               " を超えるため中止します。" & vbCrLf & vbCrLf & _
               "先にWatchlistを減らすか、分割して実行してください。", vbCritical
        Exit Sub
    End If

    ' 検証用シートが残っているとRSS登録を無駄に消費するため削除する
    Call DeleteSheetIfExists("全銘柄候補_TEST")

    Set ws = GetOrCreateSheetLocal(SHEET_SNAP)
    ws.Cells.Clear
    ws.Range("A1:N1").Value = Array( _
        "証券コード", "銘柄名", "現在値", "前日終値", "始値", "高値", "安値", _
        "出来高", "売買代金", "時価総額", "単位株数", "信用貸借区分", "決算発表日", "市場部略称")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    outRow = 2
    For r = 2 To lastRow
        If Trim(CStr(wsW.Cells(r, 1).Value)) <> "" Then
            ws.Cells(outRow, 1).Value = wsW.Cells(r, 1).Value
            ws.Cells(outRow, 2).Value = wsW.Cells(r, 2).Value
            ws.Cells(outRow, 3).Formula = RssF(outRow, "現在値")
            ws.Cells(outRow, 4).Formula = RssF(outRow, "前日終値")
            ws.Cells(outRow, 5).Formula = RssF(outRow, "始値")
            ws.Cells(outRow, 6).Formula = RssF(outRow, "高値")
            ws.Cells(outRow, 7).Formula = RssF(outRow, "安値")
            ws.Cells(outRow, 8).Formula = RssF(outRow, "出来高")
            ws.Cells(outRow, 9).Formula = RssF(outRow, "売買代金")
            ws.Cells(outRow, 10).Formula = RssF(outRow, "時価総額")
            ws.Cells(outRow, 11).Formula = RssF(outRow, "単位株数")
            ws.Cells(outRow, 12).Formula = RssF(outRow, "信用貸借区分")
            ws.Cells(outRow, 13).Formula = RssF(outRow, "決算発表日")
            ws.Cells(outRow, 14).Formula = RssF(outRow, "市場部略称")
            outRow = outRow + 1
        End If
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' [K] 手動計算モード下で入力されたRTD数式は、トピック登録に失敗して
    '     永久に空のままになることがある（2026/8/19に実測）。
    '     待機に入る前に一度だけ完全再計算を強制する。Ctrl+Alt+F9 と等価。
    On Error Resume Next
    Application.CalculateFullRebuild
    On Error GoTo 0
    DoEvents

    res = WaitSnapshotReady(outRow - 1)

    Select Case res
        Case "OK"
            MsgBox "スナップショット取得完了：" & (outRow - 2) & " 銘柄" & vbCrLf & vbCrLf & _
                   "続けて Call FreezeSnapshot を実行してください。" & vbCrLf & _
                   "（値に固定してRSS登録を解放します）", vbInformation
        Case "OVERFLOW"
            MsgBox "「表示銘柄数超過」が発生しました。" & vbCrLf & vbCrLf & _
                   "一度ファイルを保存せずに閉じ、開き直してから再実行してください。", vbExclamation
        Case Else
            MsgBox WAIT_TIMEOUT_SEC & "秒待っても取得が完了しませんでした。" & vbCrLf & _
                   "接続状態を確認してから Ctrl+Alt+F9 で再計算し、" & vbCrLf & _
                   "値が埋まっていれば Call FreezeSnapshot に進んでください。", vbCritical
    End Select
End Sub

'------------------------------------------------------
' ② 数式を値に固定する（RSS登録の解放を兼ねる）
'------------------------------------------------------
Sub FreezeSnapshot()
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SNAP)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "「" & SHEET_SNAP & "」シートがありません。", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then
        MsgBox "データがありません。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, 14)).Value = _
        ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, 14)).Value
    Application.ScreenUpdating = True

    MsgBox "値に固定しました（" & (lastRow - 1) & " 銘柄）。" & vbCrLf & _
           "RSSの登録が解放されました。上書き保存して構いません。", vbInformation
End Sub

'------------------------------------------------------
' 取得完了まで待機
'------------------------------------------------------
Private Function WaitSnapshotReady(lastRow As Long) As String
    Dim ws As Worksheet
    Dim r As Long
    Dim t As Single
    Dim allDone As Boolean, ovf As Boolean
    Dim v As Variant

    Set ws = ThisWorkbook.Sheets(SHEET_SNAP)
    t = Timer

    Do
        DoEvents
        Application.Calculate
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")

        allDone = True
        ovf = False

        For r = 2 To lastRow
            v = ws.Cells(r, 9).Value          ' 売買代金を代表としてチェック
            If v = "表示銘柄数超過" Then
                ovf = True
                Exit For
            ElseIf v = "" Then
                allDone = False
                Exit For
            End If
        Next r

        If ovf Then
            WaitSnapshotReady = "OVERFLOW"
            Exit Function
        End If
        If allDone Then
            WaitSnapshotReady = "OK"
            Exit Function
        End If

        Application.StatusBar = "RSS取得中... 経過 " & Int(Timer - t) & " 秒"
    Loop While Timer - t < WAIT_TIMEOUT_SEC

    Application.StatusBar = False
    WaitSnapshotReady = "TIMEOUT"
End Function

'------------------------------------------------------
' 補助
'------------------------------------------------------
Private Function RssF(r As Long, item As String) As String
    RssF = "=RssMarket(A" & r & ",""" & item & """)"
End Function

Private Sub DeleteSheetIfExists(nm As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

Private Function GetOrCreateSheetLocal(nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nm
    End If
    Set GetOrCreateSheetLocal = ws
End Function
