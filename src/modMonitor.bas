Attribute VB_Name = "modMonitor"
Option Explicit

' ============================================================
' modMonitor
' Q6/Q7: ギャップ率判定　Q19: タイマー定期実行　Q13: 通知
' Q20: シグナルログ記録　Q12: 1日1トレード・複数候補時は
' ギャップ率最大を優先
' ------------------------------------------------------------
' 2026/8/13 改訂
'   [D] MsgBox表示中にOnTimeが失われて連鎖が死ぬのを防ぐため、
'       「次回予約 → 判定 → 通知」の順に並べ替えた。
'       MsgBoxを出す前に必ず次回が予約済みになっている。
'   [D] modWatchdog による自動復旧に対応（Rearm系を公開）
'   [C] 直近シグナルの内容を Public 変数で公開。
'       RecordEntry がシグナル検知時刻と気配値をそのまま引き継ぐ。
' 2026/8/14 改訂
'   [E] StartGapMonitoring で G列（ギャップ率）もクリアする。
'       現在値が取得できない銘柄（8/14の5032など）に前日の数値が
'       居座り、当日の有効値と区別がつかなくなるため。
'       セッション中に値が取れなくなった銘柄も都度クリアする。
'   [F] Watchdog の復旧判定を監視開始時刻以降に限定する。
'       開始前は判定を実行せず即抜けするためハートビートが0のままで、
'       「31秒停止」と誤検知して無意味な復旧ログが出ていた。
' 2026/8/17 改訂
'   [G] [F] のガードをすり抜ける3秒の窓を塞ぐ。
'       mLastCheckTime = 0（本体が一度も回っていない）を無条件に
'       異常扱いしていたため、監視開始時刻を跨いだ直後に Watchdog が
'       先に発火すると必ず誤検知になっていた。
'       「一度も回っていない」ではなく「起点からの経過秒」で測る。
'       起点は 監視開始時刻 と 監視開始ボタンを押した時刻 の遅いほう。
'       （開始時刻を過ぎてからボタンを押した場合に即誤検知するのを防ぐ）
' 2026/8/19 改訂
'   [H] セッション終了時に Log へ「日次サマリ」を1行残す。
'       従来はシグナルが出なかった日にワークブック側の痕跡がゼロで、
'       「監視して出なかった」のか「そもそも起動していない」のかを
'       後から区別できなかった（8/17がまさにこの状態）。
'       あわせてセッション中のギャップ率最大値を保持する。G列は毎
'       サイクル上書きされるため、終了時点の値しか残らないため。
' ============================================================

' [C] 直近シグナル（modPosition.RecordEntry が参照する）
Public LastSignalValid As Boolean
Public LastSignalCode As String
Public LastSignalName As String
Public LastSignalGapPct As Double
Public LastSignalAsk As Double
Public LastSignalTime As Date

Private mIsRunning As Boolean
Private mNextRunTime As Date
Private mSignalFoundToday As Boolean   ' Q12: 1日1トレードのみ
Private mLastCheckTime As Date         ' ハートビート：最後に判定が回った時刻
Private mArmedTime As Date             ' [G] 監視開始ボタンを押した時刻
Private mErrCount As Long              ' 連続エラー回数

' [H] セッション中のギャップ率最大値（G列の上書きでは残らないため別途保持）
Private mMaxGapPct As Double
Private mMaxGapCode As String
Private mMaxGapName As String
Private mSummaryWritten As Boolean     ' 日次サマリの二重記録を防ぐ

' エラーがこの回数連続したら監視を停止する
Private Const MAX_CONSECUTIVE_ERRORS As Long = 5

' ハートビートがこの秒数以上止まっていたら連鎖切れとみなす
Private Const STALE_LIMIT_SEC As Long = 30

' ------------------------------------------------------------
' 監視開始（操作パネルのボタンから呼び出す）
' ------------------------------------------------------------
Sub StartGapMonitoring()
    LoadConfig
    mIsRunning = True
    mSignalFoundToday = False
    mLastCheckTime = 0
    mArmedTime = Now                   ' [G] 誤検知判定の起点に使う
    mErrCount = 0

    ' [H] 最大値トラッカーの初期化。全銘柄がマイナスの日でも記録できるよう
    '     初期値は0ではなく-999にする。
    mMaxGapPct = -999
    mMaxGapCode = ""
    mMaxGapName = ""
    mSummaryWritten = False

    LastSignalValid = False
    LastSignalCode = ""
    LastSignalName = ""
    LastSignalGapPct = 0
    LastSignalAsk = 0

    ThisWorkbook.Worksheets(SHEET_WATCHLIST).Range("H:H").ClearContents
    ThisWorkbook.Worksheets(SHEET_WATCHLIST).Range("I:I").ClearContents
    ThisWorkbook.Worksheets(SHEET_WATCHLIST).Cells(1, COL_SIGNAL).Value = "シグナル状態"
    ThisWorkbook.Worksheets(SHEET_WATCHLIST).Cells(1, COL_DETECTTIME).Value = "検知時刻"

    ' [E] 前日のギャップ率が残っていると、当日値が取れない銘柄で
    '     前日の数値をそのまま読んでしまうためG列も消す
    With ThisWorkbook.Worksheets(SHEET_WATCHLIST)
        .Range(.Cells(WATCHLIST_START_ROW, COL_GAPPCT), _
               .Cells(.Rows.Count, COL_GAPPCT)).ClearContents

        ' 前営業日のシグナル行ハイライトが残らないよう背景色もクリアする
        .Range(.Cells(WATCHLIST_START_ROW, COL_CODE), _
               .Cells(.Rows.Count, COL_GAPPCT)).Interior.ColorIndex = xlColorIndexNone
    End With

    Application.StatusBar = "ギャップ監視: 稼働中（" & Format(MonitorStartTime, "hh:mm") & _
        " ～ " & Format(MonitorEndTime, "hh:mm") & "）"

    ScheduleNextCheck
    modWatchdog.StartWatchdog
End Sub

' ------------------------------------------------------------
' 監視停止（手動停止用）
' ------------------------------------------------------------
Sub StopGapMonitoring()
    mIsRunning = False
    CancelPendingCheck

    ' [H] 判定が1回でも回っていればセッションの記録を残す
    WriteSessionSummaryLog

    If mLastCheckTime > 0 Then
        Application.StatusBar = "ギャップ監視: 停止しました（最終チェック " & _
            Format(mLastCheckTime, "hh:mm:ss") & "）"
    Else
        Application.StatusBar = "ギャップ監視: 停止しました"
    End If
End Sub

' ------------------------------------------------------------
' [H] セッション終了時の記録。ノーシグナル日も母集団の情報なので、
'     STATE.md ではなくワークブック側に痕跡を残す。
' ------------------------------------------------------------
Private Sub WriteSessionSummaryLog()
    If mSummaryWritten Then Exit Sub
    If mLastCheckTime = 0 Then Exit Sub   ' 判定が一度も回っていない日は記録しない
    mSummaryWritten = True

    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "日次サマリ"

    If mMaxGapCode <> "" Then
        ws.Cells(nextRow, 3).Value = mMaxGapCode
        ws.Cells(nextRow, 4).Value = mMaxGapName
        ws.Cells(nextRow, 5).Value = Round(mMaxGapPct, 2)
    End If

    ws.Cells(nextRow, 8).Value = IIf(mSignalFoundToday, "シグナルあり", "シグナルなし")

    Dim memo As String
    memo = "監視 " & Format(MonitorStartTime, "hh:mm") & "-" & _
           Format(MonitorEndTime, "hh:mm") & "／最終チェック " & _
           Format(mLastCheckTime, "hh:mm:ss") & _
           "／E列はセッション中の最大ギャップ率（G列の終了時点値ではない）"
    If mErrCount > 0 Then
        memo = memo & "／連続エラー" & mErrCount & "回で停止"
    End If
    ws.Cells(nextRow, 10).Value = memo

    On Error GoTo 0
End Sub

Private Sub CancelPendingCheck()
    On Error Resume Next
    Application.OnTime EarliestTime:=mNextRunTime, Procedure:="CheckGapSignals", Schedule:=False
    On Error GoTo 0
End Sub

' ------------------------------------------------------------
' タイマーで呼ばれる本体
' ------------------------------------------------------------
Sub CheckGapSignals()
    If Not mIsRunning Then Exit Sub

    On Error GoTo ErrHandler

    Dim nowT As Date
    nowT = TimeValue(Now)

    ' 監視終了時刻を過ぎたら自動停止
    If nowT > MonitorEndTime Then
        StopGapMonitoring
        Exit Sub
    End If

    ' [D] 判定の前に次回を予約する。
    ' この後にMsgBoxが出て操作が止まっても、予約自体は既に済んでいる。
    ScheduleNextCheck

    ' 監視開始前ならまだ判定しない
    If nowT < MonitorStartTime Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_WATCHLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_CODE).End(xlUp).row

    Dim bestRow As Long: bestRow = 0
    Dim bestGap As Double: bestGap = 0

    Dim r As Long
    For r = WATCHLIST_START_ROW To lastRow
        Dim code As String
        code = Trim(ws.Cells(r, COL_CODE).Value)
        If code = "" Then GoTo ContinueLoop

        Dim prevClose As Double, current As Double
        prevClose = SafeNum(ws.Cells(r, COL_PREVCLOSE).Value)
        current = SafeNum(ws.Cells(r, COL_CURRENT).Value)

        ' [E] 値が取れない銘柄は古い数値を残さず空にする
        If prevClose <= 0 Or current <= 0 Then
            If ws.Cells(r, COL_GAPPCT).Value <> "" Then ws.Cells(r, COL_GAPPCT).ClearContents
            GoTo ContinueLoop
        End If

        Dim gapPct As Double
        gapPct = (current - prevClose) / prevClose * 100

        ws.Cells(r, COL_GAPPCT).Value = Round(gapPct, 2)

        ' [H] セッション中の最大値を更新する
        If gapPct > mMaxGapPct Then
            mMaxGapPct = gapPct
            mMaxGapCode = code
            mMaxGapName = CStr(ws.Cells(r, COL_NAME).Value)
        End If

        ' Q7: ギャップ率閾値以上のみ候補（Q23: プラス方向の順張りのみ）
        If gapPct >= GapThresholdPct Then
            If gapPct > bestGap Then
                bestGap = gapPct
                bestRow = r
            End If
        End If

ContinueLoop:
    Next r

    ' ハートビート：ここまで到達すれば1サイクル完走
    mLastCheckTime = Now
    mErrCount = 0
    UpdateStatusBar

    ' Q12: ギャップ率最大の1件のみ・1日1回だけ発火
    ' [D] MsgBoxを出す前にフラグを立てる（ダイアログ滞留中の二重発火防止）
    If bestRow > 0 And Not mSignalFoundToday Then
        mSignalFoundToday = True
        FireSignal ws, bestRow, bestGap
    End If

    Exit Sub

ErrHandler:
    ' 一時的なエラーで連鎖が永久に切れるのを防ぐ。
    ' 次回予約は本体先頭で済んでいるため、ここでは再予約しない。
    mErrCount = mErrCount + 1
    Application.StatusBar = "ギャップ監視: エラー" & mErrCount & "回目 " & _
        Format(Now, "hh:mm:ss") & " [" & Err.Number & "] " & Err.Description
    Err.Clear

    If mErrCount >= MAX_CONSECUTIVE_ERRORS Then
        StopGapMonitoring
        MsgBox "ギャップ監視でエラーが" & MAX_CONSECUTIVE_ERRORS & "回連続したため監視を停止しました。" & vbCrLf & _
               "RSS接続とWatchlistの数式を確認してください。", vbCritical, "監視停止"
    End If
End Sub

' ------------------------------------------------------------
' ステータスバー更新（稼働状況＋最終チェック時刻）
' ------------------------------------------------------------
Private Sub UpdateStatusBar()
    Application.StatusBar = "ギャップ監視: 稼働中（" & Format(MonitorStartTime, "hh:mm") & _
        " ～ " & Format(MonitorEndTime, "hh:mm") & "）" & _
        "  最終チェック " & Format(mLastCheckTime, "hh:mm:ss") & _
        IIf(mSignalFoundToday, "  ★本日シグナル確定済", "")
End Sub

' ------------------------------------------------------------
' 次回のタイマーを予約する（Q19: 3～5秒間隔）
' ------------------------------------------------------------
Private Sub ScheduleNextCheck()
    If Not mIsRunning Then Exit Sub
    mNextRunTime = Now + TimeSerial(0, 0, CheckIntervalSec)
    Application.OnTime EarliestTime:=mNextRunTime, Procedure:="CheckGapSignals"
End Sub

' ------------------------------------------------------------
' [D] 監視役（modWatchdog）から呼ばれる自己診断・再起動
' ------------------------------------------------------------
Public Function GapChainIsRunning() As Boolean
    GapChainIsRunning = mIsRunning
End Function

Public Sub WatchdogRearmGap()
    If Not mIsRunning Then Exit Sub

    If TimeValue(Now) > MonitorEndTime Then
        StopGapMonitoring
        Exit Sub
    End If

    ' [F] 監視開始前は本体が判定せず即抜けするため、ハートビートは
    '     更新されない。ここで復旧判定すると必ず誤検知になる。
    If TimeValue(Now) < MonitorStartTime Then Exit Sub

    Dim staleSec As Double
    If mLastCheckTime = 0 Then
        ' [G] まだ本体が一度も完走していない状態。
        '     これ自体は異常ではない（監視開始時刻の直後は必ずこうなる）。
        '     起点＝「監視開始時刻」と「開始ボタンを押した時刻」の遅いほう
        '     からの経過秒で測り、それでも30秒超えなら本当に死んでいる。
        Dim baseT As Date
        baseT = Date + MonitorStartTime
        If mArmedTime > baseT Then baseT = mArmedTime
        staleSec = (Now - baseT) * 86400#
    Else
        staleSec = (Now - mLastCheckTime) * 86400#
    End If

    If staleSec > STALE_LIMIT_SEC Then
        CancelPendingCheck          ' 死んでいる可能性のある予約を消してから
        ScheduleNextCheck           ' 張り直す（二重連鎖の防止）
        LogWatchdogRecovery "ギャップ監視", staleSec
    End If
End Sub

Private Sub LogWatchdogRecovery(chainName As String, staleSec As Double)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "復旧"
    ws.Cells(nextRow, 10).Value = chainName & "のタイマー連鎖が" & Format(staleSec, "0") & _
        "秒停止していたため再起動しました"
    On Error GoTo 0
End Sub

' ------------------------------------------------------------
' シグナル発火：ハイライト・通知・ログ記録（Q13/Q20）
' ------------------------------------------------------------
Private Sub FireSignal(ws As Worksheet, r As Long, gapPct As Double)
    Dim code As String, nm As String
    code = ws.Cells(r, COL_CODE).Value
    nm = ws.Cells(r, COL_NAME).Value

    ws.Cells(r, COL_SIGNAL).Value = "選定（本日のエントリー候補）"
    ws.Cells(r, COL_DETECTTIME).Value = Now

    ' 視覚通知：行をハイライト
    ws.Range(ws.Cells(r, COL_CODE), ws.Cells(r, COL_GAPPCT)).Interior.Color = RGB(255, 199, 206)

    ' 気配値ベースのシミュレーションエントリー価格（Q22）
    Dim askPrice As Double
    askPrice = SafeNum(ws.Cells(r, COL_ASK).Value)
    If askPrice <= 0 Then askPrice = SafeNum(ws.Cells(r, COL_CURRENT).Value)

    Dim bufferedPrice As Double
    bufferedPrice = askPrice * 1.002   ' Q22: 保守的バッファ+0.2%

    ' [C] 直近シグナルを保持。RecordEntry がこの時刻・価格を引き継ぐ。
    LastSignalValid = True
    LastSignalCode = code
    LastSignalName = nm
    LastSignalGapPct = gapPct
    LastSignalAsk = askPrice
    LastSignalTime = Now

    ' ログ記録（Q20）
    WriteSignalLog code, nm, gapPct, askPrice, bufferedPrice

    ' 推奨株数（Q11）
    Dim recShares As Long
    recShares = modPosition.CalcPositionSize(askPrice)

    Dim sharesText As String
    If recShares > 0 Then
        sharesText = Format(recShares, "#,##0") & " 株" & _
                     "（概算約定代金 " & Format(askPrice * recShares, "#,##0") & " 円）"
    Else
        sharesText = "0 株 ※リスク率" & RiskPct & "%では100株未満。" & _
                     "100株なら実リスク約" & _
                     Format((100 * askPrice * Abs(StopLossPct) / 100) / AccountBalance * 100, "0.00") & "%"
    End If

    If SoundOn Then
        modNotify.PlayAlertSound 2
    End If

    If PopupOn Then
        MsgBox "【シグナル検知】" & Format(LastSignalTime, "hh:mm:ss") & vbCrLf & _
               "銘柄コード: " & code & " " & nm & vbCrLf & _
               "ギャップ率: " & Format(gapPct, "0.00") & "%" & vbCrLf & _
               "気配値(Ask): " & Format(askPrice, "0.0") & vbCrLf & _
               "推奨株数: " & sharesText & vbCrLf & _
               "損切り: " & Format(askPrice * (1 + StopLossPct / 100), "0.0") & _
               " ／ 利確: " & Format(askPrice * (1 + TakeProfitPct / 100), "0.0") & vbCrLf & vbCrLf & _
               "手動でマーケットスピードIIの取引画面から成行発注してください。" & vbCrLf & _
               "発注できたら操作パネルの「エントリー記録」を押してください。", _
               vbInformation, "デイトレ支援ツール"
    End If
End Sub

Private Function SafeNum(v As Variant) As Double
    If IsNumeric(v) Then
        SafeNum = CDbl(v)
    Else
        SafeNum = 0
    End If
End Function

' ------------------------------------------------------------
' Log シートへの追記（Q20）
' ------------------------------------------------------------
Private Sub WriteSignalLog(code As String, nm As String, gapPct As Double, _
                            askPrice As Double, bufferedPrice As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1

    If nextRow = 2 And ws.Cells(1, 1).Value = "" Then
        ws.Range("A1:J1").Value = Array("日時", "種別", "証券コード", "銘柄名", _
            "ギャップ率(%)", "気配値ベース価格", "保守的バッファ込み価格", _
            "エントリー有無", "決済理由", "備考")
    End If

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "シグナル"
    ws.Cells(nextRow, 3).Value = code
    ws.Cells(nextRow, 4).Value = nm
    ws.Cells(nextRow, 5).Value = Round(gapPct, 2)
    ws.Cells(nextRow, 6).Value = Round(askPrice, 1)
    ws.Cells(nextRow, 7).Value = Round(bufferedPrice, 1)
    ws.Cells(nextRow, 8).Value = "未定（Position記録待ち）"
End Sub

' ------------------------------------------------------------
' テスト専用：監視時間帯を無視して今すぐ1回だけ判定する
' ※ Logに偽シグナルが混入するため通常は使用しない
' ------------------------------------------------------------
Sub TestCheckGapSignalsNow()
    LoadConfig

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_WATCHLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_CODE).End(xlUp).row

    Dim bestRow As Long: bestRow = 0
    Dim bestGap As Double: bestGap = 0
    Dim checkedCount As Long: checkedCount = 0

    Dim r As Long
    For r = WATCHLIST_START_ROW To lastRow
        Dim code As String
        code = Trim(ws.Cells(r, COL_CODE).Value)
        If code = "" Then GoTo ContinueLoop

        Dim prevClose As Double, current As Double
        prevClose = SafeNum(ws.Cells(r, COL_PREVCLOSE).Value)
        current = SafeNum(ws.Cells(r, COL_CURRENT).Value)

        If prevClose <= 0 Or current <= 0 Then GoTo ContinueLoop

        checkedCount = checkedCount + 1

        Dim gapPct As Double
        gapPct = (current - prevClose) / prevClose * 100

        ws.Cells(r, COL_GAPPCT).Value = Round(gapPct, 2)

        If gapPct >= GapThresholdPct Then
            If gapPct > bestGap Then
                bestGap = gapPct
                bestRow = r
            End If
        End If

ContinueLoop:
    Next r

    If bestRow > 0 Then
        FireSignal ws, bestRow, bestGap
        MsgBox "【テスト実行】" & checkedCount & "銘柄をチェックし、条件（" & _
               GapThresholdPct & "%以上）を満たす銘柄が見つかりました。", vbInformation, "テストモード"
    Else
        MsgBox "【テスト実行】" & checkedCount & "銘柄をチェックしましたが、" & vbCrLf & _
               "ギャップ率" & GapThresholdPct & "%以上の銘柄はありませんでした。", _
               vbInformation, "テストモード"
    End If
End Sub
