Attribute VB_Name = "modPosition"
Option Explicit

' ============================================================
' modPosition
' Q9/Q10: 損切り・利確・時間切れ・強制手仕舞い
'
' 2026/8/19 改訂
'   [J] Q列（終値ベース損益率）を FinalizeTracking でのみ書くよう変更。
'       従来は追跡tickごとに上書きしていたため、連鎖が途中で死ぬと
'       「終値ではない数字」がQ列に残り、見た目で区別できなかった。
'       あわせて U列の完了注記に時刻を入れる。
' Q11: 1トレード最大損失率からのポジションサイズ逆算
' Q20: エントリー/決済結果のログ記録
' ------------------------------------------------------------
' 2026/8/13 改訂（A～D 一括対応）
'   [A] 決済時に Trades シートへ1行追記して履歴を残す。
'       Position シートは作業用（上書き）、Trades が一次データ。
'   [B] MFE/MAE を2系統に分離。
'         K/L列 = 保有中（エントリー～決済）
'         R/S列 = 大引け（終値追跡終了時刻）まで継続追跡
'       8/10のように手入力で混ぜない。両方コードが記録する。
'   [C] エントリー時刻をシグナル検知時刻に合わせられるようにした。
'       価格が9:00の気配値なのに時計が9:07、という混在を解消。
'       ボタンを押した時刻は P列に別途記録する。
'   [D] 決済アラートのMsgBoxは1回だけ。以降は音とステータスバーのみ。
'       次回OnTime予約を判定より先に行い、連鎖切れを防ぐ。
' 2026/8/14 改訂
'   [G] Trades の「日付」を退避実行日ではなくエントリー時刻から決める。
'       自動退避だと翌営業日の日付が入ってしまっていた。
'   [H] 保有時間(分)を型に依存しない実装に変更。IsDate/CDate 経由だと
'       セル書式によって判定が落ち、全行が空欄になっていた。
'   [I] 発火区分（寄り付き型／ザラ場型）を自動判定してV列に記録。
'       9:08発火の8550のような銘柄を後から分離集計できるようにする。
' ============================================================

Private Const POS_ROW As Long = 2   ' 単一ポジション運用のため2行目固定

Private mPositionOpen As Boolean
Private mNextPosCheck As Date
Private mForcedExitScheduled As Boolean
Private mExitAlerted As Boolean
Private mLastAlertSound As Date
Private mLastPosCheck As Date
Private mPosErrCount As Long

' 決済後の大引け追跡
Private mTrackOn As Boolean
Private mNextTrackCheck As Date
Private mTrackEntryPrice As Double
Private mTradesRow As Long

' [J] 追跡の最新値。Q列（終値ベース損益率）は追跡が完走したときだけ書く。
'     途中で連鎖が死んだ日に「終値ではない数字」がQ列に残るのを防ぐ。
'     8/14（15:29:47で打ち切り）と8/19（10:40:25で停止）で実害が出た。
Private mLastTrackPct As Double
Private mTrackHasValue As Boolean

Private Const MAX_CONSECUTIVE_ERRORS As Long = 5
Private Const ALERT_SOUND_INTERVAL_SEC As Long = 30
Private Const STALE_LIMIT_SEC As Long = 30
Private Const TRACK_INTERVAL_SEC As Long = 15

' ============================================================
' エントリー記録
' ============================================================
Sub RecordEntry()
    LoadConfig

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_POSITION)

    WritePositionHeaders ws

    ' [A] 前回のトレードが未退避のまま残っていたら先に Trades へ逃がす
    If Trim(ws.Cells(POS_ROW, PCOL_CODE).Value) <> "" Then
        If ws.Cells(POS_ROW, PCOL_STATUS).Value <> STATUS_ARCHIVED Then
            ArchiveTrade ws, "自動退避（RecordExit未実行の可能性あり）"
        End If
    End If

    ' [C] 直近シグナルの内容を初期値として提示する
    Dim defCode As String, defPrice As String
    If modMonitor.LastSignalValid Then
        defCode = modMonitor.LastSignalCode
        defPrice = CStr(modMonitor.LastSignalAsk)
    End If

    Dim code As String
    code = InputBox("エントリーした証券コードを入力してください", "エントリー記録", defCode)
    If code = "" Then Exit Sub

    Dim fillPrice As String
    fillPrice = InputBox("実際の約定価格を入力してください（成行約定価格）", "エントリー記録", defPrice)
    If Not IsNumeric(fillPrice) Then
        MsgBox "価格は数値で入力してください。", vbExclamation
        Exit Sub
    End If

    ' [C] 時間切れ判定の起点をどちらにするか明示的に選ばせる
    Dim entryTime As Date
    entryTime = Now

    Dim useSignalTime As Boolean
    useSignalTime = False

    If modMonitor.LastSignalValid And code = modMonitor.LastSignalCode Then
        Dim lagSec As Double
        lagSec = (Now - modMonitor.LastSignalTime) * 86400#

        Dim ans As VbMsgBoxResult
        ans = MsgBox( _
            "シグナル検知は " & Format(modMonitor.LastSignalTime, "hh:mm:ss") & _
            "、現在は " & Format(Now, "hh:mm:ss") & "（" & Format(lagSec, "0") & "秒経過）です。" & vbCrLf & vbCrLf & _
            "エントリー時刻（＝時間切れ15分の起点）をどちらにしますか？" & vbCrLf & vbCrLf & _
            "［はい］ シグナル検知時刻を使う" & vbCrLf & _
            "　　　　 → 紙トレード検証向け。価格が気配値なら時刻も揃える。" & vbCrLf & _
            "［いいえ］ 現在時刻を使う" & vbCrLf & _
            "　　　　 → 実際に今この瞬間約定した場合。", _
            vbQuestion + vbYesNo, "エントリー時刻の基準")

        If ans = vbYes Then
            entryTime = modMonitor.LastSignalTime
            useSignalTime = True
        End If
    End If

    Dim shares As Long
    shares = CalcPositionSize(CDbl(fillPrice))

    If shares = 0 Then
        Dim actualRiskPct As Double
        actualRiskPct = (100 * CDbl(fillPrice) * Abs(StopLossPct) / 100) / AccountBalance * 100

        Dim answer As VbMsgBoxResult
        answer = MsgBox( _
            "この株価（" & fillPrice & "円）では、Q11ルール（1トレード最大損失=口座資金の" & RiskPct & "%）を守ると" & vbCrLf & _
            "推奨株数が100株未満（=0株）になります。" & vbCrLf & vbCrLf & _
            "最低単元の100株で発注すると、実際のリスク率は約" & Format(actualRiskPct, "0.00") & "%になります。" & vbCrLf & vbCrLf & _
            "それでも100株として記録しますか？", _
            vbExclamation + vbYesNo, "リスク率超過の警告")

        If answer = vbYes Then shares = 100
    End If

    ' 前回のトレードの残骸を消す（E列のRSS数式は消さない）
    ws.Range(ws.Cells(POS_ROW, PCOL_PLPCT), ws.Cells(POS_ROW, PCOL_PLPCT)).ClearContents
    ws.Range(ws.Cells(POS_ROW, PCOL_EXITREASON), ws.Cells(POS_ROW, PCOL_EXITTIME)).ClearContents
    ws.Range(ws.Cells(POS_ROW, PCOL_MFE), ws.Cells(POS_ROW, PCOL_MAE)).ClearContents
    ws.Range(ws.Cells(POS_ROW, PCOL_EXITPRICE), ws.Cells(POS_ROW, PCOL_TRACK)).ClearContents
    ws.Range(ws.Cells(POS_ROW, PCOL_CODE), ws.Cells(POS_ROW, PCOL_TRACK)).Interior.ColorIndex = xlNone

    ws.Cells(POS_ROW, PCOL_CODE).Value = code
    ws.Cells(POS_ROW, PCOL_ENTRYTIME).Value = entryTime
    ws.Cells(POS_ROW, PCOL_ENTRYPRICE).Value = CDbl(fillPrice)
    ws.Cells(POS_ROW, PCOL_STATUS).Value = STATUS_WATCHING
    ws.Cells(POS_ROW, PCOL_SHARES).Value = shares
    ws.Cells(POS_ROW, PCOL_RECTIME).Value = Now

    If modMonitor.LastSignalValid And code = modMonitor.LastSignalCode Then
        ws.Cells(POS_ROW, PCOL_NAME).Value = modMonitor.LastSignalName
        ws.Cells(POS_ROW, PCOL_SIGNALTIME).Value = modMonitor.LastSignalTime
        ws.Cells(POS_ROW, PCOL_GAPPCT).Value = Round(modMonitor.LastSignalGapPct, 2)
    Else
        ws.Cells(POS_ROW, PCOL_NAME).Value = LookupName(code)
    End If

    mPositionOpen = True
    mExitAlerted = False
    mForcedExitScheduled = False
    mPosErrCount = 0
    mLastPosCheck = 0
    mTradesRow = 0

    StopAfterExitTracking

    MsgBox "エントリーを記録しました。" & vbCrLf & vbCrLf & _
           "銘柄: " & code & vbCrLf & _
           "価格: " & fillPrice & " 円 ／ 推奨株数: " & shares & " 株" & vbCrLf & _
           "エントリー時刻: " & Format(entryTime, "hh:mm:ss") & _
           IIf(useSignalTime, "（シグナル検知時刻を採用）", "（現在時刻）") & vbCrLf & _
           "時間切れ判定: " & Format(entryTime + TimeSerial(0, CLng(TimeStopMinutes), 0), "hh:mm:ss") & " 以降", _
           vbInformation, "エントリー記録"

    StartPositionMonitoring
End Sub

' ------------------------------------------------------------
' Q11: リスクベースのポジションサイズ算出
' ------------------------------------------------------------
Function CalcPositionSize(entryPrice As Double) As Long
    Dim maxLossAmount As Double
    maxLossAmount = AccountBalance * (RiskPct / 100)

    Dim lossPerShare As Double
    lossPerShare = entryPrice * Abs(StopLossPct) / 100

    If lossPerShare <= 0 Then
        CalcPositionSize = 0
        Exit Function
    End If

    Dim rawShares As Double
    rawShares = maxLossAmount / lossPerShare

    CalcPositionSize = Int(rawShares / 100) * 100
End Function

' ============================================================
' ポジション監視
' ============================================================
Sub StartPositionMonitoring()
    ScheduleNextPositionCheck
    ScheduleForcedExitReminder
    modWatchdog.StartWatchdog
End Sub

Sub StopPositionMonitoring()
    mPositionOpen = False
    CancelPendingPositionCheck
    StopAfterExitTracking
    Application.StatusBar = "ポジション監視: 停止しました"
End Sub

Private Sub CancelPendingPositionCheck()
    On Error Resume Next
    Application.OnTime EarliestTime:=mNextPosCheck, Procedure:="CheckPositionStatus", Schedule:=False
    On Error GoTo 0
End Sub

Private Sub ScheduleNextPositionCheck()
    If Not mPositionOpen Then Exit Sub
    mNextPosCheck = Now + TimeSerial(0, 0, CheckIntervalSec)
    Application.OnTime EarliestTime:=mNextPosCheck, Procedure:="CheckPositionStatus"
End Sub

' ------------------------------------------------------------
' 損切り・利確・時間切れの判定
' ------------------------------------------------------------
Sub CheckPositionStatus()
    If Not mPositionOpen Then Exit Sub

    On Error GoTo ErrHandler

    ' [D] 判定より先に次回を予約する
    ScheduleNextPositionCheck

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_POSITION)

    If ws.Cells(POS_ROW, PCOL_STATUS).Value <> STATUS_WATCHING Then
        mPositionOpen = False
        CancelPendingPositionCheck
        Exit Sub
    End If

    Dim entryPrice As Double, currentPrice As Double, entryTime As Date
    entryPrice = SafeNumP(ws.Cells(POS_ROW, PCOL_ENTRYPRICE).Value)
    currentPrice = SafeNumP(ws.Cells(POS_ROW, PCOL_CURRENT).Value)
    entryTime = ws.Cells(POS_ROW, PCOL_ENTRYTIME).Value

    mLastPosCheck = Now
    mPosErrCount = 0

    If entryPrice > 0 And currentPrice > 0 Then
        Dim plPct As Double
        plPct = (currentPrice - entryPrice) / entryPrice * 100
        ws.Cells(POS_ROW, PCOL_PLPCT).Value = Round(plPct, 2)

        ' [B] 保有中MFE/MAE（K/L）
        UpdateExcursion ws, PCOL_MFE, PCOL_MAE, plPct

        ' [B] 大引けまでのMFE/MAE（R/S）。保有中から通しで記録する。
        UpdateExcursion ws, PCOL_MFE_EOD, PCOL_MAE_EOD, plPct

        Dim reason As String
        reason = ""

        If plPct >= TakeProfitPct Then
            reason = "利確ライン到達（+" & Format(plPct, "0.00") & "%）"
        ElseIf plPct <= StopLossPct Then
            reason = "損切りライン到達（" & Format(plPct, "0.00") & "%）"
        Else
            Dim elapsedMin As Double
            elapsedMin = (Now - entryTime) * 24 * 60
            If elapsedMin >= TimeStopMinutes And plPct < TakeProfitPct / 2 Then
                reason = "時間切れ（" & TimeStopMinutes & "分経過、含み益が利確の半分未満）"
            End If
        End If

        If reason <> "" Then
            ClosePositionAlert reason, plPct
        Else
            Application.StatusBar = "ポジション監視: " & ws.Cells(POS_ROW, PCOL_CODE).Value & _
                " 損益 " & Format(plPct, "0.00") & "%  最終チェック " & Format(mLastPosCheck, "hh:mm:ss")
        End If
    Else
        Application.StatusBar = "ポジション監視: 現在値が取得できません（Position E2のRSS数式を確認）  " & _
            Format(Now, "hh:mm:ss")
    End If

    Exit Sub

ErrHandler:
    mPosErrCount = mPosErrCount + 1
    Application.StatusBar = "ポジション監視: エラー" & mPosErrCount & "回目 " & _
        Format(Now, "hh:mm:ss") & " [" & Err.Number & "] " & Err.Description
    Err.Clear

    If mPosErrCount >= MAX_CONSECUTIVE_ERRORS Then
        mPositionOpen = False
        CancelPendingPositionCheck
        MsgBox "ポジション監視でエラーが" & MAX_CONSECUTIVE_ERRORS & "回連続したため停止しました。", _
               vbCritical, "監視停止"
    End If
End Sub

' ------------------------------------------------------------
' MFE/MAE の更新（最大値列・最小値列を指定して使う）
' ------------------------------------------------------------
Private Sub UpdateExcursion(ws As Worksheet, maxCol As Long, minCol As Long, plPct As Double)
    If ws.Cells(POS_ROW, maxCol).Value = "" Then
        ws.Cells(POS_ROW, maxCol).Value = Round(plPct, 2)
    ElseIf plPct > SafeNumP(ws.Cells(POS_ROW, maxCol).Value) Then
        ws.Cells(POS_ROW, maxCol).Value = Round(plPct, 2)
    End If

    If ws.Cells(POS_ROW, minCol).Value = "" Then
        ws.Cells(POS_ROW, minCol).Value = Round(plPct, 2)
    ElseIf plPct < SafeNumP(ws.Cells(POS_ROW, minCol).Value) Then
        ws.Cells(POS_ROW, minCol).Value = Round(plPct, 2)
    End If
End Sub

' 日付でも数値でもシリアル値(Double)として取り出す。
' セルの表示形式によって IsDate / IsNumeric の結果が変わるため、
' 両方を受けられるようにしておく。
Private Function ToSerial(v As Variant) As Double
    ToSerial = 0
    On Error Resume Next
    If IsDate(v) Then
        ToSerial = CDbl(CDate(v))
    ElseIf IsNumeric(v) Then
        ToSerial = CDbl(v)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        ToSerial = 0
    End If
    On Error GoTo 0
End Function

Private Function SafeNumP(v As Variant) As Double
    If IsNumeric(v) Then
        SafeNumP = CDbl(v)
    Else
        SafeNumP = 0
    End If
End Function

' ------------------------------------------------------------
' [D] 決済アラート
' MsgBoxは最初の1回だけ。以降は音（30秒間隔）とステータスバーで知らせる。
' モーダルダイアログを繰り返し出すとOnTimeが実行できず連鎖が死ぬため。
' ------------------------------------------------------------
Private Sub ClosePositionAlert(reason As String, plPct As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_POSITION)
    ws.Range(ws.Cells(POS_ROW, PCOL_CODE), ws.Cells(POS_ROW, PCOL_MAE)).Interior.Color = RGB(255, 235, 156)

    Application.StatusBar = "★決済シグナル: " & reason & _
        "　現在 " & Format(plPct, "0.00") & "%　→ 決済したら「決済記録」ボタン　" & Format(Now, "hh:mm:ss")

    If Not mExitAlerted Then
        mExitAlerted = True
        mLastAlertSound = Now

        If SoundOn Then modNotify.PlayAlertSound 3

        If PopupOn Then
            MsgBox "【決済シグナル】" & vbCrLf & reason & vbCrLf & vbCrLf & _
                   "手動でマーケットスピードII取引画面から成行決済してください。" & vbCrLf & _
                   "決済できたら操作パネルの「決済記録」を押してください。" & vbCrLf & vbCrLf & _
                   "※このポップアップは1回だけです。以降はステータスバーと通知音で知らせます。", _
                   vbExclamation, "デイトレ支援ツール"
        End If
    Else
        If SoundOn And (Now - mLastAlertSound) * 86400# >= ALERT_SOUND_INTERVAL_SEC Then
            mLastAlertSound = Now
            modNotify.PlayAlertSound 1
        End If
    End If
End Sub

' ------------------------------------------------------------
' 引け前の強制手仕舞いリマインダー
' ------------------------------------------------------------
Private Sub ScheduleForcedExitReminder()
    If mForcedExitScheduled Then Exit Sub
    mForcedExitScheduled = True

    Dim fireTime As Date
    fireTime = DateValue(Now) + ForcedExitTime
    If fireTime < Now Then Exit Sub

    On Error Resume Next
    Application.OnTime EarliestTime:=fireTime, Procedure:="ForcedExitReminder"
    On Error GoTo 0
End Sub

Sub ForcedExitReminder()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_POSITION)

    If ws.Cells(POS_ROW, PCOL_STATUS).Value = STATUS_WATCHING Then
        If SoundOn Then modNotify.PlayAlertSound 4
        MsgBox "【引け前強制手仕舞い】" & vbCrLf & _
               Format(ForcedExitTime, "hh:mm") & "になりました。" & vbCrLf & _
               "オーバーナイトを避けるため、必ず本日中に決済してください。", _
               vbCritical, "デイトレ支援ツール"
    End If
End Sub

' ============================================================
' 決済記録
' ============================================================
Sub RecordExit()
    LoadConfig

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_POSITION)

    If ws.Cells(POS_ROW, PCOL_STATUS).Value <> STATUS_WATCHING Then
        MsgBox "現在監視中のポジションがありません。", vbExclamation
        Exit Sub
    End If

    Dim defExit As String
    defExit = CStr(SafeNumP(ws.Cells(POS_ROW, PCOL_CURRENT).Value))

    Dim exitPrice As String
    exitPrice = InputBox("実際の決済価格を入力してください", "決済記録", defExit)
    If Not IsNumeric(exitPrice) Then
        MsgBox "価格は数値で入力してください。", vbExclamation
        Exit Sub
    End If

    Dim reason As String
    reason = InputBox("決済理由を入力してください（例: 損切り／利確／時間切れ／引け前強制）", _
                      "決済記録", ws.Cells(POS_ROW, PCOL_EXITREASON).Value)

    Dim entryPrice As Double
    entryPrice = SafeNumP(ws.Cells(POS_ROW, PCOL_ENTRYPRICE).Value)

    Dim plPct As Double
    plPct = (CDbl(exitPrice) - entryPrice) / entryPrice * 100

    ws.Cells(POS_ROW, PCOL_STATUS).Value = "決済済"
    ws.Cells(POS_ROW, PCOL_EXITREASON).Value = reason
    ws.Cells(POS_ROW, PCOL_EXITTIME).Value = Now
    ws.Cells(POS_ROW, PCOL_EXITPRICE).Value = CDbl(exitPrice)
    ws.Cells(POS_ROW, PCOL_PLPCT).Value = Round(plPct, 2)
    ws.Range(ws.Cells(POS_ROW, PCOL_CODE), ws.Cells(POS_ROW, PCOL_MAE)).Interior.ColorIndex = xlNone

    ' 決済価格そのものも大引けMFE/MAEの候補に含める
    UpdateExcursion ws, PCOL_MFE, PCOL_MAE, plPct
    UpdateExcursion ws, PCOL_MFE_EOD, PCOL_MAE_EOD, plPct

    mPositionOpen = False
    CancelPendingPositionCheck
    mExitAlerted = False

    ' [A] Trades シートへ退避（ここが一次データ）
    ArchiveTrade ws, ""

    ' Log シートにも従来どおり1行残す
    WriteExitLog ws, plPct, reason

    ' [B] 大引けまで継続追跡（R/S列と Trades を更新し続ける）
    StartAfterExitTracking entryPrice

    MsgBox "決済を記録しました。損益率: " & Format(plPct, "0.00") & "%" & vbCrLf & vbCrLf & _
           "Trades シートの " & mTradesRow & " 行目に履歴を保存しました。" & vbCrLf & _
           "大引け（" & Format(TrackEndTime, "hh:mm") & "）まで値動きの追跡を続けます。" & vbCrLf & _
           "※ブックを閉じると追跡は止まります。", _
           vbInformation, "決済記録"
End Sub

' ============================================================
' [A] Trades シート（トレード履歴）
' ============================================================
Private Function EnsureTradesSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_TRADES)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_TRADES
    End If

    If ws.Cells(1, 1).Value = "" Then
        ws.Range("A1:U1").Value = Array( _
            "日付", "証券コード", "銘柄名", "シグナル検知時刻", "ギャップ率(%)", _
            "エントリー時刻", "エントリー価格", "株数", _
            "決済時刻", "決済価格", "決済理由", "損益率(%)", _
            "保有中MFE(%)", "保有中MAE(%)", "大引けMFE(%)", "大引けMAE(%)", "終値ベース損益率(%)", _
            "保有時間(分)", "概算損益(円)", "コスト後損益(円)", "備考")
        ws.Rows(1).Font.Bold = True
    End If

    ' 既存シートにも後から列・書式を足せるようにしておく
    If ws.Cells(1, 22).Value = "" Then ws.Cells(1, 22).Value = "発火区分"
    ws.Range("A:A").NumberFormat = "yyyy/m/d"
    ws.Range("D:D").NumberFormat = "yyyy/m/d hh:mm:ss"
    ws.Range("F:F").NumberFormat = "yyyy/m/d hh:mm:ss"
    ws.Range("I:I").NumberFormat = "yyyy/m/d hh:mm:ss"
    ws.Columns("A:V").AutoFit

    Set EnsureTradesSheet = ws
End Function

Private Sub ArchiveTrade(pos As Worksheet, note As String)
    Dim tr As Worksheet
    Set tr = EnsureTradesSheet()

    Dim r As Long
    r = tr.Cells(tr.Rows.Count, 1).End(xlUp).row + 1

    Dim entryPrice As Double, exitPrice As Double, shares As Double
    entryPrice = SafeNumP(pos.Cells(POS_ROW, PCOL_ENTRYPRICE).Value)
    exitPrice = SafeNumP(pos.Cells(POS_ROW, PCOL_EXITPRICE).Value)
    shares = SafeNumP(pos.Cells(POS_ROW, PCOL_SHARES).Value)

    ' [G] 日付はエントリー時刻から決める（自動退避で翌日日付になるのを防ぐ）
    Dim tEntry As Double, tExit As Double, tSignal As Double
    tEntry = ToSerial(pos.Cells(POS_ROW, PCOL_ENTRYTIME).Value)
    tExit = ToSerial(pos.Cells(POS_ROW, PCOL_EXITTIME).Value)
    tSignal = ToSerial(pos.Cells(POS_ROW, PCOL_SIGNALTIME).Value)

    If tEntry > 0 Then
        tr.Cells(r, 1).Value = CDate(Int(tEntry))
    Else
        tr.Cells(r, 1).Value = Date
    End If

    tr.Cells(r, 2).Value = pos.Cells(POS_ROW, PCOL_CODE).Value
    tr.Cells(r, 3).Value = pos.Cells(POS_ROW, PCOL_NAME).Value
    tr.Cells(r, 4).Value = pos.Cells(POS_ROW, PCOL_SIGNALTIME).Value
    tr.Cells(r, 5).Value = pos.Cells(POS_ROW, PCOL_GAPPCT).Value
    tr.Cells(r, 6).Value = pos.Cells(POS_ROW, PCOL_ENTRYTIME).Value
    tr.Cells(r, 7).Value = pos.Cells(POS_ROW, PCOL_ENTRYPRICE).Value
    tr.Cells(r, 8).Value = pos.Cells(POS_ROW, PCOL_SHARES).Value
    tr.Cells(r, 9).Value = pos.Cells(POS_ROW, PCOL_EXITTIME).Value
    tr.Cells(r, 10).Value = pos.Cells(POS_ROW, PCOL_EXITPRICE).Value
    tr.Cells(r, 11).Value = pos.Cells(POS_ROW, PCOL_EXITREASON).Value
    tr.Cells(r, 12).Value = pos.Cells(POS_ROW, PCOL_PLPCT).Value
    tr.Cells(r, 13).Value = pos.Cells(POS_ROW, PCOL_MFE).Value
    tr.Cells(r, 14).Value = pos.Cells(POS_ROW, PCOL_MAE).Value
    tr.Cells(r, 15).Value = pos.Cells(POS_ROW, PCOL_MFE_EOD).Value
    tr.Cells(r, 16).Value = pos.Cells(POS_ROW, PCOL_MAE_EOD).Value

    ' [H] 保有時間（分）。セル書式に左右されないようシリアル値で計算する
    If tEntry > 0 And tExit > tEntry Then
        tr.Cells(r, 18).Value = Round((tExit - tEntry) * 1440#, 1)
    End If

    ' 概算損益とコスト後損益
    If entryPrice > 0 And exitPrice > 0 And shares > 0 Then
        Dim grossPL As Double, costAmt As Double
        grossPL = (exitPrice - entryPrice) * shares
        costAmt = entryPrice * shares * (RoundTripCostPct / 100)
        tr.Cells(r, 19).Value = Round(grossPL, 0)
        tr.Cells(r, 20).Value = Round(grossPL - costAmt, 0)
    End If

    tr.Cells(r, 21).Value = note

    ' [I] 発火区分。寄り付き直後か、ザラ場で条件を満たしたかを分ける。
    '     20件到達時に両者を混ぜて集計しないための目印。
    If tSignal > 0 Then
        If (tSignal - Int(tSignal)) <= TimeValue("9:01:00") Then
            tr.Cells(r, 22).Value = "寄り付き型"
        Else
            tr.Cells(r, 22).Value = "ザラ場型"
        End If
    End If

    pos.Cells(POS_ROW, PCOL_STATUS).Value = STATUS_ARCHIVED
    mTradesRow = r
End Sub

Private Sub WriteExitLog(pos As Worksheet, plPct As Double, reason As String)
    Dim logWs As Worksheet
    Set logWs = ThisWorkbook.Worksheets(SHEET_LOG)

    Dim nextRow As Long
    nextRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).row + 1

    logWs.Cells(nextRow, 1).Value = Now
    logWs.Cells(nextRow, 2).Value = "決済"
    logWs.Cells(nextRow, 3).Value = pos.Cells(POS_ROW, PCOL_CODE).Value
    logWs.Cells(nextRow, 4).Value = pos.Cells(POS_ROW, PCOL_NAME).Value
    logWs.Cells(nextRow, 5).Value = pos.Cells(POS_ROW, PCOL_GAPPCT).Value
    logWs.Cells(nextRow, 8).Value = "エントリー済み"
    logWs.Cells(nextRow, 9).Value = reason
    logWs.Cells(nextRow, 10).Value = "損益率: " & Format(plPct, "0.00") & "%／Trades " & mTradesRow & "行目"
End Sub

' ============================================================
' [B] 決済後の大引け追跡
' ============================================================
Private Sub StartAfterExitTracking(entryPrice As Double)
    If entryPrice <= 0 Then Exit Sub
    If TimeValue(Now) >= TrackEndTime Then
        FinalizeTracking
        Exit Sub
    End If

    mTrackOn = True
    mTrackEntryPrice = entryPrice
    mLastTrackPct = 0
    mTrackHasValue = False
    ScheduleNextTrack
    modWatchdog.StartWatchdog
End Sub

Private Sub ScheduleNextTrack()
    If Not mTrackOn Then Exit Sub
    mNextTrackCheck = Now + TimeSerial(0, 0, TRACK_INTERVAL_SEC)
    On Error Resume Next
    Application.OnTime EarliestTime:=mNextTrackCheck, Procedure:="TrackAfterExit"
    On Error GoTo 0
End Sub

Sub StopAfterExitTracking()
    If Not mTrackOn Then Exit Sub
    mTrackOn = False
    On Error Resume Next
    Application.OnTime EarliestTime:=mNextTrackCheck, Procedure:="TrackAfterExit", Schedule:=False
    On Error GoTo 0
End Sub

Sub TrackAfterExit()
    If Not mTrackOn Then Exit Sub

    On Error Resume Next

    If TimeValue(Now) >= TrackEndTime Then
        FinalizeTracking
        Exit Sub
    End If

    ScheduleNextTrack

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_POSITION)

    Dim cur As Double
    cur = SafeNumP(ws.Cells(POS_ROW, PCOL_CURRENT).Value)
    If cur <= 0 Or mTrackEntryPrice <= 0 Then Exit Sub

    Dim p As Double
    p = (cur - mTrackEntryPrice) / mTrackEntryPrice * 100

    UpdateExcursion ws, PCOL_MFE_EOD, PCOL_MAE_EOD, p

    ws.Cells(POS_ROW, PCOL_TRACK).Value = "追跡中 " & Format(Now, "hh:mm:ss") & _
        "／現在 " & Format(p, "0.00") & "%"

    ' [J] 最新値は変数に持つだけ。Q列へは FinalizeTracking で確定時に書く。
    mLastTrackPct = p
    mTrackHasValue = True

    If mTradesRow > 0 Then
        Dim tr As Worksheet
        Set tr = ThisWorkbook.Worksheets(SHEET_TRADES)
        tr.Cells(mTradesRow, 15).Value = ws.Cells(POS_ROW, PCOL_MFE_EOD).Value
        tr.Cells(mTradesRow, 16).Value = ws.Cells(POS_ROW, PCOL_MAE_EOD).Value
    End If
End Sub

Private Sub FinalizeTracking()
    mTrackOn = False

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_POSITION)
    ws.Cells(POS_ROW, PCOL_TRACK).Value = "追跡完了 " & Format(Now, "hh:mm:ss")

    If mTradesRow > 0 Then
        Dim tr As Worksheet
        Set tr = ThisWorkbook.Worksheets(SHEET_TRADES)

        ' [J] ここで初めてQ列を確定させる。
        '     Q列に値がある＝追跡が完走した、という保証になる。
        If mTrackHasValue Then
            tr.Cells(mTradesRow, 17).Value = Round(mLastTrackPct, 2)
        End If

        If tr.Cells(mTradesRow, 21).Value = "" Then
            tr.Cells(mTradesRow, 21).Value = "大引け追跡完了 " & Format(Now, "hh:mm:ss")
        End If
    End If
    On Error GoTo 0
End Sub

' ============================================================
' [D] 監視役（modWatchdog）から呼ばれる自己診断・再起動
' ============================================================
Public Function PositionChainIsRunning() As Boolean
    PositionChainIsRunning = mPositionOpen
End Function

Public Function TrackingIsRunning() As Boolean
    TrackingIsRunning = mTrackOn
End Function

Public Sub WatchdogRearmPos()
    ' ポジション監視チェーン
    If mPositionOpen Then
        Dim staleSec As Double
        If mLastPosCheck = 0 Then
            staleSec = STALE_LIMIT_SEC + 1
        Else
            staleSec = (Now - mLastPosCheck) * 86400#
        End If

        If staleSec > STALE_LIMIT_SEC Then
            CancelPendingPositionCheck
            ScheduleNextPositionCheck
        End If
    End If

    ' 大引け追跡チェーン
    If mTrackOn Then
        If Now > mNextTrackCheck + TimeSerial(0, 1, 0) Then
            ScheduleNextTrack
        End If
    End If
End Sub

' ============================================================
' 共通ユーティリティ
' ============================================================
Private Sub WritePositionHeaders(ws As Worksheet)
    ws.Cells(1, PCOL_CODE).Value = "証券コード"
    ws.Cells(1, PCOL_ENTRYTIME).Value = "エントリー時刻"
    ws.Cells(1, PCOL_ENTRYPRICE).Value = "エントリー価格"
    ws.Cells(1, PCOL_MEMO).Value = "(E列に現在値のRSS数式を手動設定)"
    ws.Cells(1, PCOL_CURRENT).Value = "現在値"
    ws.Cells(1, PCOL_PLPCT).Value = "損益率(%)"
    ws.Cells(1, PCOL_STATUS).Value = "状態"
    ws.Cells(1, PCOL_EXITREASON).Value = "決済理由"
    ws.Cells(1, PCOL_EXITTIME).Value = "決済時刻"
    ws.Cells(1, PCOL_SHARES).Value = "参考:推奨株数"
    ws.Cells(1, PCOL_MFE).Value = "保有中MFE(%)"
    ws.Cells(1, PCOL_MAE).Value = "保有中MAE(%)"
    ws.Cells(1, PCOL_NAME).Value = "銘柄名"
    ws.Cells(1, PCOL_SIGNALTIME).Value = "シグナル検知時刻"
    ws.Cells(1, PCOL_GAPPCT).Value = "ギャップ率(%)"
    ws.Cells(1, PCOL_RECTIME).Value = "記録操作時刻"
    ws.Cells(1, PCOL_EXITPRICE).Value = "決済価格"
    ws.Cells(1, PCOL_MFE_EOD).Value = "大引けMFE(%)"
    ws.Cells(1, PCOL_MAE_EOD).Value = "大引けMAE(%)"
    ws.Cells(1, PCOL_TRACK).Value = "追跡状態"
End Sub

Private Function LookupName(code As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_WATCHLIST)

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_CODE).End(xlUp).row

    For r = WATCHLIST_START_ROW To lastRow
        If Trim(CStr(ws.Cells(r, COL_CODE).Value)) = Trim(code) Then
            LookupName = ws.Cells(r, COL_NAME).Value
            Exit Function
        End If
    Next r
    On Error GoTo 0
End Function
