Attribute VB_Name = "modWatchdog"
Option Explicit

' ============================================================
' modWatchdog（新規・2026/8/13）
'
' 目的：8/12に発生した「タイマー連鎖が黙って死ぬ」事故への対策。
'
' Application.OnTime で予約したプロシージャは、その時刻に Excel が
' モーダルダイアログ表示中・セル編集中・重い再計算中などで
' 実行不能だった場合、リトライの末に予約ごと破棄されることがある。
' 一度破棄されると自己再予約の連鎖が途切れ、以後まったく動かない。
' しかも例外は出ないので画面上は何も起きない。
'
' そこで本モジュールは60秒間隔の「見張り役」を独立して回し、
' 各チェーンの最終実行時刻が古すぎる場合に予約を張り直す。
' 見張り役自身も毎回自分を再予約する（最初にRearmを呼ぶ前に予約する）。
'
' 注意：二重連鎖を防ぐため、張り直す側（modMonitor/modPosition）は
' 「既存の予約をキャンセルしてから再予約」する実装になっている。
' ============================================================

Private mWatchdogOn As Boolean
Private mNextWatchdog As Date
Private mRecoveryCount As Long

Private Const WATCHDOG_INTERVAL_SEC As Long = 60

Sub StartWatchdog()
    If mWatchdogOn Then Exit Sub
    mWatchdogOn = True
    mRecoveryCount = 0
    ScheduleNextWatchdog
End Sub

Sub StopWatchdog()
    mWatchdogOn = False
    On Error Resume Next
    Application.OnTime EarliestTime:=mNextWatchdog, Procedure:="WatchdogTick", Schedule:=False
    On Error GoTo 0
End Sub

Private Sub ScheduleNextWatchdog()
    If Not mWatchdogOn Then Exit Sub
    mNextWatchdog = Now + TimeSerial(0, 0, WATCHDOG_INTERVAL_SEC)
    On Error Resume Next
    Application.OnTime EarliestTime:=mNextWatchdog, Procedure:="WatchdogTick"
    On Error GoTo 0
End Sub

Sub WatchdogTick()
    If Not mWatchdogOn Then Exit Sub

    ' 自分の次回予約を最優先で確保する
    ScheduleNextWatchdog

    On Error Resume Next
    modMonitor.WatchdogRearmGap
    modPosition.WatchdogRearmPos
    On Error GoTo 0

    ' 監視対象が全部止まっていたら見張りも降りる
    If Not modMonitor.GapChainIsRunning() _
       And Not modPosition.PositionChainIsRunning() _
       And Not modPosition.TrackingIsRunning() Then
        StopWatchdog
    End If
End Sub

' ------------------------------------------------------------
' 動作確認用：現在の見張り状態をメッセージで表示する
' ------------------------------------------------------------
Sub ShowWatchdogStatus()
    MsgBox "見張り役: " & IIf(mWatchdogOn, "稼働中", "停止") & vbCrLf & _
           "次回チェック: " & IIf(mWatchdogOn, Format(mNextWatchdog, "hh:mm:ss"), "-") & vbCrLf & _
           "ギャップ監視チェーン: " & IIf(modMonitor.GapChainIsRunning(), "稼働中", "停止") & vbCrLf & _
           "ポジション監視チェーン: " & IIf(modPosition.PositionChainIsRunning(), "稼働中", "停止") & vbCrLf & _
           "大引け追跡チェーン: " & IIf(modPosition.TrackingIsRunning(), "稼働中", "停止"), _
           vbInformation, "Watchdog"
End Sub
