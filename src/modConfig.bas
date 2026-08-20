Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================
' modConfig
' Config シートから全パラメータを読み込み、Public変数として
' 他モジュールに提供する。
' コード内に数値をハードコードしない方針
' ------------------------------------------------------------
' 2026/8/13 改訂
'   - Trades シート（トレード履歴）用の定数を追加
'   - Position シートの列定数を追加（マジックナンバー排除）
'   - 終値追跡終了時刻 / 往復コスト率 を追加
' ------------------------------------------------------------
' 2026/8/20 改訂（重要）
'   時刻系4項目（強制手仕舞い / 監視開始 / 監視終了 / 終値追跡終了）が
'   Config シートの値をまったく読めておらず、常に既定値で動いていた。
'
'   原因：時刻セルの .Value が Date ではなく Double（シリアル値）で
'   返るケースがあり、TimeValue() は Double を受け取れず
'   実行時エラー13（型が一致しません）になる。GetConfigTime は
'   On Error Resume Next でこれを飲み込んで既定値に落としていたため、
'   一切の警告が出ないまま設定が無視されていた。
'
'   B5/B6/B7 はシート値と既定値が同一（15:00 / 8:59 / 10:00）だったため
'   永久に観測不能だった。B13 を 15:31 に変更して初めて露見した。
'
'   対策：GetConfigTime を数値・日付・文字列の3型に対応させる。
'   あわせて Find を xlWhole 優先（失敗時のみ xlPart）に変更し、
'   既知issue#5（部分一致による誤ヒット）も塞ぐ。
' ============================================================

Public GapThresholdPct As Double      ' Q7: ギャップ率閾値(%)
Public StopLossPct As Double          ' Q9: 損切り率(%)
Public TakeProfitPct As Double        ' Q10: 利確率(%)
Public TimeStopMinutes As Double      ' Q8/Q9: 時間切れ(分)
Public ForcedExitTime As Date         ' Q10: 強制手仕舞い時刻
Public MonitorStartTime As Date       ' Q6: 気配監視開始
Public MonitorEndTime As Date         ' Q7: ギャップ判定監視終了
Public CheckIntervalSec As Long       ' Q19: チェック間隔(秒)
Public RiskPct As Double              ' Q11: 1トレード最大損失率(%)
Public AccountBalance As Double       ' Q11: 口座資金(円)
Public SoundOn As Boolean             ' Q13: 通知音ON/OFF
Public PopupOn As Boolean             ' Q13: ポップアップON/OFF
Public TrackEndTime As Date           ' 決済後の終値追跡を打ち切る時刻（既定15:30）
Public RoundTripCostPct As Double     ' 往復コスト率(%)（既定0.4）

Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_WATCHLIST As String = "Watchlist"
Public Const SHEET_POSITION As String = "Position"
Public Const SHEET_LOG As String = "Log"
Public Const SHEET_TRADES As String = "Trades"

' Watchlist列定義
Public Const COL_CODE As Long = 1          ' A: 証券コード
Public Const COL_NAME As Long = 2          ' B: 銘柄名
Public Const COL_PREVCLOSE As Long = 3     ' C: 前日終値
Public Const COL_CURRENT As Long = 4       ' D: 現在値
Public Const COL_ASK As Long = 5           ' E: 売気配値(Ask)
Public Const COL_VOLUME As Long = 6        ' F: 出来高
Public Const COL_GAPPCT As Long = 7        ' G: ギャップ率(%)
Public Const COL_SIGNAL As Long = 8        ' H: シグナル状態
Public Const COL_DETECTTIME As Long = 9    ' I: 検知時刻

Public Const WATCHLIST_START_ROW As Long = 2

' Position列定義（A～T）
Public Const PCOL_CODE As Long = 1         ' A: 証券コード
Public Const PCOL_ENTRYTIME As Long = 2    ' B: エントリー時刻（時間切れ判定の起点）
Public Const PCOL_ENTRYPRICE As Long = 3   ' C: エントリー価格
Public Const PCOL_MEMO As Long = 4         ' D: 注記
Public Const PCOL_CURRENT As Long = 5      ' E: 現在値（RSS数式を手動設定）
Public Const PCOL_PLPCT As Long = 6        ' F: 損益率(%)
Public Const PCOL_STATUS As Long = 7       ' G: 状態
Public Const PCOL_EXITREASON As Long = 8   ' H: 決済理由
Public Const PCOL_EXITTIME As Long = 9     ' I: 決済時刻
Public Const PCOL_SHARES As Long = 10      ' J: 推奨株数
Public Const PCOL_MFE As Long = 11         ' K: 保有中MFE(%)
Public Const PCOL_MAE As Long = 12         ' L: 保有中MAE(%)
Public Const PCOL_NAME As Long = 13        ' M: 銘柄名
Public Const PCOL_SIGNALTIME As Long = 14  ' N: シグナル検知時刻
Public Const PCOL_GAPPCT As Long = 15      ' O: ギャップ率(%)
Public Const PCOL_RECTIME As Long = 16     ' P: 記録操作時刻（ボタンを押した時刻）
Public Const PCOL_EXITPRICE As Long = 17   ' Q: 決済価格
Public Const PCOL_MFE_EOD As Long = 18     ' R: 大引けまでのMFE(%)
Public Const PCOL_MAE_EOD As Long = 19     ' S: 大引けまでのMAE(%)
Public Const PCOL_TRACK As Long = 20       ' T: 追跡状態

' 決済記録が Trades に退避済みであることを示す状態文字列
Public Const STATUS_ARCHIVED As String = "決済済(記録済)"
Public Const STATUS_WATCHING As String = "監視中"

' ------------------------------------------------------------
' Config シートのB列からラベル(A列)一致で値を取得する簡易読込
' ------------------------------------------------------------
Sub LoadConfig()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)

    GapThresholdPct = GetConfigValue(ws, "ギャップ率閾値(%)", 5)
    StopLossPct = GetConfigValue(ws, "損切り率(%)", -2)
    TakeProfitPct = GetConfigValue(ws, "利確率(%)", 4)
    TimeStopMinutes = GetConfigValue(ws, "時間切れ(分)", 15)
    RiskPct = GetConfigValue(ws, "1トレード最大損失率(%)", 1)
    AccountBalance = GetConfigValue(ws, "口座資金(円)", 1000000)
    CheckIntervalSec = CLng(GetConfigValue(ws, "チェック間隔(秒)", 3))
    RoundTripCostPct = GetConfigValue(ws, "往復コスト率(%)", 0.4)

    ForcedExitTime = GetConfigTime(ws, "強制手仕舞い時刻", "15:00")
    MonitorStartTime = GetConfigTime(ws, "監視開始時刻", "8:59")
    MonitorEndTime = GetConfigTime(ws, "監視終了時刻(ギャップ判定)", "10:00")
    TrackEndTime = GetConfigTime(ws, "終値追跡終了時刻", "15:30")

    SoundOn = GetConfigBool(ws, "通知音を鳴らす(TRUE/FALSE)", True)
    PopupOn = GetConfigBool(ws, "ポップアップ表示(TRUE/FALSE)", True)
End Sub

' ------------------------------------------------------------
' Config A列からラベル行を探す共通処理
' issue#5 対策：完全一致を優先し、見つからない場合のみ部分一致に落とす。
' （従来は xlPart のみだったため、部分一致するラベルを後から追加すると
'   誤った行にヒットする危険があった）
' ------------------------------------------------------------
Private Function FindConfigRow(ws As Worksheet, label As String) As Range
    Dim c As Range
    Set c = ws.Columns(1).Find(What:=label, LookIn:=xlValues, LookAt:=xlWhole)
    If c Is Nothing Then
        Set c = ws.Columns(1).Find(What:=label, LookIn:=xlValues, LookAt:=xlPart)
    End If
    Set FindConfigRow = c
End Function

Private Function GetConfigValue(ws As Worksheet, label As String, defaultVal As Double) As Double
    Dim c As Range
    Set c = FindConfigRow(ws, label)
    If c Is Nothing Then
        GetConfigValue = defaultVal
    Else
        If IsNumeric(ws.Cells(c.row, 2).Value) Then
            GetConfigValue = CDbl(ws.Cells(c.row, 2).Value)
        Else
            GetConfigValue = defaultVal
        End If
    End If
End Function

' ------------------------------------------------------------
' 時刻の取得
' セルの .Value は環境・書式によって Double / Date / String の
' いずれでも返りうる。3型すべてを明示的に処理する。
'
' Double の場合はExcelのシリアル値なので、整数部（日付）を捨てて
' 小数部（時刻）だけを Date 型に代入する。
' ------------------------------------------------------------
Private Function GetConfigTime(ws As Worksheet, label As String, defaultVal As String) As Date
    Dim c As Range
    Set c = FindConfigRow(ws, label)
    If c Is Nothing Then
        GetConfigTime = TimeValue(defaultVal)
        Exit Function
    End If

    Dim v As Variant
    v = ws.Cells(c.row, 2).Value

    Dim result As Date
    Dim ok As Boolean
    ok = False

    On Error Resume Next

    If IsNumeric(v) Then
        Dim d As Double
        d = CDbl(v)
        result = CDate(d - Int(d))
        If Err.Number = 0 Then ok = True Else Err.Clear
    ElseIf IsDate(v) Then
        result = TimeValue(CDate(v))
        If Err.Number = 0 Then ok = True Else Err.Clear
    Else
        result = TimeValue(CStr(v))
        If Err.Number = 0 Then ok = True Else Err.Clear
    End If

    On Error GoTo 0

    If ok Then
        GetConfigTime = result
    Else
        GetConfigTime = TimeValue(defaultVal)
    End If
End Function

Private Function GetConfigBool(ws As Worksheet, label As String, defaultVal As Boolean) As Boolean
    Dim c As Range
    Set c = FindConfigRow(ws, label)
    If c Is Nothing Then
        GetConfigBool = defaultVal
    Else
        On Error Resume Next
        GetConfigBool = CBool(ws.Cells(c.row, 2).Value)
        If Err.Number <> 0 Then
            Err.Clear
            GetConfigBool = defaultVal
        End If
        On Error GoTo 0
    End If
End Function

' ------------------------------------------------------------
' 動作確認用：読み込んだ設定値を一覧表示する。
' 「シートに書いた値が本当に効いているか」を目視するための唯一の手段。
' Config を変更したら必ずこれを実行して確認すること。
' ------------------------------------------------------------
Sub ShowLoadedConfig()
    LoadConfig
    MsgBox "ギャップ率閾値: " & GapThresholdPct & "%" & vbCrLf & _
           "損切り率: " & StopLossPct & "%" & vbCrLf & _
           "利確率: " & TakeProfitPct & "%" & vbCrLf & _
           "時間切れ: " & TimeStopMinutes & "分" & vbCrLf & _
           "1トレード最大損失率: " & RiskPct & "%" & vbCrLf & _
           "口座資金: " & Format(AccountBalance, "#,##0") & "円" & vbCrLf & _
           "チェック間隔: " & CheckIntervalSec & "秒" & vbCrLf & _
           "往復コスト率: " & RoundTripCostPct & "%" & vbCrLf & vbCrLf & _
           "強制手仕舞い時刻: " & Format(ForcedExitTime, "hh:mm:ss") & vbCrLf & _
           "監視開始時刻: " & Format(MonitorStartTime, "hh:mm:ss") & vbCrLf & _
           "監視終了時刻: " & Format(MonitorEndTime, "hh:mm:ss") & vbCrLf & _
           "終値追跡終了時刻: " & Format(TrackEndTime, "hh:mm:ss") & vbCrLf & vbCrLf & _
           "通知音: " & SoundOn & vbCrLf & _
           "ポップアップ: " & PopupOn, _
           vbInformation, "読み込まれた設定値"
End Sub
