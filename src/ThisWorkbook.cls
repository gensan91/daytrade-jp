' ============================================================
' ThisWorkbook
' 2026/8/19 改訂
'   [L] 起動時に「大引け追跡が未完走のまま」を警告するロジックを統合。
'
'   旧版はこのロジックを標準モジュール「ThisWorkbook1」に置いていたが、
'   標準モジュールの Private Sub Workbook_Open() はただの同名プロシージャに
'   過ぎず、Excelのイベントとしては一切発火しない（Workbook_Openイベントが
'   有効なのは、この ThisWorkbook クラスモジュール自身に書いた場合のみ）。
'   そのため [L] は「実装済み」と記録されていたが実際は完全に無効だった。
'
'   原因はおそらくインポート手順：ThisWorkbook はドキュメントモジュールで
'   あり、同名の .cls を File > Import File しても既存の ThisWorkbook を
'   置き換えることはできず、"ThisWorkbook1" という別モジュールとして
'   追加されてしまう。ThisWorkbook/Sheet系の変更は「コードペインを直接
'   全選択→削除→貼り付け」で行うこと。Import File が使えるのは
'   modXxx.bas のような標準モジュールのみ。
' ============================================================
Private Sub Workbook_Open()
    modUIButtons.ShowController
    WarnIfTrackingUnfinished
End Sub

' ------------------------------------------------------------
' [L] 前回の大引け追跡が完走したかをシート上の痕跡から確認する。
'
' modWatchdog の見張りはメモリ上の変数に依存しているため、
' VBAの状態がリセットされる事故（停止ボタン・コード編集・End・
' Excelの異常終了）では見張り役ごと消える。原理的に自己検知できない。
' そこで「Position T2 が 追跡中 のまま残っているか」を翌朝に見る。
' ------------------------------------------------------------
Private Sub WarnIfTrackingUnfinished()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Position")
    If ws Is Nothing Then Exit Sub

    Dim t As String
    t = CStr(ws.Cells(2, 20).Value)

    If Left$(t, 3) = "追跡中" Then
        MsgBox "前回の大引け追跡が完走していません。" & vbCrLf & vbCrLf & _
               "Position T2: " & t & vbCrLf & vbCrLf & _
               "該当する Trades 行の O/P列（大引けMFE/MAE）は" & vbCrLf & _
               "終値までの値ではありません。U列に打ち切り時刻を注記してください。" & vbCrLf & _
               "（Q列は完走時のみ書かれる仕様なので空欄のはずです）", _
               vbExclamation, "大引け追跡が未完走です"
    End If

    On Error GoTo 0
End Sub
