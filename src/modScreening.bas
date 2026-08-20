Attribute VB_Name = "modScreening"
Option Explicit

'======================================================
' コア30銘柄・イベント銘柄 自動スクリーニング
' 前提：MarketSpeed II RSSが起動・ログイン済みであること
'
' 使い方（JPX銘柄マスタ方式・推奨）：
'   0) BuildCandidateSheet_Test を実行し、RSSの値が正しく
'      取れるか（売買代金の単位、決算発表日の書式）を確認する
'      ※既に検証済み：売買代金は千円単位
'   1) ImportJPXMaster を実行し、JPX公式の東証上場銘柄一覧
'      （data_j.xls）からプライム・スタンダードの内国株式
'      だけを「JPX銘柄マスタ」シートに取り込む
'      ※月次更新のファイルなので、月1回実行すればよい
'   2) Call RunAllBatches を実行すると、全バッチを自動で
'      連続処理する（RSS反映待ち・表示銘柄数超過の検知も自動）
'      重い処理のため、監視時間帯（8:59-10:00）は避けること
'      途中で止まった場合は Call RunAllBatches(N) でN番目の
'      バッチから再開できる
'      （手動で1バッチずつ進めたい場合は、従来通り
'       BuildCandidateSheetBatch(番号) → FilterCandidates
'       を個別に呼び出すこともできる）
'======================================================

' 抽出条件（要調整）
Private Const BAIBAI_MIN As Double = 300000    ' 売買代金下限（千円単位、検証済み）
Private Const BAIBAI_MAX As Double = 1000000   ' 売買代金上限（千円単位、検証済み）

' RSSの同時取得可能銘柄数の上限（証券口座単位、Excel全体で共有）
' 参考：楽天証券RSS利用者情報より、上限は500銘柄
' 「全銘柄候補」シート単体でこの上限に収める必要がある
Private Const RSS_MAX_SYMBOLS As Long = 500
Private Const RSS_SAFETY_MARGIN As Long = 50   ' Watchlist側が使う分の安全マージン
Private Const BATCH_SIZE As Long = 450         ' 1バッチあたりの処理銘柄数

' JPX東証上場銘柄一覧（公式配布ファイル）のパス
' https://www.jpx.co.jp/markets/statistics-equities/misc/01.html からダウンロードしたもの
' 月次更新（毎月第3営業日）。更新されたら同じファイル名で上書き保存すればよい
'
' ユーザー名をコードに埋め込まないよう、USERPROFILE から組み立てる。
' 保存先を変えた場合はこの相対パスだけを直せばよい。
Private Const JPX_RELATIVE_PATH As String = "\OneDrive\株・資産管理\銘柄表\data_j.xls"

' RSSの同時取得待ちタイムアウト（秒）
' 再接続直後は450銘柄の登録に時間がかかることがあるため、余裕を持たせている
Private Const RSS_WAIT_TIMEOUT_SEC As Long = 180

'------------------------------------------------------
' JPXファイルの絶対パスを組み立てて返す。
' Const は式を評価できないため関数にしている。
' 呼び出し側は従来どおり JPX_FILE_PATH と書けばよい。
'------------------------------------------------------
Private Function JPX_FILE_PATH() As String
    JPX_FILE_PATH = Environ$("USERPROFILE") & JPX_RELATIVE_PATH
End Function

'------------------------------------------------------
' ① 少数銘柄（10件程度）でRSS項目の単位・書式を検証する
'    本番実行の前に必ずこれで確認すること
'------------------------------------------------------
Sub BuildCandidateSheet_Test()
    Dim ws As Worksheet
    Dim testCodes As Variant
    Dim i As Long, r As Long

    ' 検証用サンプル銘柄（任意で変更可）
    testCodes = Array(7203, 6758, 9984, 8306, 4755, 6501, 7011, 4063, 6981, 9432)

    Set ws = GetOrCreateSheet("全銘柄候補_TEST")
    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("証券コード", "銘柄名称", "売買代金", "市場区分", "決算発表日", "備考")

    Application.ScreenUpdating = False

    r = 2
    For i = LBound(testCodes) To UBound(testCodes)
        ws.Cells(r, 1).Value = testCodes(i)
        ws.Cells(r, 2).Formula = "=RssMarket(A" & r & ",""銘柄名称"")"
        ws.Cells(r, 3).Formula = "=RssMarket(A" & r & ",""売買代金"")"
        ws.Cells(r, 4).Formula = "=RssMarket(A" & r & ",""市場部略称"")"
        ws.Cells(r, 5).Formula = "=RssMarket(A" & r & ",""決算発表日"")"
        r = r + 1
    Next i

    Application.ScreenUpdating = True

    MsgBox "テストシートを作成しました。" & vbCrLf & _
           "・C列（売買代金）の桁数から単位（円/千円）を確認" & vbCrLf & _
           "・E列（決算発表日）の表示形式を確認" & vbCrLf & _
           "これらを確認後、本番のBuildCandidateSheetを実行してください。", vbInformation
End Sub

'------------------------------------------------------
' ②-A JPX公式の東証上場銘柄一覧を読み込み、
'      プライム・スタンダード（内国株式）だけを抽出して
'      「JPX銘柄マスタ」シートに保存する
'      ※ 月次更新のファイルなので、月1回実行すればよい
'------------------------------------------------------
Sub ImportJPXMaster()
    Dim wbSrc As Workbook
    Dim shSrc As Worksheet
    Dim wsMaster As Worksheet
    Dim r As Long, outRow As Long
    Dim marketType As String
    Dim importedCount As Long

    If Dir(JPX_FILE_PATH) = "" Then
        MsgBox "JPXファイルが見つかりません。" & vbCrLf & JPX_FILE_PATH & vbCrLf & _
               "パスを確認してください。", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Set wbSrc = Workbooks.Open(FileName:=JPX_FILE_PATH, ReadOnly:=True, UpdateLinks:=0)
    Set shSrc = wbSrc.Sheets(1)

    Set wsMaster = GetOrCreateSheet("JPX銘柄マスタ")
    wsMaster.Cells.Clear
    wsMaster.Range("A1:C1").Value = Array("証券コード", "銘柄名称", "市場区分")

    outRow = 2
    r = 2
    Do While shSrc.Cells(r, 2).Value <> ""
        marketType = CStr(shSrc.Cells(r, 4).Value)
        If marketType = "プライム（内国株式）" Or marketType = "スタンダード（内国株式）" Then
            wsMaster.Cells(outRow, 1).Value = shSrc.Cells(r, 2).Value
            wsMaster.Cells(outRow, 2).Value = shSrc.Cells(r, 3).Value
            wsMaster.Cells(outRow, 3).Value = marketType
            outRow = outRow + 1
        End If
        r = r + 1
    Loop

    importedCount = outRow - 2

    wbSrc.Close SaveChanges:=False
    Application.ScreenUpdating = True

    MsgBox "JPX銘柄マスタを更新しました。" & vbCrLf & _
           "プライム＋スタンダード合計： " & importedCount & " 銘柄" & vbCrLf & _
           "必要バッチ数（" & BATCH_SIZE & "件区切り）： " & GetBatchCount() & " バッチ" & vbCrLf & vbCrLf & _
           "続けて BuildCandidateSheetBatch(1) から順に実行してください。", vbInformation
End Sub

'------------------------------------------------------
' JPX銘柄マスタの件数から、必要なバッチ数を計算する
'------------------------------------------------------
Function GetBatchCount() As Long
    Dim wsMaster As Worksheet
    Dim totalCount As Long
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets("JPX銘柄マスタ")
    On Error GoTo 0
    If wsMaster Is Nothing Then
        GetBatchCount = 0
        Exit Function
    End If
    totalCount = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).row - 1
    GetBatchCount = -Int(-totalCount / BATCH_SIZE)  ' 切り上げ除算
End Function

'------------------------------------------------------
' ②-B JPX銘柄マスタから指定バッチ番号（1始まり）の範囲を取り出し、
'      「全銘柄候補」シートにRSS数式（売買代金・決算発表日のみ）を仕込む
'      銘柄名・市場区分はマスタから直接転記（RSS不要）
'------------------------------------------------------
Sub BuildCandidateSheetBatch(batchNumber As Long, Optional silent As Boolean = False)
    Dim wsMaster As Worksheet, ws As Worksheet
    Dim totalCount As Long, startRow As Long, endRow As Long
    Dim r As Long, outRow As Long
    Dim batchCount As Long

    Set wsMaster = Nothing
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets("JPX銘柄マスタ")
    On Error GoTo 0
    If wsMaster Is Nothing Then
        MsgBox "「JPX銘柄マスタ」シートがありません。先に ImportJPXMaster を実行してください。", vbExclamation
        Exit Sub
    End If

    batchCount = GetBatchCount()
    If batchNumber < 1 Or batchNumber > batchCount Then
        MsgBox "バッチ番号は 1 ～ " & batchCount & " の範囲で指定してください。", vbExclamation
        Exit Sub
    End If

    totalCount = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).row - 1
    startRow = 2 + (batchNumber - 1) * BATCH_SIZE
    endRow = startRow + BATCH_SIZE - 1
    If endRow > totalCount + 1 Then endRow = totalCount + 1

    Set ws = GetOrCreateSheet("全銘柄候補")
    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("証券コード", "銘柄名称", "売買代金", "市場区分", "決算発表日", "条件合致")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    outRow = 2
    For r = startRow To endRow
        ws.Cells(outRow, 1).Value = wsMaster.Cells(r, 1).Value
        ws.Cells(outRow, 2).Value = wsMaster.Cells(r, 2).Value
        ws.Cells(outRow, 3).Formula = "=RssMarket(A" & outRow & ",""売買代金"")"
        ws.Cells(outRow, 4).Value = wsMaster.Cells(r, 3).Value
        ws.Cells(outRow, 5).Formula = "=RssMarket(A" & outRow & ",""決算発表日"")"
        outRow = outRow + 1
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If Not silent Then
        MsgBox "バッチ " & batchNumber & " / " & batchCount & " を処理しました（" & _
           (endRow - startRow + 1) & " 銘柄）。" & vbCrLf & _
           "RSSの値が更新されたら FilterCandidates を実行してください。" & vbCrLf & _
           "次のバッチは BuildCandidateSheetBatch(" & (batchNumber + 1) & ") です。", vbInformation
    End If
End Sub

'------------------------------------------------------
' ②-C 全バッチを自動で連続実行する（推奨・最も手軽）
'      各バッチごとに RSSの反映待ち→表示銘柄数超過チェック
'      →FilterCandidates を自動で行う
'      使い方： Call RunAllBatches         （バッチ1から最後まで）
'              Call RunAllBatches(4)      （バッチ4から再開したい場合）
'------------------------------------------------------
Sub RunAllBatches(Optional startBatch As Long = 1)
    Dim batchCount As Long, i As Long
    Dim waitResult As String

    batchCount = GetBatchCount()
    If batchCount = 0 Then
        MsgBox "「JPX銘柄マスタ」がありません。先に ImportJPXMaster を実行してください。", vbExclamation
        Exit Sub
    End If

    If startBatch < 1 Or startBatch > batchCount Then
        MsgBox "開始バッチ番号は 1 ～ " & batchCount & " の範囲で指定してください。", vbExclamation
        Exit Sub
    End If

    For i = startBatch To batchCount
        Call BuildCandidateSheetBatch(i, silent:=True)
        waitResult = WaitForRSSReady()

        If waitResult = "OVERFLOW" Then
            MsgBox "バッチ " & i & " で「表示銘柄数超過」が発生しました。" & vbCrLf & vbCrLf & _
                   "処理をここで停止します。以下の手順で再開してください。" & vbCrLf & _
                   "1. このダイアログを閉じる（OKを押す）" & vbCrLf & _
                   "2. マーケットスピードIIタブの接続アイコンで一度切断→再接続する" & vbCrLf & _
                   "3. イミディエイトウィンドウで Call RunAllBatches(" & i & ") を実行する", vbExclamation
            Exit Sub
        ElseIf waitResult = "TIMEOUT" Then
            MsgBox "バッチ " & i & " のRSS取得が" & RSS_WAIT_TIMEOUT_SEC & "秒待っても完了しませんでした。" & vbCrLf & vbCrLf & _
                   "処理をここで停止します。接続状態を確認の上、" & vbCrLf & _
                   "Call RunAllBatches(" & i & ") で再開してください。", vbCritical
            Exit Sub
        End If

        Call FilterCandidates
    Next i

    MsgBox "全 " & batchCount & " バッチの処理が完了しました。" & vbCrLf & _
           "Watchlist・コア30候補一覧・イベント銘柄一覧をご確認ください。", vbInformation
End Sub

'------------------------------------------------------
' 「全銘柄候補」シートのRSS取得が完了するまで待機する
' 戻り値： "OK"（正常完了） / "OVERFLOW"（表示銘柄数超過を検知） / "TIMEOUT"（時間切れ）
'------------------------------------------------------
Private Function WaitForRSSReady() As String
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim startTime As Single
    Dim allDone As Boolean, hasOverflow As Boolean

    Set ws = ThisWorkbook.Sheets("全銘柄候補")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    startTime = Timer

    Do
        DoEvents
        Application.Calculate
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")

        allDone = True
        hasOverflow = False

        For r = 2 To lastRow
            If ws.Cells(r, 3).Value = "表示銘柄数超過" Or ws.Cells(r, 5).Value = "表示銘柄数超過" Then
                hasOverflow = True
                Exit For
            ElseIf ws.Cells(r, 3).Value = "" Then
                allDone = False
                Exit For
            End If
        Next r

        If hasOverflow Then
            WaitForRSSReady = "OVERFLOW"
            Exit Function
        End If
        If allDone Then
            WaitForRSSReady = "OK"
            Exit Function
        End If
    Loop While Timer - startTime < RSS_WAIT_TIMEOUT_SEC

    ' タイムアウト直前に、念のためもう一度だけ強制チェック
    ' （ポップアップ表示など何らかのイベント処理をきっかけにRSS更新が
    '   反映されるケースがあるため、最後にもう一押しする）
    DoEvents
    Application.Calculate
    DoEvents
    Application.Wait Now + TimeValue("0:00:02")
    Application.Calculate
    DoEvents

    allDone = True
    hasOverflow = False
    For r = 2 To lastRow
        If ws.Cells(r, 3).Value = "表示銘柄数超過" Or ws.Cells(r, 5).Value = "表示銘柄数超過" Then
            hasOverflow = True
            Exit For
        ElseIf ws.Cells(r, 3).Value = "" Then
            allDone = False
            Exit For
        End If
    Next r

    If hasOverflow Then
        WaitForRSSReady = "OVERFLOW"
        Exit Function
    End If
    If allDone Then
        WaitForRSSReady = "OK"
        Exit Function
    End If

    WaitForRSSReady = "TIMEOUT"
End Function

'------------------------------------------------------
' コア30候補一覧・イベント銘柄一覧に既に入ってしまった重複行を
' 掃除する（1回だけ実行すればよい）
'------------------------------------------------------
Sub CleanupDuplicatesInReferenceSheets()
    Call RemoveDuplicateRows("コア30候補一覧")
    Call RemoveDuplicateRows("イベント銘柄一覧")
    MsgBox "重複行の掃除が完了しました。", vbInformation
End Sub

Private Sub RemoveDuplicateRows(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long, r As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Application.ScreenUpdating = False
    For r = lastRow To 2 Step -1
        Dim c As String
        c = CStr(ws.Cells(r, 1).Value)
        If c = "" Then
            ' 何もしない
        ElseIf seen.Exists(c) Then
            ws.Rows(r).Delete
        Else
            seen.Add c, True
        End If
    Next r
    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------
' ②-legacy 手動でコード範囲を直接指定する版（従来方式）
'    JPXマスタが使えない場合の予備手段として残置
'    重い処理のため、監視時間帯（8:59-10:00）は避けること
'------------------------------------------------------
Sub BuildCandidateSheet(Optional startCode As Long = 1300, _
                        Optional endCode As Long = 9999)
    Dim ws As Worksheet
    Dim r As Long, c As Long
    Dim rangeSize As Long
    Dim allowedMax As Long

    rangeSize = endCode - startCode + 1
    allowedMax = RSS_MAX_SYMBOLS - RSS_SAFETY_MARGIN

    If rangeSize > allowedMax Then
        MsgBox "指定範囲が " & rangeSize & " 銘柄あります。" & vbCrLf & _
               "RSSの同時取得上限は500銘柄（証券口座単位・Watchlist分を含む）のため、" & vbCrLf & _
               "一度に処理できるのは最大 " & allowedMax & " 銘柄までに制限しています。" & vbCrLf & _
               "範囲を分けて複数回実行してください（例：500ずつ）。", vbExclamation
        Exit Sub
    End If

    Set ws = GetOrCreateSheet("全銘柄候補")
    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("証券コード", "銘柄名称", "売買代金", "市場区分", "決算発表日", "条件合致")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    r = 2
    For c = startCode To endCode
        ws.Cells(r, 1).Value = c
        ws.Cells(r, 2).Formula = "=RssMarket(A" & r & ",""銘柄名称"")"
        ws.Cells(r, 3).Formula = "=RssMarket(A" & r & ",""売買代金"")"
        ws.Cells(r, 4).Formula = "=RssMarket(A" & r & ",""市場部略称"")"
        ws.Cells(r, 5).Formula = "=RssMarket(A" & r & ",""決算発表日"")"
        r = r + 1
    Next c

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "全銘柄候補シートに " & (endCode - startCode + 1) & " 件分の数式を仕込みました。" & vbCrLf & _
           "RSSの値が更新されるまで少し待ってから FilterCandidates を実行してください。", vbInformation
End Sub

'------------------------------------------------------
' ③ 条件に合う銘柄を判定し、Watchlistへ転記する
'    ・Watchlist：監視対象。コア30候補／イベント銘柄の両方を追加
'      （J列に「選定理由」を記載。A～I列は既存の監視用に温存）
'    ・コア30候補一覧／イベント銘柄一覧：参照専用シート（監視対象外）
'------------------------------------------------------
Sub FilterCandidates()
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim wsCore As Worksheet, wsEvent As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long
    Dim baibaidaikin As Variant, shijou As Variant, kessan As Variant
    Dim isCoreCandidate As Boolean, isEventCandidate As Boolean
    Dim todayD As Date, tomorrowD As Date
    Dim addedCount As Long
    Dim reasonText As String

    Set wsSrc = ThisWorkbook.Sheets("全銘柄候補")
    Set wsDst = ThisWorkbook.Sheets("Watchlist")
    Set wsCore = GetOrCreateSheet("コア30候補一覧")
    Set wsEvent = GetOrCreateSheet("イベント銘柄一覧")

    ' Watchlist J列の見出し（無ければ追加）
    If wsDst.Cells(1, 10).Value = "" Then
        wsDst.Cells(1, 10).Value = "選定理由"
    End If

    ' 参照専用シート：初回のみ見出しを作成（毎回クリアせず、バッチをまたいで累積する）
    If wsCore.Cells(1, 1).Value = "" Then
        wsCore.Range("A1:E1").Value = Array("証券コード", "銘柄名称", "売買代金(千円)", "市場区分", "決算発表日")
    End If
    If wsEvent.Cells(1, 1).Value = "" Then
        wsEvent.Range("A1:E1").Value = Array("証券コード", "銘柄名称", "売買代金(千円)", "市場区分", "決算発表日")
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row
    outRow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).row + 1
    If outRow < 2 Then outRow = 2
    todayD = Date
    tomorrowD = Date + 1
    addedCount = 0

    Dim coreRow As Long, eventRow As Long
    coreRow = wsCore.Cells(wsCore.Rows.Count, 1).End(xlUp).row + 1
    If coreRow < 2 Then coreRow = 2
    eventRow = wsEvent.Cells(wsEvent.Rows.Count, 1).End(xlUp).row + 1
    If eventRow < 2 Then eventRow = 2
    Dim coreAddedThisRun As Long, eventAddedThisRun As Long
    coreAddedThisRun = 0
    eventAddedThisRun = 0

    ' Watchlistに既にある証券コードを収集（重複追加を防ぐため）
    Dim existingCodes As Object
    Set existingCodes = CreateObject("Scripting.Dictionary")
    Dim wr As Long, wLastRow As Long
    wLastRow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).row
    For wr = 2 To wLastRow
        If wsDst.Cells(wr, 1).Value <> "" Then
            If Not existingCodes.Exists(CStr(wsDst.Cells(wr, 1).Value)) Then
                existingCodes.Add CStr(wsDst.Cells(wr, 1).Value), wr
            End If
        End If
    Next wr

    Dim skippedDupCount As Long
    skippedDupCount = 0

    ' コア30候補一覧・イベント銘柄一覧の既存コードも収集（同一データへの再実行による重複を防ぐ）
    Dim existingCoreCodes As Object, existingEventCodes As Object
    Set existingCoreCodes = CreateObject("Scripting.Dictionary")
    Set existingEventCodes = CreateObject("Scripting.Dictionary")
    Dim cr As Long, crLastRow As Long
    crLastRow = wsCore.Cells(wsCore.Rows.Count, 1).End(xlUp).row
    For cr = 2 To crLastRow
        If wsCore.Cells(cr, 1).Value <> "" Then
            If Not existingCoreCodes.Exists(CStr(wsCore.Cells(cr, 1).Value)) Then
                existingCoreCodes.Add CStr(wsCore.Cells(cr, 1).Value), True
            End If
        End If
    Next cr
    Dim er As Long, erLastRow As Long
    erLastRow = wsEvent.Cells(wsEvent.Rows.Count, 1).End(xlUp).row
    For er = 2 To erLastRow
        If wsEvent.Cells(er, 1).Value <> "" Then
            If Not existingEventCodes.Exists(CStr(wsEvent.Cells(er, 1).Value)) Then
                existingEventCodes.Add CStr(wsEvent.Cells(er, 1).Value), True
            End If
        End If
    Next er

    For r = 2 To lastRow
        baibaidaikin = wsSrc.Cells(r, 3).Value
        shijou = wsSrc.Cells(r, 4).Value
        kessan = wsSrc.Cells(r, 5).Value

        isCoreCandidate = False
        If IsNumeric(baibaidaikin) Then
            If baibaidaikin >= BAIBAI_MIN And baibaidaikin <= BAIBAI_MAX Then
                ' JPXマスタ方式（"プライム（内国株式）"等）とRSS方式（"東P"等）の両方に対応
                If InStr(shijou, "プライム") > 0 Or InStr(shijou, "スタンダード") > 0 _
                   Or shijou = "東P" Or shijou = "東S" Then
                    isCoreCandidate = True
                End If
            End If
        End If

        isEventCandidate = False
        If IsDate(kessan) Then
            If CDate(kessan) = todayD Or CDate(kessan) = tomorrowD Then
                isEventCandidate = True
            End If
        End If

        reasonText = ""
        If isCoreCandidate Then reasonText = "コア30候補"
        If isEventCandidate Then
            If reasonText <> "" Then
                reasonText = reasonText & "・イベント銘柄"
            Else
                reasonText = "イベント銘柄"
            End If
        End If

        wsSrc.Cells(r, 6).Value = reasonText

        If isCoreCandidate Or isEventCandidate Then
            Dim codeStr As String
            codeStr = CStr(wsSrc.Cells(r, 1).Value)

            If existingCodes.Exists(codeStr) Then
                ' 既にWatchlistにある銘柄は追加せず、選定理由だけ更新
                Dim existRow As Long
                existRow = existingCodes(codeStr)
                wsDst.Cells(existRow, 10).Value = reasonText
                skippedDupCount = skippedDupCount + 1
            Else
                ' Watchlist（監視対象）へ新規追加：A証券コード、B銘柄名、J選定理由
                wsDst.Cells(outRow, 1).Value = wsSrc.Cells(r, 1).Value
                wsDst.Cells(outRow, 2).Value = wsSrc.Cells(r, 2).Value
                wsDst.Cells(outRow, 10).Value = reasonText
                existingCodes.Add codeStr, outRow
                outRow = outRow + 1
                addedCount = addedCount + 1
            End If

            ' カテゴリ別の参照専用シートへも記録（累積。同一コードの再登録は防ぐ）
            Dim codeStrForRef As String
            codeStrForRef = CStr(wsSrc.Cells(r, 1).Value)

            If isCoreCandidate Then
                If Not existingCoreCodes.Exists(codeStrForRef) Then
                    wsCore.Cells(coreRow, 1).Value = wsSrc.Cells(r, 1).Value
                    wsCore.Cells(coreRow, 2).Value = wsSrc.Cells(r, 2).Value
                    wsCore.Cells(coreRow, 3).Value = wsSrc.Cells(r, 3).Value
                    wsCore.Cells(coreRow, 4).Value = wsSrc.Cells(r, 4).Value
                    wsCore.Cells(coreRow, 5).Value = wsSrc.Cells(r, 5).Value
                    coreRow = coreRow + 1
                    coreAddedThisRun = coreAddedThisRun + 1
                    existingCoreCodes.Add codeStrForRef, True
                End If
            End If
            If isEventCandidate Then
                If Not existingEventCodes.Exists(codeStrForRef) Then
                    wsEvent.Cells(eventRow, 1).Value = wsSrc.Cells(r, 1).Value
                    wsEvent.Cells(eventRow, 2).Value = wsSrc.Cells(r, 2).Value
                    wsEvent.Cells(eventRow, 3).Value = wsSrc.Cells(r, 3).Value
                    wsEvent.Cells(eventRow, 4).Value = wsSrc.Cells(r, 4).Value
                    wsEvent.Cells(eventRow, 5).Value = wsSrc.Cells(r, 5).Value
                    eventRow = eventRow + 1
                    eventAddedThisRun = eventAddedThisRun + 1
                    existingEventCodes.Add codeStrForRef, True
                End If
            End If
        End If
    Next r

    MsgBox "抽出完了。" & vbCrLf & _
           "Watchlistに新規追加： " & addedCount & " 件（J列に選定理由を記載）" & vbCrLf & _
           "既存銘柄のため選定理由のみ更新： " & skippedDupCount & " 件" & vbCrLf & _
           "コア30候補一覧：今回 " & coreAddedThisRun & " 件追加（累積 " & (coreRow - 2) & " 件）" & vbCrLf & _
           "イベント銘柄一覧：今回 " & eventAddedThisRun & " 件追加（累積 " & (eventRow - 2) & " 件）", vbInformation
End Sub

'------------------------------------------------------
' シート取得（無ければ新規作成）
'------------------------------------------------------
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function
