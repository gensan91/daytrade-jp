Attribute VB_Name = "modDaily7"
Option Explicit

'======================================================
' modDaily7（modDaily6の置き換え。旧modDaily6は解放してください）
'
' 【検証で確定したRssChartPastの仕様】公式ヘルプに記載のない挙動を含む
'   1. スピル関数ではない。数式セルはステータス文字列（"完了"等）を返す
'   2. ヘッダー行引数は「出力先」ではなく「入力」。省略するとRSSが
'      全項目のヘッダーを数式セルの1行下に自動で書く → 省略が正解
'   3. データはヘッダーのさらに下（数式セルの2行下）から始まる
'   4. 起点日付から【未来方向】へ指定本数ぶん取得する。並びは古い順
'      → 当日を起点にすると1本しか返らない（最初の失敗の原因）
'   5. 起点が非営業日なら翌営業日から始まる（エラーにはならない）
'   6. 列： 1銘柄名称 2市場名称 3足種 4日付 5時刻 6始値 7高値 8安値 9終値 10出来高
'   7. 銘柄名称は最初のデータ行のみ。2行目以降は空欄
'      → 行数を数えるときは【日付列】を見ること
'   8. データの末尾に "--------" の行が入る
'   9. 銘柄コード書式（7203 / 7203.T / 72030.T）は結果に影響しない
'
' 【modDaily6からの変更】
'   ・BuildCleanList を取り込んだ（modDaily6に入れ忘れていた）
'   ・中断時に Err.Number / Err.Description を表示するようにした
'   ・選別候補シートの有無を、エラー任せにせず明示的にチェックする
'
' 【手順】
'   1) Call BuildCleanList    … スナップショットから選別候補を作る（RSS不要）
'   2) Call TestOneStock      … 1銘柄だけ完走させて数値を目視確認（検証済み）
'   3) Call BuildDailyStats   … 選別候補の全銘柄（10並列）
'======================================================

Private Const BASE_DATE As String = "2026/08/08"
Private Const START_OFFSET_DAYS As Long = 45   ' 起点＝基準日の何日前にするか
Private Const BARS As Long = 40                ' 要求本数
Private Const USE_BARS As Long = 20            ' 集計に使う直近本数
Private Const MIN_ROWS As Long = 15            ' これだけ揃えば取得完了とみなす

Private Const CHUNK As Long = 10               ' 同時投入する銘柄数
Private Const BLOCK_W As Long = 11             ' 1銘柄あたりの列幅
Private Const BLOCK_H As Long = 48             ' 1銘柄あたりの行数
Private Const DATA_ROW As Long = 3             ' 数式が1行目 → データは3行目から
Private Const CHUNK_TIMEOUT_SEC As Long = 45

Private Const KESSAN_EXCLUDE_DAYS As Long = 7   ' 決算発表日がこの日数以内なら除外

Private Const SH_SNAP As String = "スナップショット"
Private Const SH_CLEAN As String = "選別候補"
Private Const SH_STATS As String = "日足集計"
Private Const SH_WORK As String = "_work"


'------------------------------------------------------
' (0) スナップショットから選別候補を作る（RSS不要・即完了）
'    フィルタ： 信用貸借区分=1（空売り可） かつ 決算発表日が基準日±7日でない
'------------------------------------------------------
Sub BuildCleanList()
    Dim wsS As Worksheet, ws As Worksheet
    Dim lastRow As Long, r As Long, o As Long
    Dim baseD As Date, kessanD As Date
    Dim shinyo As Variant, kessanRaw As Variant
    Dim keep As Boolean, elapsed As Double
    Dim cntShinyo As Long, cntKessan As Long
    Dim genzai As Double, zenjitsu As Double, hi As Double, lo As Double, hajime As Double

    On Error Resume Next
    Set wsS = ThisWorkbook.Sheets(SH_SNAP)
    On Error GoTo 0
    If wsS Is Nothing Then
        MsgBox "「" & SH_SNAP & "」シートがありません。" & vbCrLf & _
               "先に modSnapshot の SnapshotWatchlist を実行してください。", vbCritical
        Exit Sub
    End If

    baseD = CDate(BASE_DATE)
    lastRow = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).row

    Set ws = GetSheet(SH_CLEAN)
    ws.Cells.Clear
    ws.Range("A1:K1").Value = Array("証券コード", "銘柄名", "現在値", "単元価格", _
        "売買代金(億)", "時価総額(億)", "値幅率(%)", "ギャップ率(%)", _
        "決算発表日", "決算経過日", "市場")

    Application.ScreenUpdating = False
    o = 2
    For r = 2 To lastRow
        If Trim(CStr(wsS.Cells(r, 1).Value)) <> "" Then
            keep = True
            shinyo = wsS.Cells(r, 12).Value
            kessanRaw = wsS.Cells(r, 13).Value

            If Not IsNumeric(shinyo) Then
                keep = False
            ElseIf CLng(shinyo) <> 1 Then
                keep = False: cntShinyo = cntShinyo + 1
            End If

            elapsed = 9999
            If keep Then
                If IsDate(kessanRaw) Then
                    kessanD = CDate(kessanRaw)
                    elapsed = DateDiff("d", kessanD, baseD)
                    If Abs(elapsed) <= KESSAN_EXCLUDE_DAYS Then
                        keep = False: cntKessan = cntKessan + 1
                    End If
                End If
            End If

            If keep Then
                genzai = SafeD(wsS.Cells(r, 3).Value)
                zenjitsu = SafeD(wsS.Cells(r, 4).Value)
                hajime = SafeD(wsS.Cells(r, 5).Value)
                hi = SafeD(wsS.Cells(r, 6).Value)
                lo = SafeD(wsS.Cells(r, 7).Value)

                ws.Cells(o, 1).Value = wsS.Cells(r, 1).Value
                ws.Cells(o, 2).Value = wsS.Cells(r, 2).Value
                ws.Cells(o, 3).Value = genzai
                ws.Cells(o, 4).Value = genzai * SafeD(wsS.Cells(r, 11).Value)
                ws.Cells(o, 5).Value = Round(SafeD(wsS.Cells(r, 9).Value) / 100000, 2)
                ws.Cells(o, 6).Value = Round(SafeD(wsS.Cells(r, 10).Value) / 100, 0)
                If zenjitsu > 0 Then
                    ws.Cells(o, 7).Value = Round((hi - lo) / zenjitsu * 100, 2)
                    ws.Cells(o, 8).Value = Round((hajime - zenjitsu) / zenjitsu * 100, 2)
                End If
                ws.Cells(o, 9).Value = kessanRaw
                If elapsed < 9999 Then ws.Cells(o, 10).Value = elapsed
                ws.Cells(o, 11).Value = wsS.Cells(r, 14).Value
                o = o + 1
            End If
        End If
    Next r
    Application.ScreenUpdating = True

    MsgBox "選別候補を作成しました。" & vbCrLf & vbCrLf & _
           "対象： " & (lastRow - 1) & " 銘柄" & vbCrLf & _
           "除外： 信用貸借区分≠1 … " & cntShinyo & " 件" & vbCrLf & _
           "　　　 決算±" & KESSAN_EXCLUDE_DAYS & "日 … " & cntKessan & " 件" & vbCrLf & _
           "残り： " & (o - 2) & " 銘柄", vbInformation
End Sub

'------------------------------------------------------
' ① まず1銘柄だけ。取得から集計まで通して数値を目で見る
'------------------------------------------------------
Sub TestOneStock()
    Dim ws As Worksheet
    Dim t As Single
    Dim n As Long
    Dim dat As Variant
    Dim s As String

    Set ws = GetSheet(SH_WORK)
    ws.Visible = xlSheetVisible
    ws.Activate
    ClearWork ws

    ws.Cells(1, 1).Formula = ChartFormula("7203")

    t = Timer
    Do
        DoEvents
        Application.Calculate
        n = CountRows(ws, 1)
        If n >= MIN_ROWS Then Exit Do
    Loop While Timer - t < 30

    If n < 2 Then
        ws.Activate
        MsgBox "取得できませんでした（" & n & " 行）。" & vbCrLf & _
               "シートは残してあります。A1の表示内容をご連絡ください。", vbCritical
        Exit Sub
    End If

    dat = ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(DATA_ROW + n - 1, 10)).Value
    s = CalcStats(dat)

    ws.Activate
    MsgBox "【1銘柄テスト】トヨタ自動車(7203)" & vbCrLf & vbCrLf & _
           "起点： " & Format(StartDate(), "yyyy/mm/dd") & _
           "（基準日 " & BASE_DATE & " の " & START_OFFSET_DAYS & " 日前）" & vbCrLf & _
           "取得本数： " & n & " 本（所要 " & Format(Timer - t, "0.0") & " 秒）" & vbCrLf & vbCrLf & _
           s & vbCrLf & vbCrLf & _
           "数値が現実的なら Call BuildDailyStats に進んでください。" & vbCrLf & _
           "（シートは残してあります。生データも確認できます）", vbInformation
End Sub

'------------------------------------------------------
' ② 本番：選別候補の全銘柄。10銘柄を横並びで同時投入
'    途中再開： Call BuildDailyStats(51)
'------------------------------------------------------
Sub BuildDailyStats(Optional startIdx As Long = 1)
    Dim wsC As Worksheet, wsW As Worksheet, wsO As Worksheet
    Dim lastRow As Long, total As Long
    Dim i As Long, j As Long, k As Long
    Dim codes() As String, nms() As String
    Dim t0 As Single, tc As Single
    Dim nOK As Long, nNG As Long
    Dim ready As Long, n As Long, baseCol As Long
    Dim dat As Variant

    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Aborted

    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets(SH_CLEAN)
    On Error GoTo Aborted
    If wsC Is Nothing Then
        MsgBox "「" & SH_CLEAN & "」シートがありません。" & vbCrLf & vbCrLf & _
               "先に Call BuildCleanList を実行してください。", vbCritical
        Exit Sub
    End If
    lastRow = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).row
    total = lastRow - 1
    If total < 1 Then MsgBox "「" & SH_CLEAN & "」に銘柄がありません。", vbExclamation: Exit Sub

    Set wsW = GetSheet(SH_WORK)
    wsW.Visible = xlSheetVisible
    Set wsO = GetSheet(SH_STATS)

    If startIdx <= 1 Then
        wsO.Cells.Clear
        wsO.Range("A1:M1").Value = Array("No", "証券コード", "銘柄名", "本数", _
            "20日平均売買代金(億)", "20日平均値幅率(%)", "値幅率の標準偏差", _
            "20日平均|ギャップ|(%)", "ギャップ3%以上の日数", "ギャップ5%以上の日数", _
            "最大値幅率(%)", "最新終値", "備考")
    End If

    ReDim codes(1 To CHUNK): ReDim nms(1 To CHUNK)
    Application.ScreenUpdating = False
    t0 = Timer
    i = startIdx

    Do While i <= total
        k = 0
        For j = i To Application.Min(i + CHUNK - 1, total)
            k = k + 1
            codes(k) = Trim(CStr(wsC.Cells(j + 1, 1).Value))
            nms(k) = CStr(wsC.Cells(j + 1, 2).Value)
        Next j

        ClearWork wsW
        For j = 1 To k
            wsW.Cells(1, 1 + (j - 1) * BLOCK_W).Formula = ChartFormula(codes(j))
        Next j

        tc = Timer
        Do
            DoEvents
            Application.Calculate
            ready = 0
            For j = 1 To k
                If CountRows(wsW, 1 + (j - 1) * BLOCK_W) >= MIN_ROWS Then ready = ready + 1
            Next j
            Application.StatusBar = "日足取得中 " & (i + k - 1) & "/" & total & _
                "  チャンク " & ready & "/" & k & "  経過 " & Int(Timer - t0) & "秒  （Escで中断）"
            If ready >= k Then Exit Do
        Loop While Timer - tc < CHUNK_TIMEOUT_SEC

        For j = 1 To k
            baseCol = 1 + (j - 1) * BLOCK_W
            n = CountRows(wsW, baseCol)
            wsO.Cells(i + j, 1).Value = i + j - 1
            wsO.Cells(i + j, 2).Value = wsC.Cells(i + j, 1).Value
            wsO.Cells(i + j, 3).Value = nms(j)
            If n < 2 Then
                wsO.Cells(i + j, 13).Value = "取得失敗"
                nNG = nNG + 1
            Else
                dat = wsW.Range(wsW.Cells(DATA_ROW, baseCol), _
                                wsW.Cells(DATA_ROW + n - 1, baseCol + 9)).Value
                Call WriteStats(wsO, i + j, dat)
                nOK = nOK + 1
            End If
        Next j
        i = i + k
    Loop

    ClearWork wsW
    wsW.Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "日足集計が完了しました。" & vbCrLf & vbCrLf & _
           "成功： " & nOK & " 銘柄 ／ 失敗： " & nNG & " 銘柄" & vbCrLf & _
           "所要： " & Format((Timer - t0) / 60, "0.0") & " 分" & vbCrLf & vbCrLf & _
           "「" & SH_STATS & "」を確認して保存してください。", vbInformation
    Exit Sub

Aborted:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If Err.Number = 18 Then
        MsgBox "Escで中断しました（" & (i - 1) & " 銘柄まで完了）。" & vbCrLf & vbCrLf & _
               "再開： Call BuildDailyStats(" & Application.Max(1, i) & ")", vbExclamation
    Else
        MsgBox "エラーで中断しました。" & vbCrLf & vbCrLf & _
               "エラー番号： " & Err.Number & vbCrLf & _
               "内容　　　： " & Err.Description & vbCrLf & _
               "処理位置　： " & IIf(i = 0, "開始前（シート取得や初期化の段階）", i - 1 & " 銘柄まで完了") & vbCrLf & vbCrLf & _
               "この内容をそのままご連絡ください。", vbCritical
    End If
End Sub

'------------------------------------------------------
' 数式（ヘッダー行は省略する。これが正しい使い方）
'------------------------------------------------------
Private Function StartDate() As Date
    StartDate = CDate(BASE_DATE) - START_OFFSET_DAYS
End Function

Private Function ChartFormula(code As String) As String
    Dim arg As String
    arg = """" & code & ".T"""
    ChartFormula = "=RssChartPast(," & arg & ",""D""," & _
                   Format(StartDate(), "yyyymmdd") & "," & BARS & ")"
End Function

'------------------------------------------------------
' データ行数を数える（日付列＝基準列+3 を見る。銘柄名称列は空欄が多い）
'------------------------------------------------------
Private Function CountRows(ws As Worksheet, baseCol As Long) As Long
    Dim i As Long, c As Long
    For i = DATA_ROW To DATA_ROW + BARS - 1
        If IsDate(ws.Cells(i, baseCol + 3).Value) Then c = c + 1 Else Exit For
    Next i
    CountRows = c
End Function

Private Sub ClearWork(ws As Worksheet)
    ws.Range(ws.Cells(1, 1), ws.Cells(BLOCK_H, CHUNK * BLOCK_W + 2)).Clear
End Sub

'------------------------------------------------------
' 統計量の計算（datは古い順に並んでいる前提。念のため日付で確認する）
'------------------------------------------------------
Private Function Aggregate(dat As Variant, ByRef cnt As Long, ByRef avgT As Double, _
                           ByRef avgR As Double, ByRef sdR As Double, ByRef avgG As Double, _
                           ByRef g3 As Long, ByRef g5 As Long, ByRef maxR As Double, _
                           ByRef lastCl As Double) As Boolean
    Dim n As Long, i As Long, p As Long
    Dim asc_ As Boolean
    Dim op As Double, hi As Double, lo As Double, cl As Double, vol As Double, prevCl As Double
    Dim rng As Double, gap As Double
    Dim sumT As Double, sumR As Double, sumG As Double, sumSq As Double
    Dim idx() As Long

    n = UBound(dat, 1)
    If n < 2 Then Aggregate = False: Exit Function

    asc_ = True
    If IsDate(dat(1, 4)) And IsDate(dat(2, 4)) Then asc_ = (CDate(dat(1, 4)) < CDate(dat(2, 4)))
    ReDim idx(1 To n)
    For i = 1 To n
        If asc_ Then idx(i) = i Else idx(i) = n - i + 1
    Next i   ' idx は常に古い→新しい

    cnt = 0
    For i = n To 2 Step -1
        If cnt >= USE_BARS Then Exit For
        p = idx(i)
        op = SafeD(dat(p, 6)): hi = SafeD(dat(p, 7)): lo = SafeD(dat(p, 8))
        cl = SafeD(dat(p, 9)): vol = SafeD(dat(p, 10))
        prevCl = SafeD(dat(idx(i - 1), 9))
        If prevCl > 0 And hi > 0 Then
            rng = (hi - lo) / prevCl * 100
            gap = Abs(op - prevCl) / prevCl * 100
            sumT = sumT + vol * cl / 100000000#
            sumR = sumR + rng: sumSq = sumSq + rng * rng: sumG = sumG + gap
            If gap >= 3 Then g3 = g3 + 1
            If gap >= 5 Then g5 = g5 + 1
            If rng > maxR Then maxR = rng
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then Aggregate = False: Exit Function
    avgT = sumT / cnt
    avgR = sumR / cnt
    sdR = Sqr(Application.Max(0, sumSq / cnt - avgR * avgR))
    avgG = sumG / cnt
    lastCl = SafeD(dat(idx(n), 9))
    Aggregate = True
End Function

Private Sub WriteStats(wsO As Worksheet, o As Long, dat As Variant)
    Dim cnt As Long, g3 As Long, g5 As Long
    Dim avgT As Double, avgR As Double, sdR As Double, avgG As Double
    Dim maxR As Double, lastCl As Double

    If Not Aggregate(dat, cnt, avgT, avgR, sdR, avgG, g3, g5, maxR, lastCl) Then
        wsO.Cells(o, 13).Value = "計算不能": Exit Sub
    End If
    wsO.Cells(o, 4).Value = cnt
    wsO.Cells(o, 5).Value = Round(avgT, 2)
    wsO.Cells(o, 6).Value = Round(avgR, 2)
    wsO.Cells(o, 7).Value = Round(sdR, 2)
    wsO.Cells(o, 8).Value = Round(avgG, 2)
    wsO.Cells(o, 9).Value = g3
    wsO.Cells(o, 10).Value = g5
    wsO.Cells(o, 11).Value = Round(maxR, 2)
    wsO.Cells(o, 12).Value = lastCl
    If cnt < USE_BARS Then wsO.Cells(o, 13).Value = "本数不足(" & cnt & ")"
End Sub

Private Function CalcStats(dat As Variant) As String
    Dim cnt As Long, g3 As Long, g5 As Long
    Dim avgT As Double, avgR As Double, sdR As Double, avgG As Double
    Dim maxR As Double, lastCl As Double

    If Not Aggregate(dat, cnt, avgT, avgR, sdR, avgG, g3, g5, maxR, lastCl) Then
        CalcStats = "集計できませんでした。": Exit Function
    End If
    CalcStats = _
        "集計に使った本数： " & cnt & " 本" & vbCrLf & _
        "20日平均売買代金： " & Format(avgT, "0.0") & " 億円" & vbCrLf & _
        "20日平均値幅率　： " & Format(avgR, "0.00") & " ％" & vbCrLf & _
        "値幅率の標準偏差： " & Format(sdR, "0.00") & vbCrLf & _
        "20日平均|ギャップ|： " & Format(avgG, "0.00") & " ％" & vbCrLf & _
        "ギャップ3%以上　： " & g3 & " 日 ／ 5%以上： " & g5 & " 日" & vbCrLf & _
        "最大値幅率　　　： " & Format(maxR, "0.00") & " ％" & vbCrLf & _
        "最新終値　　　　： " & lastCl
End Function

Private Function SafeD(v As Variant) As Double
    If IsNumeric(v) Then SafeD = CDbl(v) Else SafeD = 0
End Function

Private Function GetSheet(nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nm
    End If
    Set GetSheet = ws
End Function
