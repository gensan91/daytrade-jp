Attribute VB_Name = "modNotify"
Option Explicit

' ============================================================
' modNotify
' Q13: 通知音（音+ポップアップ）
' VBA標準のBeepステートメントは、内蔵PCスピーカー（マザーボード
' 直結のビープ）を鳴らす古い命令で、近年のノートPCや外部スピーカー
' 環境では物理的に音が出ないことが多い。
' そこでWindows APIのPlaySoundを使い、システム標準のwavファイルを
' 実際のサウンドカード経由で鳴らす（こちらは現代のPC環境でも確実）。
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
#End If

Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000

' ------------------------------------------------------------
' times回、通知音を鳴らす。system\media配下のwavが見つからない
' 場合は、フォールバックとして従来のBeepを使う。
' ------------------------------------------------------------
Sub PlayAlertSound(Optional times As Integer = 1)
    Dim soundFile As String
    ' Windows標準搭載の通知音（環境によって多少ファイル名が異なる場合あり）
    soundFile = Environ$("WINDIR") & "\Media\Windows Notify System Generic.wav"

    If Dir(soundFile) = "" Then
        ' 見つからない場合、別の標準サウンドを試す
        soundFile = Environ$("WINDIR") & "\Media\Alarm01.wav"
    End If

    Dim i As Integer
    If Dir(soundFile) <> "" Then
        For i = 1 To times
            PlaySound soundFile, 0, SND_FILENAME Or SND_ASYNC
            WaitShort 0.4
        Next i
    Else
        ' wavファイルが一切見つからない場合の最終フォールバック
        For i = 1 To times
            Beep
            WaitShort 0.2
        Next i
    End If
End Sub

Private Sub WaitShort(seconds As Single)
    Dim t As Single
    t = Timer
    Do While Timer < t + seconds
        DoEvents
    Loop
End Sub
