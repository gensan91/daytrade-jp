Attribute VB_Name = "modUIButtons"
Option Explicit

' ============================================================
' modUIButtons
' CommandBars方式は本環境で座標指定が効かない不具合が確認できたため撤去。
' 代わりに frmController（UserForm）をモードレス表示する方式に変更。
'
' 前提：frmController という名前のUserFormを手動で作成し、
'       CommandButton×8個とそれぞれのClickイベントコードを仕込んでおくこと。
' ============================================================

Sub ShowController()
    On Error Resume Next
    Unload frmController   ' 既に開いていたら一旦閉じて作り直す（多重表示防止）
    On Error GoTo 0

    frmController.Show vbModeless
End Sub
