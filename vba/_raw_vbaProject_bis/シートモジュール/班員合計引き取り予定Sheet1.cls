VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()

    Dim response As VbMsgBoxResult
    ' メッセージボックスを表示して確認を求める
    response = MsgBox("転写を実行します。よろしいですか？", vbYesNo + vbQuestion, "確認")
    
    ' ユーザーが「はい」を選択した場合に処理を実行
    If response = vbYes Then
        Call TransferDataByProductNumber
    End If
    
End Sub

Private Sub CommandButton2_Click()

    Dim response As VbMsgBoxResult
    ' メッセージボックスを表示して確認を求める
    response = MsgBox("ファイルを保存します。よろしいですか？", vbYesNo + vbQuestion, "確認")
    
    ' ユーザーが「はい」を選択した場合に処理を実行
    If response = vbYes Then
        Call SaveSheetsAsFiles
    End If
    
End Sub
