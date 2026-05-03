- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private m_tbl        As ListObject
Private m_col担当者   As Long
Private m_col現場     As Long
Private m_dicBukken  As Object   ' 担当者名 → 物件名Collection

'--------------------------------------------------
' 呼び出し元から初期データを受け取る
'--------------------------------------------------
Public Sub Init(tbl As ListObject, col担当者 As Long, col現場 As Long)
    Set m_tbl = tbl
    m_col担当者 = col担当者
    m_col現場 = col現場
End Sub

Private Sub ListBox1_Click()

End Sub

'--------------------------------------------------
' フォーム表示時：担当者リストを構築
'--------------------------------------------------
    Private Sub UserForm_Activate()
    Dim dicTanto As Object
    Set dicTanto = CreateObject("Scripting.Dictionary")
    Set m_dicBukken = CreateObject("Scripting.Dictionary")

    Dim lr  As ListRow
    Dim nm  As String, pn As String

    For Each lr In m_tbl.ListRows
        nm = Trim(CStr(lr.Range(m_col担当者).Value))
        pn = Trim(CStr(lr.Range(m_col現場).Value))

        If nm <> "" Then
            If Not dicTanto.Exists(nm) Then
                dicTanto.Add nm, 1
                lstTanto.AddItem nm
                m_dicBukken.Add nm, New Collection
            End If
            If pn <> "" Then
                Dim col As Collection
                Set col = m_dicBukken(nm)
                Dim already As Boolean: already = False
                Dim itm As Variant
                For Each itm In col
                    If CStr(itm) = pn Then already = True: Exit For
                Next itm
                If Not already Then col.Add pn
            End If
        End If
    Next lr

    Me.Caption = "見積削除"
End Sub

'--------------------------------------------------
' 担当者を選択 → 物件リストを更新
'--------------------------------------------------
Private Sub lstTanto_Click()
    lstBukken.Clear
    btnDelete.Enabled = False

    Dim selected As String
    selected = lstTanto.Value
    If selected = "" Then Exit Sub

    lstBukken.AddItem "【すべて削除】"

    Dim col As Collection
    Set col = m_dicBukken(selected)
    Dim itm As Variant
    For Each itm In col
        lstBukken.AddItem CStr(itm)
    Next itm

    lstBukken.Enabled = True
End Sub

'--------------------------------------------------
' 物件を選択 → 削除ボタン有効化
'--------------------------------------------------
Private Sub lstBukken_Click()
    btnDelete.Enabled = (lstBukken.ListIndex >= 0)
End Sub

'--------------------------------------------------
' 削除実行
'--------------------------------------------------
Private Sub btnDelete_Click()
    Dim 削除担当者 As String
    削除担当者 = lstTanto.Value

    If lstBukken.ListIndex < 0 Then
        MsgBox "物件を選択してください。", vbExclamation: Exit Sub
    End If

    Dim bAll As Boolean
    bAll = (lstBukken.ListIndex = 0)   ' 先頭「すべて削除」

    Dim 削除物件 As String
    If Not bAll Then 削除物件 = lstBukken.Value

    Dim confirmMsg As String
    If bAll Then
        confirmMsg = "【" & 削除担当者 & "】のデータをすべて削除します。よろしいですか？"
    Else
        confirmMsg = "【" & 削除担当者 & " / " & 削除物件 & "】を削除します。よろしいですか？"
    End If
    If MsgBox(confirmMsg, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Dim r           As Long
    Dim deleteCount As Long
    Dim dicPhysical As Object: Set dicPhysical = CreateObject("Scripting.Dictionary")
    Dim tgtTanto As String, tgtBukken As String

    Application.ScreenUpdating = False

    For r = m_tbl.ListRows.count To 1 Step -1
        tgtTanto = Trim(CStr(m_tbl.ListRows(r).Range(m_col担当者).Value))
        tgtBukken = Trim(CStr(m_tbl.ListRows(r).Range(m_col現場).Value))
        If tgtTanto = 削除担当者 Then
            If bAll Or tgtBukken = 削除物件 Then
                If tgtBukken <> "" And Not dicPhysical.Exists(tgtBukken) Then dicPhysical.Add tgtBukken, 1
                m_tbl.ListRows(r).Delete
                deleteCount = deleteCount + 1
            End If
        End If
    Next r

    Dim delKey   As Variant
    Dim delSheet As Worksheet
    Application.DisplayAlerts = False
    For Each delKey In dicPhysical.Keys
        Set delSheet = Nothing
        On Error Resume Next
        Set delSheet = ThisWorkbook.Sheets(シート名サニタイズ(CStr(delKey)))
        On Error GoTo 0
        If Not delSheet Is Nothing Then delSheet.Delete
    Next delKey
    Application.DisplayAlerts = True

    ログから削除 削除担当者, IIf(bAll, "", 削除物件)

    Application.ScreenUpdating = True
    MsgBox deleteCount & " 行削除しました。", vbInformation
    Unload Me
End Sub

'--------------------------------------------------
' キャンセル
'--------------------------------------------------
Private Sub btnCancel_Click()
    Unload Me
End Sub