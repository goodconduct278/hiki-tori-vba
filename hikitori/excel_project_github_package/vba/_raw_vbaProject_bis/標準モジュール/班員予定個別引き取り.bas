Attribute VB_Name = "Module1"
' 見積取込マクロ v10（UR列自動付与対応版）
'====================================================

Option Explicit

' ========= 共通定数 =========
Private Const TABLE_NAME     As String = "集計テーブル"
Private Const LOG_SHEET      As String = "取込ログ"
Private Const WARN_SHEET     As String = "単位不一致ログ"
Private Const SHEET_NAME_MAX As Long = 28

' ========= UR印 =========
' 集計テーブルに「UR」列がある場合、UR用見積書から取り込んだ行に「UR」を入れる。
' 通常見積書から取り込んだ行は空欄にする。

' ========= フォルダ保存用プロパティ名 =========
Private Const FOLDER_PROP_NORMAL As String = "見積書フォルダパス_通常"
Private Const FOLDER_PROP_UR     As String = "見積書フォルダパス_UR"

' ========= 通常見積書 レイアウト =========
Private Const NORMAL_CELL_担当者 As String = "B1"
Private Const NORMAL_CELL_物件名 As String = "F10"
Private Const NORMAL_COL_材料    As Long = 1    ' A列
Private Const NORMAL_COL_数量    As Long = 13   ' M列
Private Const NORMAL_COL_単位    As Long = 17   ' Q列

' ========= UR用見積書 レイアウト =========
Private Const UR_CELL_担当者 As String = "B1"
Private Const UR_CELL_物件名 As String = "H11"
Private Const UR_COL_材料    As Long = 3    ' C列
Private Const UR_COL_数量    As Long = 10   ' J列
Private Const UR_COL_単位    As Long = 13   ' M列

'====================================================
' 明細の読み込み範囲（開始行, 終了行）
' ★【修正①】開始行を 22 → 21 に修正（File2 の startRow=21 に合致）
' ★【修正②】NormalブロックとURブロックは現時点で同一レイアウト（バグ予防のため意図的に分離）
'            将来レイアウトが変わった場合はそれぞれ個別に修正すること
'====================================================
Private Function GetNormalBlocks() As Variant
    GetNormalBlocks = Array( _
        Array(21, 46), _
        Array(70, 94), _
        Array(118, 142), _
        Array(166, 190), _
        Array(214, 238) _
    )
End Function

Private Function GetURBlocks() As Variant
    ' 現時点では通常見積書と同一レイアウト（意図的分離）
    GetURBlocks = Array( _
        Array(21, 46), _
        Array(70, 94), _
        Array(118, 142), _
        Array(166, 190), _
        Array(214, 238) _
    )
End Function


'====================================================
' 初回セットアップ（ボタン用）
'====================================================
Sub 見積書フォルダ設定()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "見積書フォルダを作成する場所（親フォルダ）を選択してください"
    If fd.Show = False Then Exit Sub

    Dim parentPath As String, pathNormal As String, pathUR As String
    parentPath = fd.SelectedItems(1)

    pathNormal = parentPath & "\通常見積書"
    On Error Resume Next: If Dir(pathNormal, vbDirectory) = "" Then MkDir pathNormal: On Error GoTo 0
    フォルダパス保存 FOLDER_PROP_NORMAL, pathNormal

    If MsgBox("「UR用見積書」も取り込みますか？", vbQuestion + vbYesNo, "UR用の利用確認") = vbYes Then
        pathUR = parentPath & "\UR用見積書"
        On Error Resume Next: If Dir(pathUR, vbDirectory) = "" Then MkDir pathUR: On Error GoTo 0
        フォルダパス保存 FOLDER_PROP_UR, pathUR
        MsgBox "① " & pathNormal & vbCrLf & "② " & pathUR & vbCrLf & "を作成しました。", vbInformation
    Else
        フォルダパス保存 FOLDER_PROP_UR, ""
        MsgBox "・ " & pathNormal & vbCrLf & "を作成しました。", vbInformation
    End If
End Sub


'====================================================
' 起動時自動取込
'====================================================
  Sub 起動時自動取込()
    ' フォルダ未設定の場合はメッセージなしでスキップ
    Dim fPathNormal As String
    fPathNormal = フォルダパス取得(FOLDER_PROP_NORMAL)
    If fPathNormal = "" Then Exit Sub

    Dim tbl As ListObject          ' ← 追加
    Set tbl = テーブル取得()       ' ← 追加
    If tbl Is Nothing Then Exit Sub

    Dim col() As Long
    If Not テーブル列取得(tbl, col) Then Exit Sub

    Dim fPathUR As String          ' ← 追加
    fPathUR = フォルダパス取得(FOLDER_PROP_UR)


    Dim 成功数 As Long, スキップ数 As Long
    Dim 成功リスト As String, スキップリスト As String, 新着あり As Boolean

    Application.ScreenUpdating = False

    フォルダ内一括処理 fPathNormal, tbl, col, NORMAL_CELL_担当者, NORMAL_CELL_物件名, GetNormalBlocks(), _
        NORMAL_COL_材料, NORMAL_COL_数量, NORMAL_COL_単位, False, 成功数, スキップ数, 成功リスト, スキップリスト, 新着あり

    If fPathUR <> "" Then
        フォルダ内一括処理 fPathUR, tbl, col, UR_CELL_担当者, UR_CELL_物件名, GetURBlocks(), _
            UR_COL_材料, UR_COL_数量, UR_COL_単位, True, 成功数, スキップ数, 成功リスト, スキップリスト, 新着あり
    End If

    Application.ScreenUpdating = True
    If 新着あり Then 結果表示 "自動取込完了", 成功数, 成功リスト, スキップ数, スキップリスト
End Sub


'====================================================
' 手動取込
'====================================================
Sub 手動取込_通常()
    手動取込共通 "通常見積書", NORMAL_CELL_担当者, NORMAL_CELL_物件名, GetNormalBlocks(), _
        NORMAL_COL_材料, NORMAL_COL_数量, NORMAL_COL_単位, False
End Sub

Sub 手動取込_UR用()
    手動取込共通 "UR用見積書", UR_CELL_担当者, UR_CELL_物件名, GetURBlocks(), _
        UR_COL_材料, UR_COL_数量, UR_COL_単位, True
End Sub

' ★【修正】ScreenUpdating の制御を continueProcess フラグで明示化
'           Exit For 後も確実に True へ戻るよう整理
Private Sub 手動取込共通(titlePrefix As String, cTanto As String, cBukken As String, _
                          blocks As Variant, cZairyo As Long, cSuryo As Long, cTani As Long, _
                          ByVal isUR As Boolean)
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "取り込む【" & titlePrefix & "】を選択（複数可）"
    fd.Filters.Add "Excelファイル", "*.xlsx;*.xlsm;*.xls"
    fd.AllowMultiSelect = True
    If fd.Show = False Then Exit Sub

    Dim tbl As ListObject: Set tbl = テーブル取得()
    If tbl Is Nothing Then Exit Sub
    Dim col() As Long: If Not テーブル列取得(tbl, col) Then Exit Sub

    Dim 成功数 As Long, スキップ数 As Long, fileIdx As Long
    Dim filePath As String, resultMsg As String
    Dim 物件名Out As String, 取込件数Out As Long
    Dim 成功リスト As String, スキップリスト As String
    Dim continueProcess As Boolean: continueProcess = True  ' ★ フラグ管理

    Application.ScreenUpdating = False

    For fileIdx = 1 To fd.SelectedItems.count
        If Not continueProcess Then Exit For
        filePath = fd.SelectedItems(fileIdx)
        resultMsg = ファイル取込1件(filePath, tbl, col, cTanto, cBukken, blocks, _
                                    cZairyo, cSuryo, cTani, isUR, 物件名Out, 取込件数Out)
        Select Case resultMsg
            Case ""
                成功数 = 成功数 + 1
                成功リスト = 成功リスト & "  ○ " & 物件名Out & "（" & 取込件数Out & " 行）" & vbCrLf
            Case "SKIP"
                スキップ数 = スキップ数 + 1
                スキップリスト = スキップリスト & "  - " & GetFileName(filePath) & "（スキップ）" & vbCrLf
            Case Else
                スキップ数 = スキップ数 + 1
                スキップリスト = スキップリスト & "  x " & GetFileName(filePath) & "（エラー）" & vbCrLf
                Application.ScreenUpdating = True
                If MsgBox("エラー：" & resultMsg & vbCrLf & "残りを続けますか？", vbCritical + vbYesNo) = vbNo Then
                    continueProcess = False  ' ★ Exit For は使わずフラグで制御
                End If
                Application.ScreenUpdating = False
        End Select
    Next fileIdx

    Application.ScreenUpdating = True  ' ★ continueProcess = False 経由でも必ず到達
    結果表示 titlePrefix & " 取込完了", 成功数, 成功リスト, スキップ数, スキップリスト
End Sub


'====================================================
' フォルダ内一括処理（自動取込用）
'====================================================
Private Sub フォルダ内一括処理(fPath As String, tbl As ListObject, col() As Long, _
    cTanto As String, cBukken As String, blocks As Variant, _
    cZairyo As Long, cSuryo As Long, cTani As Long, ByVal isUR As Boolean, _
    ByRef 成功数 As Long, ByRef スキップ数 As Long, _
    ByRef 成功リスト As String, ByRef スキップリスト As String, ByRef 新着あり As Boolean)

    Dim fileName As String, filePath As String, resultMsg As String
    Dim 物件名Out As String, 取込件数Out As Long
    Dim ext As Variant

    For Each ext In Array("*.xlsx", "*.xlsm", "*.xls")
        fileName = Dir(fPath & "\" & CStr(ext))
        Do While fileName <> ""
            filePath = fPath & "\" & fileName
            If Not 取込済みチェック(filePath) Then
                新着あり = True
                resultMsg = ファイル取込1件(filePath, tbl, col, cTanto, cBukken, blocks, _
                                            cZairyo, cSuryo, cTani, isUR, 物件名Out, 取込件数Out)
                Select Case resultMsg
                    Case ""
                        成功数 = 成功数 + 1
                        成功リスト = 成功リスト & "  ○ " & 物件名Out & "（" & 取込件数Out & " 行）" & vbCrLf
                    Case "SKIP"
                        スキップ数 = スキップ数 + 1
                        スキップリスト = スキップリスト & "  - " & fileName & "（スキップ）" & vbCrLf
                    Case Else
                        スキップ数 = スキップ数 + 1
                        スキップリスト = スキップリスト & "  x " & fileName & "（エラー：" & resultMsg & "）" & vbCrLf
                End Select
            End If
            fileName = Dir()
        Loop
    Next ext
End Sub


'====================================================
' 1ファイル取込（オーケストレーター）
' ★【修正】シートコピー・データ走査・テーブル書込を分離
'====================================================
Private Function ファイル取込1件(filePath As String, tbl As ListObject, col() As Long, _
    cTanto As String, cBukken As String, blocks As Variant, _
    cZairyo As Long, cSuryo As Long, cTani As Long, ByVal isUR As Boolean, _
    ByRef 物件名Out As String, ByRef 取込件数Out As Long) As String

    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim 担当者 As String, 物件名 As String
    Dim dicQty  As Object: Set dicQty = CreateObject("Scripting.Dictionary")
    Dim dicUnit As Object: Set dicUnit = CreateObject("Scripting.Dictionary")
    Dim orderList As New Collection
    Dim unitWarnings As String

    On Error GoTo エラー

    ' --- 取込済みチェック ---
    If 取込済みチェック(filePath) Then
        If MsgBox("すでに取り込み済みです。" & vbCrLf & filePath & vbCrLf & _
                  "もう一度取り込みますか？", vbQuestion + vbYesNo) = vbNo Then
            ファイル取込1件 = "SKIP": Exit Function
        End If
    End If

    ' --- ソースファイルを開く ---
    Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True)
    Dim s As Worksheet
    For Each s In wbSrc.Worksheets
        If s.Visible = xlSheetVisible Then Set wsSrc = s: Exit For
    Next s
    If wsSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
        ファイル取込1件 = "表示シートが見つかりません。": Exit Function
    End If

    ' --- ヘッダ情報取得 ---
    担当者 = Trim(CStr(Nz(tbl.Parent.Range(cTanto).Value)))
    If 担当者 = "" Then 担当者 = "（担当者不明）"
    物件名 = Trim(CStr(Nz(wsSrc.Range(cBukken).Value)))
    If 物件名 = "" Then 物件名 = "（物件名なし）"

    ' --- シートコピー ---
    シートコピー実行 wsSrc, 物件名

    ' --- データ走査（指定ブロックのみ）---
    ブロックデータ走査 wsSrc, blocks, cZairyo, cSuryo, cTani, dicQty, dicUnit, orderList, unitWarnings

    wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing

    ' --- 単位不一致警告 ---
    If unitWarnings <> "" Then
        単位不一致ログ記録 物件名, unitWarnings
        MsgBox "【" & 物件名 & "】単位が異なる材料があります。" & vbCrLf & unitWarnings, vbExclamation
    End If

    If orderList.count = 0 Then
        ファイル取込1件 = "指定範囲内に取込データがありません。": Exit Function
    End If

    ' --- 再取込時：同一物件名の既存行を事前削除 ★【修正③】---
    既存データ削除 tbl, col(2), 物件名, col(6), isUR

    ' --- テーブル書込 ---
    取込件数Out = テーブルに書き込む(tbl, col, 担当者, 物件名, dicQty, dicUnit, orderList, isUR)

    ログ記録 filePath, 担当者, 物件名, 取込件数Out
    物件名Out = 物件名
    ファイル取込1件 = ""
    Exit Function

エラー:
    ' ★【修正④】On Error Resume Next を Close の直前に分離し、
    '            エラー番号をクリアする前に Description を退避
    Dim errMsg As String: errMsg = Err.Description
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    ファイル取込1件 = errMsg
End Function


'====================================================
' シートコピー実行 ★【分離】
' 既存シートがある場合はユーザーに確認して上書きまたはスキップ
'====================================================
Private Sub シートコピー実行(wsSrc As Worksheet, 物件名 As String)
    Dim sheetName As String
    Dim existSheet As Worksheet

    sheetName = シート名サニタイズ(物件名)
    On Error Resume Next: Set existSheet = ThisWorkbook.Sheets(sheetName): On Error GoTo 0

    If Not existSheet Is Nothing Then
        If MsgBox("「" & sheetName & "」シートが既にあります。削除して上書きしますか？", _
                  vbQuestion + vbYesNo) = vbYes Then
            Application.DisplayAlerts = False
            existSheet.Delete
            Application.DisplayAlerts = True
        Else
            Exit Sub  ' スキップ（コピーしない）
        End If
    End If

    wsSrc.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

    Dim copiedSheet As Worksheet
    Set copiedSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    On Error Resume Next
    copiedSheet.Name = sheetName
    If Err.Number <> 0 Then copiedSheet.Name = Left(sheetName, 24) & "_" & Format(Now, "mmdd")
    On Error GoTo 0
End Sub


'====================================================
' ブロックデータ走査 ★【分離】
' 指定ブロックを走査し、dicQty / dicUnit / orderList を構築
'====================================================
Private Sub ブロックデータ走査(wsSrc As Worksheet, blocks As Variant, _
    cZairyo As Long, cSuryo As Long, cTani As Long, _
    dicQty As Object, dicUnit As Object, _
    ByRef orderList As Collection, ByRef unitWarnings As String)

    Dim b As Variant, i As Long, startR As Long, endR As Long
    Dim 材料 As String, 単位 As String, 数量Raw As Variant, 数量 As Double

    For Each b In blocks
        startR = b(0): endR = b(1)
        For i = startR To endR
            材料 = Trim(CStr(Nz(wsSrc.Cells(i, cZairyo).Value)))
            数量Raw = wsSrc.Cells(i, cSuryo).Value
            単位 = Trim(CStr(Nz(wsSrc.Cells(i, cTani).Value)))

            If 材料 <> "" And 数値として有効(数量Raw) Then
                数量 = CDbl(数量Raw)
                If dicQty.Exists(材料) Then
                    If dicUnit(材料) <> 単位 And 単位 <> "" Then
                        unitWarnings = unitWarnings & "・" & 材料 & _
                                       "（" & dicUnit(材料) & " / " & 単位 & "）" & vbCrLf
                    End If
                    dicQty(材料) = dicQty(材料) + 数量
                Else
                    dicQty.Add 材料, 数量
                    dicUnit.Add 材料, 単位
                    orderList.Add 材料   ' ★ ReDim Preserve 廃止 → Collection で管理
                End If
            End If
        Next i
    Next b
End Sub


'====================================================
' テーブルに書き込む ★【分離】
' 戻り値：書き込んだ行数
'====================================================
Private Function テーブルに書き込む(tbl As ListObject, col() As Long, _
    担当者 As String, 物件名 As String, _
    dicQty As Object, dicUnit As Object, orderList As Collection, _
    ByVal isUR As Boolean) As Long

    Dim 空行キュー As Collection
    Set 空行キュー = 空行インデックス取得(tbl, col(2), col(3))

    Dim targetRow As ListRow, 取込件数 As Long
    Dim k As Variant

    For Each k In orderList
        If 空行キュー.count > 0 Then
            Set targetRow = tbl.ListRows(CLng(空行キュー(1)))
            空行キュー.Remove 1
        Else
            Set targetRow = tbl.ListRows.Add
        End If
        If col(1) > 0 Then targetRow.Range(col(1)).Value = 担当者
        targetRow.Range(col(2)).Value = 物件名
        targetRow.Range(col(3)).Value = CStr(k)
        targetRow.Range(col(4)).Value = dicQty(k)
        targetRow.Range(col(5)).Value = dicUnit(k)
        If col(6) > 0 Then
            If isUR Then
                targetRow.Range(col(6)).Value = "UR"
            Else
                targetRow.Range(col(6)).ClearContents
            End If
        End If
        取込件数 = 取込件数 + 1
    Next k

    テーブルに書き込む = 取込件数
End Function


'====================================================
' 再取込時：同一物件名の既存テーブル行を事前削除 ★【追加】
'====================================================
Private Sub 既存データ削除(tbl As ListObject, colBukken As Long, 物件名 As String, _
                             ByVal colUR As Long, ByVal isUR As Boolean)
    Dim i As Long
    Dim rowUR As String

    For i = tbl.ListRows.count To 1 Step -1
        If Trim(CStr(tbl.ListRows(i).Range(colBukken).Value)) = 物件名 Then

            If colUR > 0 Then
                rowUR = Trim(CStr(tbl.ListRows(i).Range(colUR).Value))

                ' 同一物件名でも、通常取込とUR取込は別物として扱う。
                ' 通常の再取込では通常行だけ削除。
                ' URの再取込ではUR行だけ削除。
                If isUR Then
                    If InStr(1, UCase(rowUR), "UR", vbTextCompare) > 0 Then
                        tbl.ListRows(i).Delete
                    End If
                Else
                    If InStr(1, UCase(rowUR), "UR", vbTextCompare) = 0 Then
                        tbl.ListRows(i).Delete
                    End If
                End If

            Else
                ' UR列がない場合は従来通り、同一物件名を削除
                tbl.ListRows(i).Delete
            End If

        End If
    Next i
End Sub


'====================================================
' 見積削除（frmDelete 呼び出し版）
'====================================================
Sub 見積削除()
    Dim tbl As ListObject
    Set tbl = テーブル取得()
    If tbl Is Nothing Then Exit Sub

    Dim col担当者 As Long, col現場 As Long
    col担当者 = 列番号取得(tbl, "担当者")
    col現場 = 列番号取得(tbl, "現場")

    If col担当者 = 0 Or col現場 = 0 Then
        MsgBox "テーブルに「担当者」または「現場」列が見つかりません。", vbExclamation
        Exit Sub
    End If

    With frmDelete
        .Init tbl, col担当者, col現場
        .Show
    End With
End Sub


'====================================================
' ログ表示
'====================================================
Sub 取込ログ表示()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(LOG_SHEET): On Error GoTo 0
    If ws Is Nothing Then MsgBox "取込ログがありません。", vbInformation Else ws.Activate
End Sub

Sub 単位不一致ログ表示()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(WARN_SHEET): On Error GoTo 0
    If ws Is Nothing Then MsgBox "単位不一致ログはありません。", vbInformation Else ws.Activate
End Sub


'====================================================
' 補助関数
'====================================================
Private Function フォルダパス取得(propName As String) As String
    Dim saved As String
    On Error Resume Next
    saved = ThisWorkbook.CustomDocumentProperties(propName).Value
    On Error GoTo 0
    If saved <> "" Then
        If Dir(saved, vbDirectory) <> "" Then フォルダパス取得 = saved: Exit Function
    End If
    フォルダパス取得 = ""
End Function

Private Sub フォルダパス保存(propName As String, folderPath As String)
    On Error Resume Next: ThisWorkbook.CustomDocumentProperties(propName).Delete: On Error GoTo 0
    If folderPath <> "" Then
        ThisWorkbook.CustomDocumentProperties.Add Name:=propName, LinkToContent:=False, _
            Type:=msoPropertyTypeString, Value:=folderPath
    End If
End Sub

Private Function テーブル取得() As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = TABLE_NAME Then Set テーブル取得 = lo: Exit Function
        Next lo
    Next ws
    MsgBox "「" & TABLE_NAME & "」が見つかりません。", vbExclamation
End Function

Private Function テーブル列取得(tbl As ListObject, ByRef col() As Long) As Boolean
    ReDim col(1 To 6)
    col(1) = 列番号取得(tbl, "担当者")
    col(2) = 列番号取得(tbl, "現場")
    col(3) = 列番号取得(tbl, "材料")
    col(4) = 列番号取得(tbl, "数量")
    col(5) = 列番号取得(tbl, "単位")
    col(6) = 列番号取得(tbl, "UR")

    If col(2) = 0 Or col(3) = 0 Or col(4) = 0 Or col(5) = 0 Then
        MsgBox "テーブルに必要な列（現場・材料・数量・単位）がありません。", vbExclamation
        テーブル列取得 = False: Exit Function
    End If

    If col(6) = 0 Then
        MsgBox "テーブルに「UR」列が見つかりません。" & vbCrLf & _
               "取込自体は可能ですが、UR印は付与されません。", vbExclamation
    End If

    テーブル列取得 = True
End Function

Private Function 列番号取得(tbl As ListObject, headerName As String) As Long
    Dim c As ListColumn
    For Each c In tbl.ListColumns
        If Trim(c.Name) = headerName Then 列番号取得 = c.Index: Exit Function
    Next c
    列番号取得 = 0
End Function

Private Sub 結果表示(title As String, cSuccess As Long, lSuccess As String, cSkip As Long, lSkip As String)
    Dim msg As String
    msg = "【" & title & "】" & vbCrLf & vbCrLf & "■ 成功：" & cSuccess & " ファイル" & vbCrLf
    If lSuccess <> "" Then msg = msg & lSuccess & vbCrLf
    If cSkip > 0 Then msg = msg & "■ スキップ／エラー：" & cSkip & " ファイル" & vbCrLf & lSkip
    MsgBox msg, vbInformation
End Sub

Private Function 空行インデックス取得(tbl As ListObject, colA As Long, colB As Long) As Collection
    Dim c As New Collection, i As Long
    For i = 1 To tbl.ListRows.count
        If Trim(CStr(tbl.ListRows(i).Range(colA).Value)) = "" And _
           Trim(CStr(tbl.ListRows(i).Range(colB).Value)) = "" Then c.Add i
    Next i
    Set 空行インデックス取得 = c
End Function

Private Function Nz(v As Variant) As Variant
    If IsNull(v) Or IsEmpty(v) Then Nz = "" Else Nz = v
End Function

Private Function 数値として有効(v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then If Trim(CStr(v)) = "" Then Exit Function
    If IsNumeric(v) Then 数値として有効 = True
End Function

' ★ frmDelete から呼ぶため Public
Public Function シート名サニタイズ(s As String) As String
    Dim tmp As String, c As Variant
    tmp = Trim(s)
    For Each c In Array("\", "/", "?", "*", "[", "]", ":")
        tmp = Replace(tmp, CStr(c), "_")
    Next c
    If Len(tmp) > SHEET_NAME_MAX Then tmp = Left(tmp, SHEET_NAME_MAX)
    If tmp = "" Then tmp = "無題"
    シート名サニタイズ = tmp
End Function

Private Function 取込済みチェック(filePath As String) As Boolean
    Dim ws As Worksheet, i As Long
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(LOG_SHEET): On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For i = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = filePath Then 取込済みチェック = True: Exit Function
    Next i
End Function

Private Sub ログ記録(filePath As String, 担当者 As String, 物件名 As String, 件数 As Long)
    Dim ws As Worksheet, nextRow As Long
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(LOG_SHEET): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = LOG_SHEET
        ws.Cells(1, 1).Resize(, 5).Value = Array("ファイルパス", "担当者", "物件名", "取込件数", "取込日時")
        ws.Rows(1).Font.Bold = True
    End If
    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Resize(, 5).Value = Array(filePath, 担当者, 物件名, 件数, Now())
    ws.Cells(nextRow, 5).NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Columns("A:E").AutoFit
End Sub

Private Sub 単位不一致ログ記録(物件名 As String, 内容 As String)
    Dim ws As Worksheet, nextRow As Long
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(WARN_SHEET): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = WARN_SHEET
        ws.Cells(1, 1).Resize(, 3).Value = Array("日時", "物件名", "不一致内容")
        ws.Rows(1).Font.Bold = True
    End If
    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 1).NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Cells(nextRow, 2).Value = 物件名
    ws.Cells(nextRow, 3).Value = 内容
    ws.Cells(nextRow, 3).WrapText = True
End Sub

' ★ frmDelete から呼ぶため Public
Public Sub ログから削除(担当者 As String, 物件名 As String)
    Dim ws As Worksheet, i As Long
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(LOG_SHEET): On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    For i = ws.Cells(ws.Rows.count, 1).End(xlUp).Row To 2 Step -1
        If Trim(CStr(ws.Cells(i, 2).Value)) = 担当者 Then
            If 物件名 = "" Or Trim(CStr(ws.Cells(i, 3).Value)) = 物件名 Then
                ws.Rows(i).Delete
            End If
        End If
    Next i
End Sub

Private Function GetFileName(filePath As String) As String
    GetFileName = Mid(filePath, InStrRev(filePath, "\") + 1)
End Function


