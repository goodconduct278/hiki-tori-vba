- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Sub TransferDataByProductNumber()

    Dim wsSource As Worksheet
    Dim wsFour As Worksheet
    Dim wsransferData As Worksheet

    Dim lastRowSource As Long
    Dim lastRowFour As Long
    Dim lastRowransferData As Long

    Dim targetCol As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim productNumber As String
    Dim targetValue1 As Variant
    Dim targetValue2 As Variant

    ' 現在のExcelファイルを上書き保存
    ThisWorkbook.Save

    ' 使用するシートを設定
    Set wsSource = ThisWorkbook.Sheets("担当者別 引取り予定表")
    Set wsFour = ThisWorkbook.Sheets("首都圏(四工品)")
    Set wsransferData = ThisWorkbook.Sheets("首都圏(仕入品)")

    ' ===== ① 転写前に既存データをクリア =====
    Dim pageRangesFour As Variant
    Dim pageRangesProc As Variant

    ' 四工品側のクリア範囲
    pageRangesFour = Array( _
        "N18:U71", _
        "N91:U144", _
        "N164:U217", _
        "N237:U290", _
        "N310:U363" _
    )

    ' 仕入品側のクリア範囲
    pageRangesProc = Array( _
        "N15:U24", _
        "N27:U68", _
        "N87:U142", _
        "N159:U214", _
        "N224:U287", _
        "N297:U357" _
    )

    ' 仕入品側 pageRangesProc を基準にクリアする
    ' 四工品側 pageRangesFour は範囲数が少ないため、存在する範囲だけクリアする
    For k = LBound(pageRangesProc) To UBound(pageRangesProc)

        If k <= UBound(pageRangesFour) Then
            wsFour.Range(pageRangesFour(k)).ClearContents
        End If

        wsransferData.Range(pageRangesProc(k)).ClearContents

    Next k

    ' ===== 元データ行数確認 =====
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastRowFour = wsFour.Cells(wsFour.Rows.Count, 2).End(xlUp).Row
    lastRowransferData = wsransferData.Cells(wsransferData.Rows.Count, 2).End(xlUp).Row

    ' 「担当者別 引取り予定表」の「合計」の列を検索
    targetCol = 0

    For i = 1 To wsSource.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column

        If wsSource.Cells(2, i).Value = "合計" Then
            targetCol = i
            Exit For
        End If

    Next i

    If targetCol = 0 Then
        MsgBox "「合計」と一致する列が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    ' 転写対象列
    Dim sourceCol1 As Long
    Dim sourceCol2 As Long

    sourceCol1 = targetCol
    sourceCol2 = targetCol + 1

    ' 転写先列
    Dim destCol1 As Long
    Dim destCol2 As Long

    destCol1 = 14 ' N列
    destCol2 = 18 ' R列

    ' ===== ② データ転写 =====
    For i = 3 To lastRowSource

        productNumber = Trim(wsSource.Cells(i, 1).Value) ' 製品番号 A列

        ' 製品番号が空の場合はスキップ
        If productNumber = "" Then
            GoTo SkipRow
        End If

        targetValue1 = wsSource.Cells(i, sourceCol1).Value
        targetValue2 = wsSource.Cells(i, sourceCol2).Value

        ' 四工品
        For j = 2 To lastRowFour

            If wsFour.Cells(j, 2).Value = productNumber Then
                wsFour.Cells(j, destCol1).Value = targetValue1
                wsFour.Cells(j, destCol2).Value = targetValue2
                Exit For
            End If

        Next j

        ' 仕入品
        For j = 2 To lastRowransferData

            If wsransferData.Cells(j, 2).Value = productNumber Then
                wsransferData.Cells(j, destCol1).Value = targetValue1
                wsransferData.Cells(j, destCol2).Value = targetValue2
                Exit For
            End If

        Next j

SkipRow:
    Next i

    ' 年月を転写
    wsFour.Range("B2").Value = wsSource.Range("M1").Value

    ' ===== ③ N1セル空白時の Left(..., Len(...) - 1) エラー防止 =====
    If Len(wsSource.Range("N1").Value) > 0 Then
        wsFour.Range("G2").Value = Left(wsSource.Range("N1").Value, Len(wsSource.Range("N1").Value) - 1)
    Else
        wsFour.Range("G2").Value = ""
    End If

    wsransferData.Range("C2").Value = wsFour.Range("G2").Value
    wsFour.Range("B4").Value = Format(Date, "yyyy/m/d")
    wsransferData.Range("B4").Value = Format(Date, "yyyy/m/d")

    MsgBox "データの転写が完了しました。", vbInformation

End Sub

Sub SaveSheetsAsFiles()

    On Error GoTo ErrorHandler

    Dim wsFour As Worksheet
    Dim wsProcurement As Worksheet
    Dim wsSettings As Worksheet
    Dim saveFolder As String
    Dim fileNameFour As String
    Dim fileNameProcurement As String
    Dim wbNew As Workbook
    Dim currentFolderPath As String

    Application.ScreenUpdating = False

    ' 使用するシートを設定
    Set wsFour = ThisWorkbook.Sheets("首都圏(四工品)")
    Set wsProcurement = ThisWorkbook.Sheets("首都圏(仕入品)")
    Set wsSettings = ThisWorkbook.Sheets("プログラム設定")

    ' 自身が保存されているフォルダのパスを取得して「プログラム設定」のセルA1に出力
    currentFolderPath = ThisWorkbook.Path
    wsSettings.Range("A1").Value = currentFolderPath

    ' 保存先のフォルダパスを取得
    saveFolder = wsSettings.Range("A2").Value

    ' 保存先フォルダが存在しない場合は作成
    If Dir(saveFolder, vbDirectory) = "" Then
        MkDir saveFolder
    End If

    ' ファイル名を取得
    fileNameFour = wsSettings.Range("A4").Value
    fileNameProcurement = wsSettings.Range("A5").Value

    ' 「首都圏(四工品)」シートを新しいファイルとして保存
    wsFour.Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs Filename:=saveFolder & "\" & fileNameFour, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False

    ' 「首都圏(仕入品)」シートを新しいファイルとして保存
    wsProcurement.Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs Filename:=saveFolder & "\" & fileNameProcurement, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False

    Application.ScreenUpdating = True

    MsgBox "ファイルの保存が完了しました。", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical

End Sub

Sub ClearData()

End Sub