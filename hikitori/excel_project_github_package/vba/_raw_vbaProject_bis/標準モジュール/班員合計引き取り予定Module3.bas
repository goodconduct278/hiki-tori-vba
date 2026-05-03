Attribute VB_Name = "Module3"

'====================================================
' 合計ファイル クエリ自動設定マクロ v4
' Workbook_Open から自動実行
' 名前定義「班員データフォルダ」を参照する動的版
' UR列対応版
'====================================================

Option Explicit

Private Const DATA_SUBFOLDER As String = "★班員データ"
Private Const Q_MAIN         As String = "★班員データ"
Private Const NM_DATA_FOLDER As String = "班員データフォルダ"
Private Const SETTING_SHEET  As String = "_自動設定"

'====================================================
' クエリ自動設定（ThisWorkbook.Workbook_Open / ボタン兼用）
'====================================================
Public Sub クエリ自動設定()

    Dim dataFolder As String
    dataFolder = 班員データフォルダ取得()

    If Dir(dataFolder, vbDirectory) = "" Then
        Exit Sub
    End If

    班員データフォルダ名前定義更新 dataFolder

    Dim hasMain As Boolean
    Dim q As WorkbookQuery

    For Each q In ThisWorkbook.Queries
        If q.Name = Q_MAIN Then
            hasMain = True
            Exit For
        End If
    Next q

    If hasMain Then
        クエリパス更新
        ThisWorkbook.RefreshAll
        Exit Sub
    End If

    クエリ新規作成
    ThisWorkbook.RefreshAll

End Sub

'====================================================
' 班員データフォルダ取得
' 名前定義があれば優先。なければ従来どおり同階層の★班員データを使う。
'====================================================
Private Function 班員データフォルダ取得() As String

    Dim dataFolder As String

    On Error Resume Next
    dataFolder = CStr(ThisWorkbook.Names(NM_DATA_FOLDER).RefersToRange.Value)
    On Error GoTo 0

    If Len(Trim(dataFolder)) = 0 Then
        dataFolder = ThisWorkbook.Path & "\" & DATA_SUBFOLDER
    End If

    If Right(dataFolder, 1) = "\" Then
        dataFolder = Left(dataFolder, Len(dataFolder) - 1)
    End If

    班員データフォルダ取得 = dataFolder

End Function

'====================================================
' 名前定義「班員データフォルダ」を作成・更新
' Power Query の Excel.CurrentWorkbook() から読めるよう、隠しシートのA1を名前定義にする。
'====================================================
Private Sub 班員データフォルダ名前定義更新(ByVal dataFolder As String)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTING_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SETTING_SHEET
    End If

    ws.Range("A1").Value = dataFolder
    ws.Visible = xlSheetVeryHidden

    On Error Resume Next
    ThisWorkbook.Names(NM_DATA_FOLDER).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=NM_DATA_FOLDER, RefersTo:="='" & SETTING_SHEET & "'!$A$1"

End Sub

'====================================================
' クエリパス更新
' 既存クエリを、名前定義「班員データフォルダ」参照版へ統一する。
'====================================================
Private Sub クエリパス更新()

    Dim q As WorkbookQuery

    For Each q In ThisWorkbook.Queries
        If q.Name = Q_MAIN Then
            q.Formula = クエリ式作成()
            Exit For
        End If
    Next q

End Sub

'====================================================
' クエリ新規作成
'====================================================
Private Sub クエリ新規作成()

    Dim wb As Workbook
    Set wb = ThisWorkbook

    On Error Resume Next
    wb.Queries(Q_MAIN).Delete
    On Error GoTo 0

    wb.Queries.Add Name:=Q_MAIN, Formula:=クエリ式作成()

End Sub

'====================================================
' Power Query M式作成
'====================================================
Private Function クエリ式作成() As String

    Dim mMain As String

    mMain = "let" & vbCrLf & _
            "    フォルダパス表 = Excel.CurrentWorkbook(){[Name=""" & NM_DATA_FOLDER & """]}[Content]," & vbCrLf & _
            "    フォルダパス = Text.From(フォルダパス表{0}[Column1])," & vbCrLf & _
            "    ソース = Folder.Files(フォルダパス)," & vbCrLf & _
            "    非表示ファイルのフィルタ = Table.SelectRows(ソース, each [Attributes]?[Hidden]? <> true)," & vbCrLf & _
            "    必要な列の選択 = Table.SelectColumns(非表示ファイルのフィルタ, {""Name"", ""Content""})," & vbCrLf & _
            "    ブックの展開 = Table.AddColumn(必要な列の選択, ""Data"", each Excel.Workbook([Content], null, true))," & vbCrLf & _
            "    テーブルの展開 = Table.ExpandTableColumn(ブックの展開, ""Data"", {""Name"", ""Data"", ""Item"", ""Kind""}, {""Sheet.Name"", ""Sheet.Data"", ""Item"", ""Kind""})," & vbCrLf & _
            "    集計テーブルのフィルタ = Table.SelectRows(テーブルの展開, each ([Item] = ""集計テーブル"" and [Kind] = ""Table""))," & vbCrLf & _
            "    不要な列の削除 = Table.SelectColumns(集計テーブルのフィルタ, {""Name"", ""Sheet.Data""})," & vbCrLf & _
            "    UR列補完 = Table.TransformColumns(不要な列の削除, {{""Sheet.Data"", each if Table.HasColumns(_, ""UR"") then _ else Table.AddColumn(_, ""UR"", each null)}})," & vbCrLf & _
            "    データの展開 = Table.ExpandTableColumn(UR列補完, ""Sheet.Data"", " & _
            "{""担当者"", ""得意先"", ""現場"", ""材料"", ""数量"", ""単位"", ""納品日"", ""注文状況"", ""注文日"", ""現場状況"", ""チェック"", ""日時"", ""UR""})," & vbCrLf & _
            "    型の変更 = Table.TransformColumnTypes(データの展開, {" & _
            "{""Name"", type text}, {""担当者"", type any}, {""得意先"", type any}, {""現場"", type any}, " & _
            "{""材料"", type any}, {""数量"", type any}, {""単位"", type any}, {""納品日"", type any}, " & _
            "{""注文状況"", type any}, {""注文日"", type any}, {""現場状況"", type any}, {""チェック"", type any}, " & _
            "{""日時"", type any}, {""UR"", type any}})" & vbCrLf & _
            "in" & vbCrLf & _
            "    型の変更"

    クエリ式作成 = mMain

End Function

