- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

'====================================================
' ファイル作成マクロ v10 修正版・丸ごと上書き版
'
' 修正内容：
' ・月入力が1～12以外なら停止
' ・名前定義「班員データフォルダ」を Power Query が読める形で作成
' ・プログラム設定シートが無い場合は自動作成
' ・UR列対応
' ・Power Queryは固定パスではなく、名前定義「班員データフォルダ」を参照
' ・作成ボタン1回で、個人ファイル作成、合計ファイル作成、クエリ作成、読込先設定まで実行
'====================================================

Private Const DB_SHEET      As String = "担当者DB"
Private Const CREATE_SHEET  As String = "ファイル作成シート"
Private Const TEMPLATE_個人 As String = "班員個人引き取り予定.xlsm"
Private Const TEMPLATE_合計 As String = "班員合計引き取り予定.xlsm"

Private Const FOLDER_PROP_NORMAL As String = "見積書フォルダパス_通常"
Private Const FOLDER_PROP_UR     As String = "見積書フォルダパス_UR"

' クエリ名・シート名
Private Const Q_MAIN      As String = "★班員データ"
Private Const Q_SHEET     As String = "★班員データ"

' 作成した合計ファイル内に保存する名前定義
Private Const NAME_MEMBER_DATA_FOLDER As String = "班員データフォルダ"
Private Const SETTING_SHEET As String = "プログラム設定"


'====================================================
' 作成ボタン
'====================================================
Sub ファイル作成()

    Dim wsCreate As Worksheet, wsDB As Worksheet
    Set wsCreate = ThisWorkbook.Sheets(CREATE_SHEET)
    Set wsDB = ThisWorkbook.Sheets(DB_SHEET)

    ' ① 入力チェック
    Dim エリア As String, 月 As Long, 月文字 As String
    エリア = Trim(CStr(wsCreate.Range("A3").Value))

    If エリア = "" Then
        MsgBox "エリアを選択してください。", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(wsCreate.Range("B3").Value) Or _
       Trim(CStr(wsCreate.Range("B3").Value)) = "" Then
        MsgBox "月を数字で入力してください。", vbExclamation
        Exit Sub
    End If

    月 = CLng(wsCreate.Range("B3").Value)

    If 月 < 1 Or 月 > 12 Then
        MsgBox "月は1～12の数字で入力してください。", vbExclamation
        Exit Sub
    End If

    月文字 = CStr(月)

    ' ② 元ファイルフォルダ選択
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "元ファイルフォルダを選択してください（★元ファイル）"

    If fd.Show = False Then Exit Sub

    Dim 元フォルダ As String
    元フォルダ = fd.SelectedItems(1)

    If Dir(元フォルダ & "\" & TEMPLATE_個人) = "" Then
        MsgBox "「" & TEMPLATE_個人 & "」が見つかりません。" & vbCrLf & 元フォルダ, vbCritical
        Exit Sub
    End If

    If Dir(元フォルダ & "\" & TEMPLATE_合計) = "" Then
        MsgBox "「" & TEMPLATE_合計 & "」が見つかりません。" & vbCrLf & 元フォルダ, vbCritical
        Exit Sub
    End If

    ' ③ 保存先フォルダ選択
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "保存先フォルダを選択してください"

    If fd.Show = False Then Exit Sub

    Dim 保存先親 As String
    保存先親 = fd.SelectedItems(1)

    ' ④ 担当者DBからエリアで絞り込み
    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, "B").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "担当者DBにデータがありません。", vbCritical
        Exit Sub
    End If

    Dim 担当者リスト As New Collection
    Dim i As Long, personName As String

    For i = 2 To lastRow

        If Trim(CStr(wsDB.Cells(i, 1).Value)) = エリア Then
            personName = Trim(CStr(wsDB.Cells(i, 2).Value))

            If personName <> "" Then
                担当者リスト.Add personName
            End If
        End If

    Next i

    If 担当者リスト.Count = 0 Then
        MsgBox "「" & エリア & "」に該当する担当者が見つかりません。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrHandler

    ' ⑤ フォルダ構造作成
    Dim 月フォルダ As String
    月フォルダ = 保存先親 & "\" & 月文字 & "月引き取り"
    フォルダ作成 月フォルダ

    Dim 班員データフォルダ As String
    班員データフォルダ = 月フォルダ & "\★班員データ"
    フォルダ作成 班員データフォルダ

    Dim person As Variant

    For Each person In 担当者リスト

        Dim 個人フォルダ As String
        個人フォルダ = 月フォルダ & "\" & CStr(person) & 月文字 & "月"

        フォルダ作成 個人フォルダ
        フォルダ作成 個人フォルダ & "\通常見積書"
        フォルダ作成 個人フォルダ & "\UR用見積書"

    Next person

    ' ⑥ 個人ファイル作成
    Dim 成功リスト As String
    Dim wbSrc As Workbook, wbNew As Workbook
    Dim savePath As String

    For Each person In 担当者リスト

        savePath = 班員データフォルダ & "\" & 月文字 & "月引き取り(" & CStr(person) & ").xlsm"

        If Dir(savePath) <> "" Then

            Application.ScreenUpdating = True
            Application.DisplayAlerts = True

            If MsgBox("「" & 月文字 & "月引き取り(" & CStr(person) & ").xlsm」は既に存在します。" & vbCrLf & _
                      "上書きしますか？", vbQuestion + vbYesNo) = vbNo Then

                Application.ScreenUpdating = False
                Application.DisplayAlerts = False
                GoTo NextPerson

            End If

            Application.ScreenUpdating = False
            Application.DisplayAlerts = False

        End If

        Set wbSrc = Workbooks.Open(元フォルダ & "\" & TEMPLATE_個人, ReadOnly:=True)
        wbSrc.SaveCopyAs savePath
        wbSrc.Close SaveChanges:=False
        Set wbSrc = Nothing

        Set wbNew = Workbooks.Open(savePath)

        Dim ws個人 As Worksheet, s As Worksheet
        Set ws個人 = Nothing

        For Each s In wbNew.Worksheets

            If s.Visible = xlSheetVisible Then
                Set ws個人 = s
                Exit For
            End If

        Next s

        If Not ws個人 Is Nothing Then

            ws個人.Range("B1").Value = CStr(person)

            On Error Resume Next
            ws個人.Name = CStr(person)

            If Err.Number <> 0 Then
                Err.Clear
                ws個人.Name = Left(CStr(person), 28)
            End If

            On Error GoTo ErrHandler

        End If

        ' 個人ファイル内に通常見積書/UR用見積書フォルダパスを保存
        Dim 通常Path As String, URPath As String
        通常Path = 月フォルダ & "\" & CStr(person) & 月文字 & "月\通常見積書"
        URPath = 月フォルダ & "\" & CStr(person) & 月文字 & "月\UR用見積書"

        On Error Resume Next
        wbNew.CustomDocumentProperties(FOLDER_PROP_NORMAL).Delete
        wbNew.CustomDocumentProperties(FOLDER_PROP_UR).Delete
        On Error GoTo ErrHandler

        wbNew.CustomDocumentProperties.Add Name:=FOLDER_PROP_NORMAL, _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=通常Path

        wbNew.CustomDocumentProperties.Add Name:=FOLDER_PROP_UR, _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=URPath

        wbNew.Save
        wbNew.Close SaveChanges:=False
        Set wbNew = Nothing

        成功リスト = 成功リスト & "  ○ " & CStr(person) & vbCrLf

NextPerson:
    Next person

    ' ⑦ 合計ファイル作成
    Dim 合計Path As String
    合計Path = 月フォルダ & "\" & 月文字 & "月合計引き取り予定.xlsm"

    If Dir(合計Path) <> "" Then

        Application.ScreenUpdating = True
        Application.DisplayAlerts = True

        If MsgBox("「" & 月文字 & "月合計引き取り予定.xlsm」は既に存在します。上書きしますか？", _
                  vbQuestion + vbYesNo) = vbNo Then
            GoTo 完了
        End If

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

    End If

    Set wbSrc = Workbooks.Open(元フォルダ & "\" & TEMPLATE_合計, ReadOnly:=True)
    wbSrc.SaveCopyAs 合計Path
    wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing

    Set wbNew = Workbooks.Open(合計Path)

    ' 月数字を各シートに記入
    On Error Resume Next

    If WbSheetExists(wbNew, "担当者別 引取り予定表") Then
        wbNew.Sheets("担当者別 引取り予定表").Range("N1").Value = 月
    End If

    If WbSheetExists(wbNew, "首都圏(仕入品)") Then
        wbNew.Sheets("首都圏(仕入品)").Range("C2").Value = 月
    End If

    If WbSheetExists(wbNew, "首都圏(四工品)") Then
        wbNew.Sheets("首都圏(四工品)").Range("C2").Value = 月
    End If

    On Error GoTo ErrHandler

    ' ★作成した合計ファイル側に、参照する班員データフォルダを保存
    班員データフォルダパス保存 wbNew, 班員データフォルダ

    ' ★クエリを新規作成
    クエリ新規作成 wbNew

    ' ★クエリ出力先を★班員データシートに設定して読み込み
    クエリ出力先設定 wbNew

    wbNew.Save
    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing

完了:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "作成完了！" & vbCrLf & vbCrLf & _
           "■ 作成フォルダ" & vbCrLf & _
           "  " & 月フォルダ & vbCrLf & vbCrLf & _
           "■ 個人ファイル（" & 担当者リスト.Count & "名）" & vbCrLf & _
           成功リスト & vbCrLf & _
           "■ 合計ファイル" & vbCrLf & _
           "  ○ " & 月文字 & "月合計引き取り予定.xlsm", vbInformation, "完了"

    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    On Error Resume Next

    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False

    On Error GoTo 0

    MsgBox "ファイル作成中にエラーが発生しました。" & vbCrLf & _
           "エラー番号：" & Err.Number & vbCrLf & _
           "内容：" & Err.Description, vbCritical

End Sub


'====================================================
' クエリ出力先設定
'====================================================
Private Sub クエリ出力先設定(ByVal wb As Workbook)

    Dim wsData As Worksheet

    On Error Resume Next
    Set wsData = wb.Sheets(Q_SHEET)
    On Error GoTo 0

    If wsData Is Nothing Then
        Set wsData = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsData.Name = Q_SHEET
    End If

    ' 既存テーブル削除
    Dim lo As ListObject

    For Each lo In wsData.ListObjects
        lo.Delete
    Next lo

    wsData.Cells.Clear

    ' クエリの存在確認
    Dim q As WorkbookQuery

    On Error Resume Next
    Set q = wb.Queries(Q_MAIN)
    On Error GoTo 0

    If q Is Nothing Then
        MsgBox "クエリ「" & Q_MAIN & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    On Error GoTo クエリエラー

    Dim newLo As ListObject

    Set newLo = wsData.ListObjects.Add(SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & Q_MAIN & ";Extended Properties=""""", _
        Destination:=wsData.Range("A1"))

    With newLo.QueryTable
        .CommandType = xlCmdDefault
        .CommandText = Array(Q_MAIN)
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = True
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With

    newLo.Name = "TBL_班員データ"

    Exit Sub

クエリエラー:
    MsgBox "クエリ出力先の設定中にエラーが発生しました。" & vbCrLf & _
           "手動でデータ → クエリと接続 → ★班員データ → 右クリックから「読み込み先」でテーブルを設定してください。" & vbCrLf & _
           "エラー内容：" & Err.Description, vbExclamation

End Sub


'====================================================
' クエリ新規作成
'====================================================
Private Sub クエリ新規作成(ByVal wb As Workbook)

    On Error Resume Next
    wb.Queries(Q_MAIN).Delete
    On Error GoTo 0

    Dim part1 As String, part2 As String, part3 As String
    Dim part4 As String, part5 As String

    part1 = _
        "let" & vbCrLf & _
        "    フォルダ設定 = Excel.CurrentWorkbook(){[Name=""" & NAME_MEMBER_DATA_FOLDER & """]}[Content]," & vbCrLf & _
        "    FolderPath = Text.From(フォルダ設定{0}[Column1])," & vbCrLf & _
        "    ソース = Folder.Files(FolderPath)," & vbCrLf & _
        "    非表示ファイルのフィルタ = Table.SelectRows(ソース, each [Attributes]?[Hidden]? <> true)," & vbCrLf & _
        "    必要な列の選択 = Table.SelectColumns(非表示ファイルのフィルタ, {""Name"", ""Content""})," & vbCrLf & _
        "    ブックの展開 = Table.AddColumn(必要な列の選択, ""Data"", each Excel.Workbook([Content], null, true))," & vbCrLf & _
        "    テーブルの展開 = Table.ExpandTableColumn(ブックの展開, ""Data"", {""Name"", ""Data"", ""Item"", ""Kind""}, {""Sheet.Name"", ""Sheet.Data"", ""Item"", ""Kind""}),"

    part2 = vbCrLf & _
        "    集計テーブルのフィルタ = Table.SelectRows(テーブルの展開, each ([Item] = ""集計テーブル"" and [Kind] = ""Table""))," & vbCrLf & _
        "    標準化Data追加 = Table.AddColumn(" & vbCrLf & _
        "        集計テーブルのフィルタ," & vbCrLf & _
        "        ""標準化Data""," & vbCrLf & _
        "        each" & vbCrLf & _
        "            Table.SelectColumns(" & vbCrLf & _
        "                [Sheet.Data]," & vbCrLf & _
        "                {""担当者"", ""得意先"", ""現場"", ""材料"", ""数量"", ""単位"", ""納品日"", ""注文状況"", ""注文日"", ""現場状況"", ""チェック"", ""日時"", ""UR""}," & vbCrLf & _
        "                MissingField.UseNull" & vbCrLf & _
        "            )" & vbCrLf & _
        "    ),"

    part3 = vbCrLf & _
        "    不要な列の削除 = Table.SelectColumns(標準化Data追加, {""Name"", ""標準化Data""})," & vbCrLf & _
        "    データの展開 = Table.ExpandTableColumn(" & vbCrLf & _
        "        不要な列の削除," & vbCrLf & _
        "        ""標準化Data""," & vbCrLf & _
        "        {""担当者"", ""得意先"", ""現場"", ""材料"", ""数量"", ""単位"", ""納品日"", ""注文状況"", ""注文日"", ""現場状況"", ""チェック"", ""日時"", ""UR""}," & vbCrLf & _
        "        {""担当者"", ""得意先"", ""現場"", ""材料"", ""数量"", ""単位"", ""納品日"", ""注文状況"", ""注文日"", ""現場状況"", ""チェック"", ""日時"", ""UR""}" & vbCrLf & _
        "    ),"

    part4 = vbCrLf & _
        "    UR整備 = Table.TransformColumns(" & vbCrLf & _
        "        データの展開," & vbCrLf & _
        "        {" & vbCrLf & _
        "            {" & vbCrLf & _
        "                ""UR""," & vbCrLf & _
        "                each if _ = null then null else if Text.Contains(Text.Upper(Text.From(_)), ""UR"") then ""UR"" else null," & vbCrLf & _
        "                type nullable text" & vbCrLf & _
        "            }" & vbCrLf & _
        "        }" & vbCrLf & _
        "    )," & vbCrLf & _
        "    UR列名変更 = Table.RenameColumns(UR整備, {{""UR"", ""取込UR""}}, MissingField.Ignore)," & vbCrLf & _
        "    区分追加 = Table.AddColumn(UR列名変更, ""区分"", each if [取込UR] = ""UR"" then ""UR"" else ""通常"", type text),"

    part5 = vbCrLf & _
        "    型の変更 = Table.TransformColumnTypes(" & vbCrLf & _
        "        区分追加," & vbCrLf & _
        "        {" & vbCrLf & _
        "            {""Name"", type text}," & vbCrLf & _
        "            {""担当者"", type any}," & vbCrLf & _
        "            {""得意先"", type any}," & vbCrLf & _
        "            {""現場"", type any}," & vbCrLf & _
        "            {""材料"", type any}," & vbCrLf & _
        "            {""数量"", type any}," & vbCrLf & _
        "            {""単位"", type any}," & vbCrLf & _
        "            {""納品日"", type any}," & vbCrLf & _
        "            {""注文状況"", type any}," & vbCrLf & _
        "            {""注文日"", type any}," & vbCrLf & _
        "            {""現場状況"", type any}," & vbCrLf & _
        "            {""チェック"", type any}," & vbCrLf & _
        "            {""日時"", type any}," & vbCrLf & _
        "            {""取込UR"", type text}," & vbCrLf & _
        "            {""区分"", type text}" & vbCrLf & _
        "        }" & vbCrLf & _
        "    )" & vbCrLf & _
        "in" & vbCrLf & _
        "    型の変更"

    wb.Queries.Add Name:=Q_MAIN, Formula:=part1 & part2 & part3 & part4 & part5

End Sub


'====================================================
' 作成した合計ファイル側に班員データフォルダパスを保存
'====================================================
Private Sub 班員データフォルダパス保存(ByVal wb As Workbook, ByVal folderPath As String)

    Dim ws As Worksheet

    If Trim(folderPath) = "" Then
        MsgBox "保存する班員データフォルダパスが空です。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set ws = wb.Worksheets(SETTING_SHEET)
    On Error GoTo 0

    ' プログラム設定シートが無い場合は作成
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SETTING_SHEET
    End If

    ' C1 に ★班員データ フォルダのパスを保存
    ws.Range("C1").Value = folderPath

    ' 見出しも入れておく
    ws.Range("B1").Value = NAME_MEMBER_DATA_FOLDER

    ' 既存の名前定義を削除して作り直す
    On Error Resume Next
    wb.Names(NAME_MEMBER_DATA_FOLDER).Delete
    On Error GoTo 0

    ' Power Query の Excel.CurrentWorkbook() で読める形の名前定義にする
    wb.Names.Add _
        Name:=NAME_MEMBER_DATA_FOLDER, _
        RefersTo:="='" & ws.Name & "'!$C$1"

End Sub


'====================================================
' 補助：フォルダ作成
'====================================================
Private Sub フォルダ作成(ByVal folderPath As String)

    On Error Resume Next

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    On Error GoTo 0

End Sub


'====================================================
' 補助：指定ブック内のシート存在確認
'====================================================
Private Function WbSheetExists(ByVal wb As Workbook, ByVal sName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Sheets(sName)
    WbSheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function



+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|Suspicious|Open                |May open a file                              |
|Suspicious|MkDir               |May create a directory                       |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Hex String|'\x00\x02\x08\x19'  |00020819                                     |
|Hex String|'\x00\x00\x00\x00\x0|000000000046                                 |
|          |0F'                 |                                             |
|Hex String|'\x00\x02\x08 '     |00020820                                     |
|Base64    |'5'               |Name                                         |
|String    |                    |                                             |
|Base64    |'\rZ'              |Data                                         |
|String    |                    |                                             |
|Base64    |'"צ'                |Item                                         |
|String    |                    |                                             |
|Base64    |'*)'               |Kind                                         |
|String    |                    |                                             |
+----------+--------------------+---------------------------------------------+
MACRO SOURCE CODE WITH DEOBFUSCATED VBA STRINGS (EXPERIMENTAL):



Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "Sheet2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "Sheet3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "Sheet4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "Sheet5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "Module1"
Option Explicit

'====================================================
' ファイル作成マクロ v10 修正版・丸ごと上書き版
'
' 修正内容：
' ・月入力が1～12以外なら停止
' ・名前定義「班員データフォルダ」を Power Query が読める形で作成
' ・プログラム設定シートが無い場合は自動作成
' ・UR列対応
' ・Power Queryは固定パスではなく、名前定義「班員データフォルダ」を参照
' ・作成ボタン1回で、個人ファイル作成、合計ファイル作成、クエリ作成、読込先設定まで実行
'====================================================

Private Const DB_SHEET      As String = "担当者DB"
Private Const CREATE_SHEET  As String = "ファイル作成シート"
Private Const TEMPLATE_個人 As String = "班員個人引き取り予定.xlsm"
Private Const TEMPLATE_合計 As String = "班員合計引き取り予定.xlsm"

Private Const FOLDER_PROP_NORMAL As String = "見積書フォルダパス_通常"
Private Const FOLDER_PROP_UR     As String = "見積書フォルダパス_UR"

' クエリ名・シート名
Private Const Q_MAIN      As String = "★班員データ"
Private Const Q_SHEET     As String = "★班員データ"

' 作成した合計ファイル内に保存する名前定義
Private Const NAME_MEMBER_DATA_FOLDER As String = "班員データフォルダ"
Private Const SETTING_SHEET As String = "プログラム設定"


'====================================================
' 作成ボタン
'====================================================
Sub ファイル作成()

    Dim wsCreate As Worksheet, wsDB As Worksheet
    Set wsCreate = ThisWorkbook.Sheets(CREATE_SHEET)
    Set wsDB = ThisWorkbook.Sheets(DB_SHEET)

    ' ① 入力チェック
    Dim エリア As String, 月 As Long, 月文字 As String
    エリア = Trim(CStr(wsCreate.Range("A3").Value))

    If エリア = "" Then
        MsgBox "エリアを選択してください。", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(wsCreate.Range("B3").Value) Or        Trim(CStr(wsCreate.Range("B3").Value)) = "" Then
        MsgBox "月を数字で入力してください。", vbExclamation
        Exit Sub
    End If

    月 = CLng(wsCreate.Range("B3").Value)

    If 月 < 1 Or 月 > 12 Then
        MsgBox "月は1～12の数字で入力してください。", vbExclamation
        Exit Sub
    End If

    月文字 = CStr(月)

    ' ② 元ファイルフォルダ選択
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "元ファイルフォルダを選択してください（★元ファイル）"

    If fd.Show = False Then Exit Sub

    Dim 元フォルダ As String
    元フォルダ = fd.SelectedItems(1)

    If Dir(元フォルダ & "\" & TEMPLATE_個人) = "" Then
        MsgBox "「" & TEMPLATE_個人 & "」が見つかりません。" & vbCrLf & 元フォルダ, vbCritical
        Exit Sub
    End If

    If Dir(元フォルダ & "\" & TEMPLATE_合計) = "" Then
        MsgBox "「" & TEMPLATE_合計 & "」が見つかりません。" & vbCrLf & 元フォルダ, vbCritical
        Exit Sub
    End If

    ' ③ 保存先フォルダ選択
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "保存先フォルダを選択してください"

    If fd.Show = False Then Exit Sub

    Dim 保存先親 As String
    保存先親 = fd.SelectedItems(1)

    ' ④ 担当者DBからエリアで絞り込み
    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, "B").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "担当者DBにデータがありません。", vbCritical
        Exit Sub
    End If

    Dim 担当者リスト As New Collection
    Dim i As Long, personName As String

    For i = 2 To lastRow

        If Trim(CStr(wsDB.Cells(i, 1).Value)) = エリア Then
            personName = Trim(CStr(wsDB.Cells(i, 2).Value))

            If personName <> "" Then
                担当者リスト.Add personName
            End If
        End If

    Next i

    If 担当者リスト.Count = 0 Then
        MsgBox "「" & エリア & "」に該当する担当者が見つかりません。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrHandler

    ' ⑤ フォルダ構造作成
    Dim 月フォルダ As String
    月フォルダ = 保存先親 & "\" & 月文字 & "月引き取り"
    フォルダ作成 月フォルダ

    Dim 班員データフォルダ As String
    班員データフォルダ = 月フォルダ & "\★班員データ"
    フォルダ作成 班員データフォルダ

    Dim person As Variant

    For Each person In 担当者リスト

        Dim 個人フォルダ As String
        個人フォルダ = 月フォルダ & "\" & CStr(person) & 月文字 & "月"

        フォルダ作成 個人フォルダ
        フォルダ作成 個人フォルダ & "\通常見積書"
        フォルダ作成 個人フォルダ & "\UR用見積書"

    Next person

    ' ⑥ 個人ファイル作成
    Dim 成功リスト As String
    Dim wbSrc As Workbook, wbNew As Workbook
    Dim savePath As String

    For Each person In 担当者リスト

        savePath = 班員データフォルダ & "\" & 月文字 & "月引き取り(" & CStr(person) & ").xlsm"

        If Dir(savePath) <> "" Then

            Application.ScreenUpdating = True
            Application.DisplayAlerts = True

            If MsgBox("「" & 月文字 & "月引き取り(" & CStr(person) & ").xlsm」は既に存在します。" & vbCrLf &                       "上書きしますか？", vbQuestion + vbYesNo) = vbNo Then

                Application.ScreenUpdating = False
                Application.DisplayAlerts = False
                GoTo NextPerson

            End If

            Application.ScreenUpdating = False
            Application.DisplayAlerts = False

        End If

        Set wbSrc = Workbooks.Open(元フォルダ & "\" & TEMPLATE_個人, ReadOnly:=True)
        wbSrc.SaveCopyAs savePath
        wbSrc.Close SaveChanges:=False
        Set wbSrc = Nothing

        Set wbNew = Workbooks.Open(savePath)

        Dim ws個人 As Worksheet, s As Worksheet
        Set ws個人 = Nothing

        For Each s In wbNew.Worksheets

            If s.Visible = xlSheetVisible Then
                Set ws個人 = s
                Exit For
            End If

        Next s

        If Not ws個人 Is Nothing Then

            ws個人.Range("B1").Value = CStr(person)

            On Error Resume Next
            ws個人.Name = CStr(person)

            If Err.Number <> 0 Then
                Err.Clear
                ws個人.Name = Left(CStr(person), 28)
            End If

            On Error GoTo ErrHandler

        End If

        ' 個人ファイル内に通常見積書/UR用見積書フォルダパスを保存
        Dim 通常Path As String, URPath As String
        通常Path = 月フォルダ & "\" & CStr(person) & 月文字 & "月\通常見積書"
        URPath = 月フォルダ & "\" & CStr(person) & 月文字 & "月\UR用見積書"

        On Error Resume Next
        wbNew.CustomDocumentProperties(FOLDER_PROP_NORMAL).Delete
        wbNew.CustomDocumentProperties(FOLDER_PROP_UR).Delete
        On Error GoTo ErrHandler

        wbNew.CustomDocumentProperties.Add Name:=FOLDER_PROP_NORMAL,             LinkToContent:=False, Type:=msoPropertyTypeString, Value:=通常Path

        wbNew.CustomDocumentProperties.Add Name:=FOLDER_PROP_UR,             LinkToContent:=False, Type:=msoPropertyTypeString, Value:=URPath

        wbNew.Save
        wbNew.Close SaveChanges:=False
        Set wbNew = Nothing

        成功リスト = 成功リスト & "  ○ " & CStr(person) & vbCrLf

NextPerson:
    Next person

    ' ⑦ 合計ファイル作成
    Dim 合計Path As String
    合計Path = 月フォルダ & "\" & 月文字 & "月合計引き取り予定.xlsm"

    If Dir(合計Path) <> "" Then

        Application.ScreenUpdating = True
        Application.DisplayAlerts = True

        If MsgBox("「" & 月文字 & "月合計引き取り予定.xlsm」は既に存在します。上書きしますか？",                   vbQuestion + vbYesNo) = vbNo Then
            GoTo 完了
        End If

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

    End If

    Set wbSrc = Workbooks.Open(元フォルダ & "\" & TEMPLATE_合計, ReadOnly:=True)
    wbSrc.SaveCopyAs 合計Path
    wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing

    Set wbNew = Workbooks.Open(合計Path)

    ' 月数字を各シートに記入
    On Error Resume Next

    If WbSheetExists(wbNew, "担当者別 引取り予定表") Then
        wbNew.Sheets("担当者別 引取り予定表").Range("N1").Value = 月
    End If

    If WbSheetExists(wbNew, "首都圏(仕入品)") Then
        wbNew.Sheets("首都圏(仕入品)").Range("C2").Value = 月
    End If

    If WbSheetExists(wbNew, "首都圏(四工品)") Then
        wbNew.Sheets("首都圏(四工品)").Range("C2").Value = 月
    End If

    On Error GoTo ErrHandler

    ' ★作成した合計ファイル側に、参照する班員データフォルダを保存
    班員データフォルダパス保存 wbNew, 班員データフォルダ

    ' ★クエリを新規作成
    クエリ新規作成 wbNew

    ' ★クエリ出力先を★班員データシートに設定して読み込み
    クエリ出力先設定 wbNew

    wbNew.Save
    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing

完了:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "作成完了！" & vbCrLf & vbCrLf &            "■ 作成フォルダ" & vbCrLf &            "  " & 月フォルダ & vbCrLf & vbCrLf &            "■ 個人ファイル（" & 担当者リスト.Count & "名）" & vbCrLf &            成功リスト & vbCrLf &            "■ 合計ファイル" & vbCrLf &            "  ○ " & 月文字 & "月合計引き取り予定.xlsm", vbInformation, "完了"

    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    On Error Resume Next

    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False

    On Error GoTo 0

    MsgBox "ファイル作成中にエラーが発生しました。" & vbCrLf &            "エラー番号：" & Err.Number & vbCrLf &            "内容：" & Err.Description, vbCritical

End Sub


'====================================================
' クエリ出力先設定
'====================================================
Private Sub クエリ出力先設定(ByVal wb As Workbook)

    Dim wsData As Worksheet

    On Error Resume Next
    Set wsData = wb.Sheets(Q_SHEET)
    On Error GoTo 0

    If wsData Is Nothing Then
        Set wsData = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsData.Name = Q_SHEET
    End If

    ' 既存テーブル削除
    Dim lo As ListObject

    For Each lo In wsData.ListObjects
        lo.Delete
    Next lo

    wsData.Cells.Clear

    ' クエリの存在確認
    Dim q As WorkbookQuery

    On Error Resume Next
    Set q = wb.Queries(Q_MAIN)
    On Error GoTo 0

    If q Is Nothing Then
        MsgBox "クエリ「" & Q_MAIN & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    On Error GoTo クエリエラー

    Dim newLo As ListObject

    Set newLo = wsData.ListObjects.Add(SourceType:=0,         Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & Q_MAIN & ";Extended Properties=""""",         Destination:=wsData.Range("A1"))

    With newLo.QueryTable
        .CommandType = xlCmdDefault
        .CommandText = Array(Q_MAIN)
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = True
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With

    newLo.Name = "TBL_班員データ"

    Exit Sub

クエリエラー:
    MsgBox "クエリ出力先の設定中にエラーが発生しました。" & vbCrLf &            "手動でデータ → クエリと接続 → ★班員データ → 右クリックから「読み込み先」でテーブルを設定してください。" & vbCrLf &            "エラー内容：" & Err.Description, vbExclamation

End Sub


'====================================================
' クエリ新規作成
'====================================================
Private Sub クエリ新規作成(ByVal wb As Workbook)

    On Error Resume Next
    wb.Queries(Q_MAIN).Delete
    On Error GoTo 0

    Dim part1 As String, part2 As String, part3 As String
    Dim part4 As String, part5 As String

    part1 =         "let" & vbCrLf &         "    フォルダ設定 = Excel.CurrentWorkbook(){[Name=""" & NAME_MEMBER_DATA_FOLDER & """]}[Content]," & vbCrLf &         "    FolderPath = Text.From(フォルダ設定{0}[Column1])," & vbCrLf &         "    ソース = Folder.Files(FolderPath)," & vbCrLf &         "    非表示ファイルのフィルタ = Table.SelectRows(ソース, each [Attributes]?[Hidden]? <> true)," & vbCrLf &         "    必要な列の選択 = Table.SelectColumns(非表示ファイルのフィルタ, {""Name"", ""Content""})," & vbCrLf &         "    ブックの展開 = Table.AddColumn(必要な列の選択, ""Data"", each Excel.Workbook([Content], null, true))," & vbCrLf &         "    テーブルの展開 = Table.ExpandTableColumn(ブックの展開, ""Data"", {""Name"", ""Data"", ""Item"", ""Kind""}, {""Sheet.Name"", ""Sheet.Data"", ""Item"", ""Kind""}),"

    part2 = vbCrLf &         "    集計テーブルのフィルタ = Table.SelectRows(テーブルの展開, each ([Item] = ""集計テーブル"" and [Kind] = ""Table""))," & vbCrLf &         "    標準化Data追加 = Table.AddColumn(" & vbCrLf &         "        集計テーブルのフィルタ," & vbCrLf &         "        ""標準化Data""," & vbCrLf &         "        each" & vbCrLf &         "            Table.SelectColumns(" & vbCrLf &         "                [Sheet.Data]," & vbCrLf &         "                {""担当者"", ""得意先"", ""現場"", ""材料"", ""数量"", ""単位"", ""納品日"", ""注文状況"", ""注文日"", ""現場状況"", ""チェック"", ""日時"", ""UR""}," & vbCrLf &         "                MissingField.UseNull" & vbCrLf &         "            )" & vbCrLf &         "    ),"

    part3 = vbCrLf &         "    不要な列の削除 = Table.SelectColumns(標準化Data追加, {""Name"", ""標準化Data""})," & vbCrLf &         "    データの展開 = Table.ExpandTableColumn(" & vbCrLf &         "        不要な列の削除," & vbCrLf &         "        ""標準化Data""," & vbCrLf &         "        {""担当者"", ""得意先"", ""現場"", ""材料"", ""数量"", ""単位"", ""納品日"", ""注文状況"", ""注文日"", ""現場状況"", ""チェック"", ""日時"", ""UR""}," & vbCrLf &         "        {""担当者"", ""得意先"", ""現場"", ""材料"", ""数量"", ""単位"", ""納品日"", ""注文状況"", ""注文日"", ""現場状況"", ""チェック"", ""日時"", ""UR""}" & vbCrLf &         "    ),"

    part4 = vbCrLf &         "    UR整備 = Table.TransformColumns(" & vbCrLf &         "        データの展開," & vbCrLf &         "        {" & vbCrLf &         "            {" & vbCrLf &         "                ""UR""," & vbCrLf &         "                each if _ = null then null else if Text.Contains(Text.Upper(Text.From(_)), ""UR"") then ""UR"" else null," & vbCrLf &         "                type nullable text" & vbCrLf &         "            }" & vbCrLf &         "        }" & vbCrLf &         "    )," & vbCrLf &         "    UR列名変更 = Table.RenameColumns(UR整備, {{""UR"", ""取込UR""}}, MissingField.Ignore)," & vbCrLf &         "    区分追加 = Table.AddColumn(UR列名変更, ""区分"", each if [取込UR] = ""UR"" then ""UR"" else ""通常"", type text),"

    part5 = vbCrLf &         "    型の変更 = Table.TransformColumnTypes(" & vbCrLf &         "        区分追加," & vbCrLf &         "        {" & vbCrLf &         "            {""Name"", type text}," & vbCrLf &         "            {""担当者"", type any}," & vbCrLf &         "            {""得意先"", type any}," & vbCrLf &         "            {""現場"", type any}," & vbCrLf &         "            {""材料"", type any}," & vbCrLf &         "            {""数量"", type any}," & vbCrLf &         "            {""単位"", type any}," & vbCrLf &         "            {""納品日"", type any}," & vbCrLf &         "            {""注文状況"", type any}," & vbCrLf &         "            {""注文日"", type any}," & vbCrLf &         "            {""現場状況"", type any}," & vbCrLf &         "            {""チェック"", type any}," & vbCrLf &         "            {""日時"", type any}," & vbCrLf &         "            {""取込UR"", type text}," & vbCrLf &         "            {""区分"", type text}" & vbCrLf &         "        }" & vbCrLf &         "    )" & vbCrLf &         "in" & vbCrLf &         "    型の変更"

    wb.Queries.Add Name:=Q_MAIN, Formula:=part1 & part2 & part3 & part4 & part5

End Sub


'====================================================
' 作成した合計ファイル側に班員データフォルダパスを保存
'====================================================
Private Sub 班員データフォルダパス保存(ByVal wb As Workbook, ByVal folderPath As String)

    Dim ws As Worksheet

    If Trim(folderPath) = "" Then
        MsgBox "保存する班員データフォルダパスが空です。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set ws = wb.Worksheets(SETTING_SHEET)
    On Error GoTo 0

    ' プログラム設定シートが無い場合は作成
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SETTING_SHEET
    End If

    ' C1 に ★班員データ フォルダのパスを保存
    ws.Range("C1").Value = folderPath

    ' 見出しも入れておく
    ws.Range("B1").Value = NAME_MEMBER_DATA_FOLDER

    ' 既存の名前定義を削除して作り直す
    On Error Resume Next
    wb.Names(NAME_MEMBER_DATA_FOLDER).Delete
    On Error GoTo 0

    ' Power Query の Excel.CurrentWorkbook() で読める形の名前定義にする
    wb.Names.Add         Name:=NAME_MEMBER_DATA_FOLDER,         RefersTo:="='" & ws.Name & "'!$C$1"

End Sub


'====================================================
' 補助：フォルダ作成
'====================================================
Private Sub フォルダ作成(ByVal folderPath As String)

    On Error Resume Next

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    On Error GoTo 0

End Sub


'====================================================
' 補助：指定ブック内のシート存在確認
'====================================================
Private Function WbSheetExists(ByVal wb As Workbook, ByVal sName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Sheets(sName)
    WbSheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function