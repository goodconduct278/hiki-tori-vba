Attribute VB_Name = "Module2"
Option Explicit

' =========================================================
' 一括入力マクロ UR対応・未変換候補表示対応版
'
' 目的：
' ・★変換済みデータから未変換行を弾かない
' ・未変換行も一括入力フォームに表示する
' ・変換候補一覧の候補1?3を一括入力フォームに表示する
' ・採用候補を選んで、製品名へ反映できるようにする
' ・UR対応後の列ズレを修正する
' =========================================================

' =========================================================
' 【事前準備】
' PQの「★変換済みデータ」クエリをシート出力に変更すること
' 「クエリと接続」→ ★変換済みデータ を右クリック
' →「読み込み先」→「テーブル」→ 新しいワークシート → OK
' =========================================================

' ========= シート名定数 =========
Private Const MAIN_SHEET      As String = "担当者別 引取り予定表"
Private Const LOG_SHEET       As String = "入力ログ"
Private Const INPUT_SHEET     As String = "一括入力フォーム"
Private Const TABLE_NAME      As String = "入力フォーム"
Private Const DATA_SHEET      As String = "★班員データ"
Private Const CONV_SHEET      As String = "★変換済みデータ"   ' 追加：PQ出力先シート
Private Const MANAGE_SHEET    As String = "★取込管理"
Private Const UNMATCHED_SHEET As String = "未一致一覧"

' ========= ★班員データ 列定数 =========
' PQの列順：
' Name/担当者/得意先/現場/材料/数量/単位/納品日/注文状況/注文日/現場状況/チェック/日時/取込UR/区分
Private Const DATA_COL_ファイル   As Long = 1
Private Const DATA_COL_担当者     As Long = 2
Private Const DATA_COL_得意先     As Long = 3
Private Const DATA_COL_現場       As Long = 4
Private Const DATA_COL_材料       As Long = 5
Private Const DATA_COL_数量       As Long = 6
Private Const DATA_COL_単位       As Long = 7
Private Const DATA_COL_納品日     As Long = 8
Private Const DATA_COL_注文状況   As Long = 9
Private Const DATA_COL_注文日     As Long = 10
Private Const DATA_COL_現場状況   As Long = 11
Private Const DATA_COL_チェック   As Long = 12
Private Const DATA_COL_日時       As Long = 13
Private Const DATA_COL_取込UR     As Long = 14
Private Const DATA_COL_区分       As Long = 15

' ========= ★変換済みデータ 列定数 =========
' 現在のPQ出力列順：
' A Name
' B 担当者
' C 得意先
' D 現場
' E 区分
' F 材料
' G 変換後製品名
' H メーカー
' I 数量
' J 単位
' K 納品日
' L 注文状況
' M 注文日
' N UR
' O 現場状況
' P チェック
' Q 日時
' R 入力フォーム表示名
' S 変換状態
Private Const CONV_COL_ファイル             As Long = 1
Private Const CONV_COL_担当者               As Long = 2
Private Const CONV_COL_得意先               As Long = 3
Private Const CONV_COL_現場                 As Long = 4
Private Const CONV_COL_区分                 As Long = 5
Private Const CONV_COL_材料                 As Long = 6
Private Const CONV_COL_変換後製品名         As Long = 7
Private Const CONV_COL_メーカー             As Long = 8
Private Const CONV_COL_数量                 As Long = 9
Private Const CONV_COL_単位                 As Long = 10
Private Const CONV_COL_納品日               As Long = 11
Private Const CONV_COL_注文状況             As Long = 12
Private Const CONV_COL_注文日               As Long = 13
Private Const CONV_COL_取込UR               As Long = 14
Private Const CONV_COL_現場状況             As Long = 15
Private Const CONV_COL_チェック             As Long = 16
Private Const CONV_COL_日時                 As Long = 17
Private Const CONV_COL_入力フォーム表示名   As Long = 18
Private Const CONV_COL_変換状態             As Long = 19

' ========= モジュール変数 =========
Private m_ProductMap As Object   ' 変換リストのキャッシュ


' =========================================================
' 共通：文字列の正規化
' =========================================================
Private Function Normalize(ByVal s As Variant) As String
    Dim result As String

    result = CStr(s)

    On Error Resume Next
    result = StrConv(result, vbNarrow)
    On Error GoTo 0

    result = LCase$(result)

    result = Replace(result, " ", "")
    result = Replace(result, "　", "")
    result = Replace(result, Chr(160), "")
    result = Replace(result, vbTab, "")

    result = Replace(result, "（", "(")
    result = Replace(result, "）", ")")
    result = Replace(result, "［", "[")
    result = Replace(result, "］", "]")
    result = Replace(result, "｛", "{")
    result = Replace(result, "｝", "}")
    result = Replace(result, "－", "-")
    result = Replace(result, "―", "-")
    result = Replace(result, "ｰ", "-")
    result = Replace(result, "／", "/")
    result = Replace(result, "・", "")
    result = Replace(result, ".", "")
    result = Replace(result, "．", "")

    result = Replace(result, "ｍｌ", "ml")
    result = Replace(result, "?", "ml")
    result = Replace(result, "ｌ", "l")

    Normalize = Trim$(result)
End Function

' =========================================================
' 共通：キー用の値正規化
' =========================================================
Private Function KeyPart(ByVal v As Variant) As String
    If IsError(v) Then
        KeyPart = "#ERR"
    ElseIf Trim$(CStr(v)) = "" Then
        KeyPart = ""
    ElseIf IsDate(v) Then
        KeyPart = Format$(CDate(v), "yyyymmddhhnnss")
    ElseIf IsNumeric(v) Then
        KeyPart = Replace(Format$(CDbl(v), "0.############"), ",", "")
    Else
        KeyPart = Normalize(v)
    End If
End Function

' =========================================================
' 共通：シートの存在確認
' =========================================================
Private Function SheetExists(ByVal sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' =========================================================
' 共通：テーブルの列インデックスを名前で取得
' =========================================================
Private Function GetTableColumnIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name = colName Then
            GetTableColumnIndex = col.Index
            Exit Function
        End If
    Next col
    GetTableColumnIndex = 0
End Function

' =========================================================
' 共通：入力フォームテーブルに「元データキー」列を確保する
' =========================================================
Private Function EnsureSourceKeyColumn(ByVal tbl As ListObject) As Long
    Dim idx As Long
    Dim lc As ListColumn

    idx = GetTableColumnIndex(tbl, "元データキー")

    If idx = 0 Then
        Set lc = tbl.ListColumns.Add
        lc.Name = "元データキー"
        idx = lc.Index
    End If

    On Error Resume Next
    tbl.ListColumns(idx).Range.EntireColumn.Hidden = True
    On Error GoTo 0

    EnsureSourceKeyColumn = idx
End Function

' =========================================================
' 共通：入力フォームテーブルに列を確保する
' =========================================================
Private Function EnsureTableColumn(ByVal tbl As ListObject, ByVal colName As String, Optional ByVal hideColumn As Boolean = False) As Long
    Dim idx As Long
    Dim lc As ListColumn

    idx = GetTableColumnIndex(tbl, colName)

    If idx = 0 Then
        Set lc = tbl.ListColumns.Add
        lc.Name = colName
        idx = lc.Index
    End If

    If hideColumn Then
        On Error Resume Next
        tbl.ListColumns(idx).Range.EntireColumn.Hidden = True
        On Error GoTo 0
    End If

    EnsureTableColumn = idx
End Function

' =========================================================
' 共通：入力フォーム候補列を確保する
' =========================================================
Private Sub EnsureCandidateColumns(ByVal tbl As ListObject, _
                                   ByRef col候補1 As Long, _
                                   ByRef col候補2 As Long, _
                                   ByRef col候補3 As Long, _
                                   ByRef col採用候補 As Long, _
                                   ByRef col手入力製品名 As Long, _
                                   ByRef col元材料名 As Long)

    col候補1 = EnsureTableColumn(tbl, "候補1")
    col候補2 = EnsureTableColumn(tbl, "候補2")
    col候補3 = EnsureTableColumn(tbl, "候補3")
    col採用候補 = EnsureTableColumn(tbl, "採用候補")
    col手入力製品名 = EnsureTableColumn(tbl, "手入力製品名")
    col元材料名 = EnsureTableColumn(tbl, "元材料名")

    On Error Resume Next
    With tbl.ListColumns(col採用候補).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="候補1,候補2,候補3,手入力,見送り"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

' =========================================================
' 共通：変換候補一覧から候補1?3を読み込む
' 想定列：
' A 元の材料名 / B 区分 / C 候補1 / D スコア1 / E 候補2 / F スコア2 / G 候補3 / H スコア3
' =========================================================
Private Function LoadCandidateMap() As Object

    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim k As String
    Dim arr(1 To 3) As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If Not SheetExists("変換候補一覧") Then
        Set LoadCandidateMap = dict
        Exit Function
    End If

    Set ws = ThisWorkbook.Sheets("変換候補一覧")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 5 Then
        Set LoadCandidateMap = dict
        Exit Function
    End If

    For i = 5 To lastRow
        k = Normalize(ws.Cells(i, "B").Value) & "|" & Normalize(ws.Cells(i, "A").Value)

        If Trim$(Replace(k, "|", "")) <> "" Then
            arr(1) = Trim$(CStr(ws.Cells(i, "C").Value))
            arr(2) = Trim$(CStr(ws.Cells(i, "E").Value))
            arr(3) = Trim$(CStr(ws.Cells(i, "G").Value))

            If Not dict.exists(k) Then
                dict.Add k, Array(arr(1), arr(2), arr(3))
            End If
        End If
    Next i

    Set LoadCandidateMap = dict
End Function

' =========================================================
' 共通：候補マップのキー作成
' =========================================================
Private Function CandidateKey(ByVal kubun As Variant, ByVal materialName As Variant) As String
    CandidateKey = Normalize(kubun) & "|" & Normalize(materialName)
End Function


' =========================================================
' 共通：数値として有効かつゼロ以外かを判定
' =========================================================
Private Function IsNonZeroNumber(ByVal v As Variant) As Boolean
    If IsNumeric(v) Then
        If CDbl(v) <> 0 Then IsNonZeroNumber = True
    End If
End Function

' =========================================================
' 共通：ログテーブルに1行追記
' =========================================================
Private Sub AppendLog(ByVal wsLog As Worksheet, ByVal sPerson As String, ByVal sProduct As String, _
                      ByVal sHalf As String, ByVal nValue As Double)

    Dim tbl As ListObject
    Dim newRow As ListRow

    If wsLog.ListObjects.Count = 0 Then Exit Sub

    Set tbl = wsLog.ListObjects(1)
    Set newRow = tbl.ListRows.Add

    With newRow.Range
        .Cells(1, 1).Value = Now()
        .Cells(1, 1).NumberFormat = "yyyy/mm/dd hh:mm"
        .Cells(1, 2).Value = sPerson
        .Cells(1, 3).Value = sProduct
        .Cells(1, 4).Value = sHalf
        .Cells(1, 5).Value = nValue
    End With
End Sub

' =========================================================
' 共通：変換リストを読み込む
' A列：変換前材料名 / B列：変換後製品名
' =========================================================
Private Sub LoadProductMap(Optional ByVal forceReload As Boolean = False)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim aliasKey As String
    Dim officialName As String
    Dim officialKey As String

    If (Not forceReload) Then
        If Not m_ProductMap Is Nothing Then Exit Sub
    End If

    Set m_ProductMap = CreateObject("Scripting.Dictionary")
    m_ProductMap.CompareMode = vbTextCompare

    If Not SheetExists("変換リスト") Then Exit Sub
    Set ws = ThisWorkbook.Sheets("変換リスト")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For i = 2 To lastRow
        aliasKey = Normalize(ws.Cells(i, "A").Value)
        officialName = Trim$(CStr(ws.Cells(i, "B").Value))
        officialKey = Normalize(officialName)

        If aliasKey <> "" And officialName <> "" Then
            m_ProductMap(aliasKey) = officialName
        End If

        If officialKey <> "" And officialName <> "" Then
            If Not m_ProductMap.exists(officialKey) Then
                m_ProductMap.Add officialKey, officialName
            End If
        End If
    Next i
End Sub

' =========================================================
' 共通：材料名/製品名を正式名へ寄せる
' =========================================================
Private Function CanonicalProductName(ByVal s As String) As String
    Dim k As String

    If m_ProductMap Is Nothing Then LoadProductMap

    k = Normalize(s)

    If k = "" Then
        CanonicalProductName = ""
    ElseIf Not m_ProductMap Is Nothing And m_ProductMap.exists(k) Then
        CanonicalProductName = Trim$(CStr(m_ProductMap(k)))
    Else
        CanonicalProductName = Trim$(s)
    End If
End Function

' =========================================================
' 変換リストで材料名を製品名に変換する（見積書取込用）
' =========================================================
Private Function 材料名変換(ByVal 材料名 As String) As String
    材料名変換 = CanonicalProductName(材料名)
End Function

' =========================================================
' 共通：担当者の列番号を検索する関数
' =========================================================
Private Function GetPersonCol(ByVal wsMain As Worksheet, ByVal sPerson As String, ByVal sHalf As String) As Long
    Const hRow As Long = 2
    Dim lastCol As Long
    Dim col As Long
    Dim normKey As String

    normKey = Normalize(sPerson)
    If normKey = "" Then
        GetPersonCol = 0
        Exit Function
    End If

    lastCol = wsMain.Cells(hRow, wsMain.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol
        If Normalize(wsMain.Cells(hRow, col).Value) = normKey Then
            GetPersonCol = col + IIf(sHalf = "後半", 1, 0)
            Exit Function
        End If
    Next col

    GetPersonCol = 0
End Function

' =========================================================
' 共通：製品の行番号を検索する関数（正式名＋正規化比較）
' =========================================================
Private Function FindProductRow(ByVal wsMain As Worksheet, ByVal sProduct As String) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim normKey As String
    Dim cellKey As String

    normKey = Normalize(CanonicalProductName(sProduct))
    If normKey = "" Then
        FindProductRow = 0
        Exit Function
    End If

    lastRow = wsMain.Cells(wsMain.Rows.Count, 2).End(xlUp).Row
    If lastRow < 4 Then
        FindProductRow = 0
        Exit Function
    End If

    For r = 4 To lastRow
        cellKey = Normalize(CanonicalProductName(wsMain.Cells(r, 2).Value))
        If cellKey = normKey Then
            FindProductRow = r
            Exit Function
        End If
    Next r

    FindProductRow = 0
End Function

' =========================================================
' 共通：★班員データの1行から管理用キーを作成
' =========================================================
Private Function BuildSourceKey(ByVal wsData As Worksheet, ByVal rowNum As Long) As String

    BuildSourceKey = _
          "f=" & KeyPart(wsData.Cells(rowNum, DATA_COL_ファイル).Value) _
        & "|p=" & KeyPart(wsData.Cells(rowNum, DATA_COL_担当者).Value) _
        & "|c=" & KeyPart(wsData.Cells(rowNum, DATA_COL_得意先).Value) _
        & "|g=" & KeyPart(wsData.Cells(rowNum, DATA_COL_現場).Value) _
        & "|m=" & KeyPart(wsData.Cells(rowNum, DATA_COL_材料).Value) _
        & "|q=" & KeyPart(wsData.Cells(rowNum, DATA_COL_数量).Value) _
        & "|u=" & KeyPart(wsData.Cells(rowNum, DATA_COL_単位).Value) _
        & "|d=" & KeyPart(wsData.Cells(rowNum, DATA_COL_納品日).Value) _
        & "|o=" & KeyPart(wsData.Cells(rowNum, DATA_COL_注文日).Value) _
        & "|t=" & KeyPart(wsData.Cells(rowNum, DATA_COL_日時).Value)

End Function

' =========================================================
' 共通：★変換済みデータの配列行から管理用キーを作成
' ※ 元材料名ベースで生成するので★取込管理との整合性が保たれる
' =========================================================
Private Function BuildSourceKeyFromConv(ByVal dataArr As Variant, ByVal i As Long) As String

    BuildSourceKeyFromConv = _
          "f=" & KeyPart(dataArr(i, CONV_COL_ファイル)) _
        & "|p=" & KeyPart(dataArr(i, CONV_COL_担当者)) _
        & "|c=" & KeyPart(dataArr(i, CONV_COL_得意先)) _
        & "|g=" & KeyPart(dataArr(i, CONV_COL_現場)) _
        & "|m=" & KeyPart(dataArr(i, CONV_COL_材料)) _
        & "|q=" & KeyPart(dataArr(i, CONV_COL_数量)) _
        & "|u=" & KeyPart(dataArr(i, CONV_COL_単位)) _
        & "|d=" & KeyPart(dataArr(i, CONV_COL_納品日)) _
        & "|o=" & KeyPart(dataArr(i, CONV_COL_注文日)) _
        & "|t=" & KeyPart(dataArr(i, CONV_COL_日時))

End Function

' =========================================================
' 共通：★取込管理シートを取得（なければ作成）
' =========================================================
Private Function EnsureManageSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(MANAGE_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = MANAGE_SHEET
    End If

    With ws
        If Application.WorksheetFunction.CountA(.Rows(1)) = 0 Then
            .Cells(1, 1).Value = "元データキー"
            .Cells(1, 2).Value = "取込日時"
            .Cells(1, 3).Value = "担当者"
            .Cells(1, 4).Value = "製品名"
            .Cells(1, 5).Value = "前半/後半"
            .Cells(1, 6).Value = "数量"
            .Cells(1, 7).Value = "処理区分"
            .Cells(1, 8).Value = "メモ"

            .Rows(1).Font.Bold = True
            .Rows(1).Interior.Color = RGB(226, 239, 218)
            .Columns("A:H").AutoFit
        End If
    End With

    Set EnsureManageSheet = ws
End Function

' =========================================================
' 共通：★取込管理のキー一覧を読み込む
' =========================================================
Private Function LoadImportedKeyMap() As Object

    Dim ws As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim k As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Set ws = EnsureManageSheet()

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        Set LoadImportedKeyMap = dict
        Exit Function
    End If

    For i = 2 To lastRow
        k = Trim$(CStr(ws.Cells(i, 1).Value))
        If k <> "" Then
            If Not dict.exists(k) Then
                dict.Add k, i
            Else
                dict(k) = i
            End If
        End If
    Next i

    Set LoadImportedKeyMap = dict
End Function

' =========================================================
' 共通：元データキーを★取込管理へ記録する（Upsert）
' =========================================================
Private Sub MarkSourceKeyImported(ByVal srcKey As String, _
                                  ByVal sPerson As String, _
                                  ByVal sProduct As String, _
                                  ByVal sHalf As String, _
                                  ByVal vQty As Variant, _
                                  Optional ByVal processName As String = "転写完了", _
                                  Optional ByVal note As String = "")

    Dim ws As Worksheet
    Dim dict As Object
    Dim targetRow As Long

    If Trim$(srcKey) = "" Then Exit Sub

    Set ws = EnsureManageSheet()
    Set dict = LoadImportedKeyMap()

    If dict.exists(srcKey) Then
        targetRow = CLng(dict(srcKey))
    Else
        targetRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If targetRow < 2 Then targetRow = 2
    End If

    With ws
        .Cells(targetRow, 1).Value = srcKey
        .Cells(targetRow, 2).Value = Now()
        .Cells(targetRow, 2).NumberFormat = "yyyy/mm/dd hh:mm"
        .Cells(targetRow, 3).Value = sPerson
        .Cells(targetRow, 4).Value = sProduct
        .Cells(targetRow, 5).Value = sHalf
        .Cells(targetRow, 6).Value = vQty
        .Cells(targetRow, 7).Value = processName
        .Cells(targetRow, 8).Value = note
    End With
End Sub

' =========================================================
' 共通：未一致一覧シートを取得（なければ作成）
' =========================================================
Private Function EnsureUnmatchedSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(UNMATCHED_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = UNMATCHED_SHEET
    End If

    With ws
        If Application.WorksheetFunction.CountA(.Rows(1)) = 0 Then
            .Cells(1, 1).Value = "記録日時"
            .Cells(1, 2).Value = "処理種別"
            .Cells(1, 3).Value = "担当者"
            .Cells(1, 4).Value = "入力製品名"
            .Cells(1, 5).Value = "変換後製品名"
            .Cells(1, 6).Value = "正規化キー"
            .Cells(1, 7).Value = "前半/後半"
            .Cells(1, 8).Value = "数量"
            .Cells(1, 9).Value = "元データキー"
            .Cells(1, 10).Value = "メモ"

            .Rows(1).Font.Bold = True
            .Rows(1).Interior.Color = RGB(220, 230, 241)
            .Columns("A:J").AutoFit
        End If
    End With

    Set EnsureUnmatchedSheet = ws
End Function

' =========================================================
' 共通：未一致一覧に1件追記（転写処理失敗時用）
' =========================================================
Private Sub AppendUnmatchedProduct(ByVal processName As String, _
                                   ByVal sPerson As String, _
                                   ByVal sProduct As String, _
                                   ByVal sHalf As String, _
                                   ByVal vQty As Variant, _
                                   Optional ByVal srcKey As String = "", _
                                   Optional ByVal note As String = "")

    Dim ws As Worksheet
    Dim NextRow As Long
    Dim canonicalName As String
    Dim normKey As String

    Set ws = EnsureUnmatchedSheet()

    canonicalName = CanonicalProductName(sProduct)
    normKey = Normalize(canonicalName)

    With ws
        NextRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        If NextRow < 2 Then NextRow = 2

        .Cells(NextRow, 1).Value = Now()
        .Cells(NextRow, 1).NumberFormat = "yyyy/mm/dd hh:mm"
        .Cells(NextRow, 2).Value = processName
        .Cells(NextRow, 3).Value = sPerson
        .Cells(NextRow, 4).Value = sProduct
        .Cells(NextRow, 5).Value = canonicalName
        .Cells(NextRow, 6).Value = normKey
        .Cells(NextRow, 7).Value = sHalf
        .Cells(NextRow, 8).Value = vQty
        .Cells(NextRow, 9).Value = srcKey
        .Cells(NextRow, 10).Value = note
    End With
End Sub

' =========================================================
' 共通：テーブルの入力列をすべてクリアする
' =========================================================
Private Sub ClearInputTable(ByVal tbl As ListObject, ByVal colProduct As Long, ByVal colHalf As Long, _
                            ByVal colQty As Long, ByVal colResult As Long, _
                            Optional ByVal colSourceKey As Long = 0)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    tbl.ListColumns(colProduct).DataBodyRange.ClearContents
    tbl.ListColumns(colHalf).DataBodyRange.ClearContents
    tbl.ListColumns(colQty).DataBodyRange.ClearContents
    tbl.ListColumns(colResult).DataBodyRange.ClearContents

    If colSourceKey > 0 Then
        tbl.ListColumns(colSourceKey).DataBodyRange.ClearContents
    End If
End Sub

' =========================================================
' 共通：テーブルの全行削除
' =========================================================
Private Sub DeleteAllTableRows(ByVal tbl As ListObject)
    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = tbl.ListRows.Count To 1 Step -1
        tbl.ListRows(i).Delete
    Next i
End Sub

' =========================================================
' 共通：テーブルから成功行・空行を後ろから削除して詰める
' =========================================================
Private Sub CompactTableRows(ByVal tbl As ListObject, ByVal colProduct As Long, ByVal colResult As Long)
    Dim i As Long
    Dim status As String, prodVal As String

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = tbl.ListRows.Count To 1 Step -1
        prodVal = Trim$(CStr(tbl.ListRows(i).Range.Cells(1, colProduct).Value))
        status = Trim$(CStr(tbl.ListRows(i).Range.Cells(1, colResult).Value))

        If prodVal = "" Then
            tbl.ListRows(i).Delete
        ElseIf InStr(1, status, "転写完了", vbTextCompare) = 1 _
            Or InStr(1, status, "合算完了", vbTextCompare) = 1 Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

' =========================================================
' 共通：PQクエリを同期更新する
' =========================================================
Private Sub RefreshPowerQuery()
    Dim conn As WorkbookConnection

    Application.StatusBar = "PQを更新中..."

    On Error Resume Next
    For Each conn In ThisWorkbook.Connections
        conn.OLEDBConnection.BackgroundQuery = False
        conn.Refresh
    Next conn
    On Error GoTo 0

    Application.StatusBar = False
End Sub


' =========================================================
' 【機能1】一括入力フォーム：リストを合算転写してログを残す
' =========================================================
Sub 一括入力フォームから転写する()

    Dim wsInput As Worksheet, wsMain As Worksheet, wsLog As Worksheet
    Dim tbl As ListObject
    Dim rowRng As Range

    Dim sPerson As String, sProduct As String, sHalf As String
    Dim nValue As Variant, currentValue As Variant
    Dim targetRow As Long, targetCol As Long
    Dim i As Long
    Dim successCount As Long, errorCount As Long
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    Dim askConfirm As Boolean
    Dim doMerge As Boolean
    Dim hasInput As Boolean

    Dim colProduct As Long, colHalf As Long, colQty As Long, colResult As Long
    Dim colSourceKey As Long
    Dim srcKey As String

    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    On Error GoTo ErrHandler

    If Not SheetExists(INPUT_SHEET) Or Not SheetExists(MAIN_SHEET) Or Not SheetExists(LOG_SHEET) Then
        MsgBox "必要なシートが見つかりません。", vbCritical
        Exit Sub
    End If

    LoadProductMap True

    Set wsInput = ThisWorkbook.Sheets(INPUT_SHEET)
    Set wsMain = ThisWorkbook.Sheets(MAIN_SHEET)
    Set wsLog = ThisWorkbook.Sheets(LOG_SHEET)

    On Error Resume Next
    Set tbl = wsInput.ListObjects(TABLE_NAME)
    On Error GoTo ErrHandler

    If tbl Is Nothing Then
        MsgBox "入力フォームのテーブル「" & TABLE_NAME & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    colProduct = GetTableColumnIndex(tbl, "製品名")
    colHalf = GetTableColumnIndex(tbl, "前半/後半")
    colQty = GetTableColumnIndex(tbl, "数量")
    colResult = GetTableColumnIndex(tbl, "処理結果（自動）")
    colSourceKey = EnsureSourceKeyColumn(tbl)

    If colProduct = 0 Or colHalf = 0 Or colQty = 0 Or colResult = 0 Then
        MsgBox "テーブルの見出し名を確認してください。" & vbCrLf & _
               "必要: 製品名 / 前半/後半 / 数量 / 処理結果（自動）", vbCritical
        Exit Sub
    End If

    sPerson = Trim$(CStr(wsInput.Range("C3").Value))
    If sPerson = "" Then
        MsgBox "担当者を先に選択してください。", vbExclamation
        Exit Sub
    End If

    hasInput = False
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            Set rowRng = tbl.ListRows(i).Range
            If Trim$(CStr(rowRng.Cells(1, colProduct).Value)) <> "" _
            Or Trim$(CStr(rowRng.Cells(1, colHalf).Value)) <> "" _
            Or Trim$(CStr(rowRng.Cells(1, colQty).Value)) <> "" Then
                hasInput = True
                Exit For
            End If
        Next i
    End If

    If Not hasInput Then
        MsgBox "製品リストが入力されていません。", vbInformation
        Exit Sub
    End If

    askConfirm = (wsInput.Range("F3").Value = True)

    If MsgBox("担当者「" & sPerson & "」のデータを一気に転写しますか？" & vbCrLf & _
              "合算確認：" & IIf(askConfirm, "ON（1件ずつ確認）", "OFF（自動合算）"), _
              vbQuestion + vbYesNo, "確認") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If Not tbl.DataBodyRange Is Nothing Then
        tbl.ListColumns(colResult).DataBodyRange.ClearContents
    End If

    successCount = 0
    errorCount = 0

    For i = 1 To tbl.ListRows.Count
        Set rowRng = tbl.ListRows(i).Range

        sProduct = Trim$(CStr(rowRng.Cells(1, colProduct).Value))
        sHalf = Trim$(CStr(rowRng.Cells(1, colHalf).Value))
        nValue = rowRng.Cells(1, colQty).Value
        srcKey = Trim$(CStr(rowRng.Cells(1, colSourceKey).Value))

        If sProduct = "" And sHalf = "" And Trim$(CStr(nValue)) = "" Then
            rowRng.Cells(1, colResult).ClearContents
            GoTo NextRow
        End If

        If sProduct = "" Then
            rowRng.Cells(1, colResult).Value = "製品名なし"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        If sHalf = "" Then
            rowRng.Cells(1, colResult).Value = "未選択"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        If Not IsNumeric(nValue) Or Trim$(CStr(nValue)) = "" Then
            rowRng.Cells(1, colResult).Value = "数量不正"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        targetRow = FindProductRow(wsMain, sProduct)
        If targetRow = 0 Then
            rowRng.Cells(1, colResult).Value = "製品なし"
            AppendUnmatchedProduct "一括転写", sPerson, sProduct, sHalf, nValue, _
                                   srcKey, _
                                   "担当者別 引取り予定表の製品行が見つかりません"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        targetCol = GetPersonCol(wsMain, sPerson, sHalf)
        If targetCol = 0 Then
            rowRng.Cells(1, colResult).Value = "担当者なし"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        currentValue = wsMain.Cells(targetRow, targetCol).Value

        If IsNonZeroNumber(currentValue) Then
            doMerge = True
            If askConfirm Then
                Application.ScreenUpdating = True
                If MsgBox(sProduct & "（" & sHalf & "）に既に「" & currentValue & "」があります。" & vbCrLf & _
                          "今回の「" & nValue & "」を合算しますか？", _
                          vbExclamation + vbYesNo, "合算確認") = vbNo Then
                    doMerge = False
                End If
                Application.ScreenUpdating = False
            End If

            If doMerge Then
                wsMain.Cells(targetRow, targetCol).Value = CDbl(currentValue) + CDbl(nValue)
                rowRng.Cells(1, colResult).Value = "合算完了 (元:" & currentValue & ")"
                successCount = successCount + 1
                AppendLog wsLog, sPerson, sProduct, sHalf, CDbl(nValue)
                MarkSourceKeyImported srcKey, sPerson, sProduct, sHalf, nValue, "合算完了"
            Else
                rowRng.Cells(1, colResult).Value = "合算スキップ"
                errorCount = errorCount + 1
            End If
        Else
            wsMain.Cells(targetRow, targetCol).Value = CDbl(nValue)
            rowRng.Cells(1, colResult).Value = "転写完了"
            successCount = successCount + 1
            AppendLog wsLog, sPerson, sProduct, sHalf, CDbl(nValue)
            MarkSourceKeyImported srcKey, sPerson, sProduct, sHalf, nValue, "転写完了"
        End If

NextRow:
    Next i

    Call CompactTableRows(tbl, colProduct, colResult)

    Application.Calculation = prevCalculation
    Application.ScreenUpdating = prevScreenUpdating

    If errorCount > 0 Then
        MsgBox successCount & "件完了（エラー・スキップ " & errorCount & "件）。" & vbCrLf & _
               "完了行を削除しました。残行を確認してください。", vbExclamation
    Else
        MsgBox "全件完了しました。完了行を削除しました。", vbInformation, "完了"
    End If

    Exit Sub

ErrHandler:
    Application.Calculation = prevCalculation
    Application.ScreenUpdating = prevScreenUpdating
    MsgBox "エラーが発生しました（テーブル " & i & " 行目付近）：" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub


' =========================================================
' 【機能2】「製品なし」行だけを再転写し、成功行を削除して詰める
' =========================================================
Sub 製品なし行を再転写する()

    Dim wsInput As Worksheet, wsMain As Worksheet, wsLog As Worksheet
    Dim tbl As ListObject
    Dim rowRng As Range

    Dim sPerson As String, sProduct As String, sHalf As String
    Dim nValue As Variant, currentValue As Variant
    Dim targetRow As Long, targetCol As Long
    Dim i As Long
    Dim successCount As Long, errorCount As Long, targetCount As Long
    Dim askConfirm As Boolean
    Dim doMerge As Boolean

    Dim colProduct As Long, colHalf As Long, colQty As Long, colResult As Long
    Dim colSourceKey As Long

    Dim prevSU As Boolean
    Dim prevCalc As XlCalculation
    Dim srcKey As String

    prevSU = Application.ScreenUpdating
    prevCalc = Application.Calculation

    On Error GoTo ErrHandler

    If Not SheetExists(INPUT_SHEET) Or Not SheetExists(MAIN_SHEET) Or Not SheetExists(LOG_SHEET) Then
        MsgBox "必要なシートが見つかりません。", vbCritical
        Exit Sub
    End If

    LoadProductMap True

    Set wsInput = ThisWorkbook.Sheets(INPUT_SHEET)
    Set wsMain = ThisWorkbook.Sheets(MAIN_SHEET)
    Set wsLog = ThisWorkbook.Sheets(LOG_SHEET)

    On Error Resume Next
    Set tbl = wsInput.ListObjects(TABLE_NAME)
    On Error GoTo ErrHandler

    If tbl Is Nothing Then
        MsgBox "入力フォームのテーブル「" & TABLE_NAME & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    colProduct = GetTableColumnIndex(tbl, "製品名")
    colHalf = GetTableColumnIndex(tbl, "前半/後半")
    colQty = GetTableColumnIndex(tbl, "数量")
    colResult = GetTableColumnIndex(tbl, "処理結果（自動）")
    colSourceKey = EnsureSourceKeyColumn(tbl)

    If colProduct = 0 Or colHalf = 0 Or colQty = 0 Or colResult = 0 Then
        MsgBox "テーブルの見出し名を確認してください。", vbCritical
        Exit Sub
    End If

    sPerson = Trim$(CStr(wsInput.Range("C3").Value))
    If sPerson = "" Then
        MsgBox "担当者を先に選択してください。", vbExclamation
        Exit Sub
    End If

    If tbl.DataBodyRange Is Nothing Then
        MsgBox "製品リストが入力されていません。", vbInformation
        Exit Sub
    End If

    targetCount = 0
    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListRows(i).Range.Cells(1, colResult).Value)) = "製品なし" Then
            targetCount = targetCount + 1
        End If
    Next i

    If targetCount = 0 Then
        MsgBox "「製品なし」の行が見つかりません。" & vbCrLf & _
               "先に「一気に転写する」を実行してください。", vbInformation
        Exit Sub
    End If

    If MsgBox(targetCount & "件の「製品なし」行を再転写しますか？" & vbCrLf & _
              "※修正した製品名で再検索します。" & vbCrLf & _
              "※成功した行は自動で削除され、残りは上に詰められます。", _
              vbQuestion + vbYesNo, "確認") = vbNo Then
        Exit Sub
    End If

    askConfirm = (wsInput.Range("F3").Value = True)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    successCount = 0
    errorCount = 0

    For i = 1 To tbl.ListRows.Count
        Set rowRng = tbl.ListRows(i).Range

        If Trim$(CStr(rowRng.Cells(1, colResult).Value)) <> "製品なし" Then GoTo NextRow

        sProduct = Trim$(CStr(rowRng.Cells(1, colProduct).Value))
        sHalf = Trim$(CStr(rowRng.Cells(1, colHalf).Value))
        nValue = rowRng.Cells(1, colQty).Value
        srcKey = Trim$(CStr(rowRng.Cells(1, colSourceKey).Value))

        If sProduct = "" Then
            rowRng.Cells(1, colResult).Value = "製品名空欄"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        If sHalf = "" Then
            rowRng.Cells(1, colResult).Value = "未選択"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        If Not IsNumeric(nValue) Or Trim$(CStr(nValue)) = "" Then
            rowRng.Cells(1, colResult).Value = "数量不正"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        targetRow = FindProductRow(wsMain, sProduct)
        If targetRow = 0 Then
            rowRng.Cells(1, colResult).Value = "製品なし"
            AppendUnmatchedProduct "再転写", sPerson, sProduct, sHalf, nValue, _
                                   srcKey, _
                                   "再転写でも製品行が見つかりません"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        targetCol = GetPersonCol(wsMain, sPerson, sHalf)
        If targetCol = 0 Then
            rowRng.Cells(1, colResult).Value = "担当者なし"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        currentValue = wsMain.Cells(targetRow, targetCol).Value

        If IsNonZeroNumber(currentValue) Then
            doMerge = True
            If askConfirm Then
                Application.ScreenUpdating = True
                If MsgBox(sProduct & "（" & sHalf & "）に既に「" & currentValue & "」があります。" & vbCrLf & _
                          "今回の「" & nValue & "」を合算しますか？", _
                          vbExclamation + vbYesNo, "合算確認") = vbNo Then
                    doMerge = False
                End If
                Application.ScreenUpdating = False
            End If

            If doMerge Then
                wsMain.Cells(targetRow, targetCol).Value = CDbl(currentValue) + CDbl(nValue)
                rowRng.Cells(1, colResult).Value = "合算完了 (元:" & currentValue & ")"
                successCount = successCount + 1
                AppendLog wsLog, sPerson, sProduct, sHalf, CDbl(nValue)
                MarkSourceKeyImported srcKey, sPerson, sProduct, sHalf, nValue, "合算完了"
            Else
                rowRng.Cells(1, colResult).Value = "合算スキップ"
                errorCount = errorCount + 1
            End If
        Else
            wsMain.Cells(targetRow, targetCol).Value = CDbl(nValue)
            rowRng.Cells(1, colResult).Value = "転写完了"
            successCount = successCount + 1
            AppendLog wsLog, sPerson, sProduct, sHalf, CDbl(nValue)
            MarkSourceKeyImported srcKey, sPerson, sProduct, sHalf, nValue, "転写完了"
        End If

NextRow:
    Next i

    Call CompactTableRows(tbl, colProduct, colResult)

    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevSU

    MsgBox "再転写完了：成功 " & successCount & " 件／エラー・スキップ " & errorCount & " 件" & vbCrLf & _
           "完了行を削除し、残りを上に詰めました。", vbInformation
    Exit Sub

ErrHandler:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevSU
    MsgBox "エラーが発生しました（テーブル " & i & " 行目付近）：" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub


' =========================================================
' 【機能3】専用ログシートのテーブルデータをクリアする
' =========================================================
Sub ログをクリアする()

    Dim wsLog As Worksheet
    Dim tbl As ListObject

    On Error GoTo ErrHandler

    If Not SheetExists(LOG_SHEET) Then
        MsgBox "ログシートが見つかりません。", vbCritical
        Exit Sub
    End If

    Set wsLog = ThisWorkbook.Sheets(LOG_SHEET)

    If wsLog.ListObjects.Count = 0 Then
        MsgBox "ログシートにテーブルが見つかりません。", vbExclamation
        Exit Sub
    End If

    Set tbl = wsLog.ListObjects(1)

    If tbl.DataBodyRange Is Nothing Then
        MsgBox "削除するログがありません。", vbInformation
        Exit Sub
    End If

    If MsgBox("これまでの転写ログをすべて削除しますか？", _
              vbExclamation + vbYesNo, "ログ削除確認") = vbYes Then
        tbl.DataBodyRange.Delete
        MsgBox "ログをクリアしました！", vbInformation, "完了"
    End If
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました：" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub


' =========================================================
' 【機能4】見積書ファイルから一括入力フォームへ取り込み
' =========================================================
Sub 見積書から取り込む()

    Dim wsInput As Worksheet
    Dim tbl As ListObject
    Dim srcBook As Workbook, srcWs As Worksheet
    Dim filePath As Variant
    Dim startRow As Long, i As Long, lastRow As Long
    Dim sMat As String, vQty As Variant
    Dim dict As Object
    Dim normKey As String
    Dim readCount As Long, mergeCount As Long
    Dim prevScreenUpdating As Boolean
    Dim entry As Variant
    Dim newRow As ListRow
    Dim key As Variant, item As Variant
    Dim canonicalName As String

    Dim colProduct As Long, colQty As Long, colSourceKey As Long

    On Error GoTo ErrHandler

    If Not SheetExists(INPUT_SHEET) Then
        MsgBox "一括入力フォームが見つかりません。", vbCritical
        Exit Sub
    End If

    LoadProductMap True

    Set wsInput = ThisWorkbook.Sheets(INPUT_SHEET)

    On Error Resume Next
    Set tbl = wsInput.ListObjects(TABLE_NAME)
    On Error GoTo ErrHandler

    If tbl Is Nothing Then
        MsgBox "入力フォームのテーブル「" & TABLE_NAME & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    colProduct = GetTableColumnIndex(tbl, "製品名")
    colQty = GetTableColumnIndex(tbl, "数量")
    colSourceKey = EnsureSourceKeyColumn(tbl)

    If colProduct = 0 Or colQty = 0 Then
        MsgBox "テーブルに「製品名」または「数量」列が見つかりません。", vbCritical
        Exit Sub
    End If

    filePath = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="見積書ファイルを選択してください")
    If filePath = False Then Exit Sub

    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set srcBook = Workbooks.Open(Filename:=CStr(filePath), ReadOnly:=True)
    Set srcWs = srcBook.ActiveSheet

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    startRow = 21
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    If lastRow < startRow Then
        MsgBox "見積書に有効なデータが見つかりません。", vbExclamation
        GoTo CleanUp
    End If

    readCount = 0
    mergeCount = 0

    For i = startRow To lastRow
        sMat = Trim$(CStr(srcWs.Cells(i, 1).Value))
        vQty = srcWs.Cells(i, 13).Value

        If sMat = "" Then GoTo NextRow
        If InStr(sMat, "小計") > 0 Then GoTo NextRow
        If Not IsNumeric(vQty) Or Trim$(CStr(vQty)) = "" Then GoTo NextRow
        If CDbl(vQty) = 0 Then GoTo NextRow

        canonicalName = 材料名変換(sMat)
        normKey = Normalize(canonicalName)

        If dict.exists(normKey) Then
            entry = dict(normKey)
            entry(1) = entry(1) + CDbl(vQty)
            dict(normKey) = entry
            mergeCount = mergeCount + 1
        Else
            dict.Add normKey, Array(canonicalName, CDbl(vQty))
        End If

        readCount = readCount + 1
NextRow:
    Next i

    srcBook.Close SaveChanges:=False
    Set srcBook = Nothing

    If dict.Count = 0 Then
        MsgBox "取り込める材料が見つかりませんでした。", vbExclamation
        GoTo CleanUp
    End If

    Call DeleteAllTableRows(tbl)

    For Each key In dict.Keys
        item = dict(key)
        Set newRow = tbl.ListRows.Add
        newRow.Range.Cells(1, colProduct).Value = CStr(item(0))
        newRow.Range.Cells(1, colQty).Value = item(1)
        newRow.Range.Cells(1, colSourceKey).ClearContents
    Next key

    Application.ScreenUpdating = prevScreenUpdating

    MsgBox "取り込み完了：" & dict.Count & "件" & vbCrLf & _
           "（読み込み " & readCount & " 行、合算 " & mergeCount & " 回）" & vbCrLf & _
           "※前半/後半を入力してから「一気に転写する」を押してください。", _
           vbInformation, "完了"
    Exit Sub

CleanUp:
    If Not srcBook Is Nothing Then srcBook.Close SaveChanges:=False
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    If Not srcBook Is Nothing Then srcBook.Close SaveChanges:=False
    Application.ScreenUpdating = prevScreenUpdating
    MsgBox "エラーが発生しました：" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub


' =========================================================
' 【機能5】★変換済みデータから担当者名でデータを抽出する
' ・PQで変換済みの製品名をそのまま使用
' ・抽出前にPQを自動更新
' ・取込済み判定は★取込管理のキーで行う
' =========================================================
Sub 担当者データ抽出()

    Dim wsInput As Worksheet, wsConv As Worksheet
    Dim tbl As ListObject
    Dim targetName As String
    Dim lastRow As Long, i As Long
    Dim dataArr As Variant
    Dim hitCount As Long, skipCount As Long
    Dim newRow As ListRow
    Dim convName As String
    Dim colProduct As Long, colQty As Long, colResult As Long, colSourceKey As Long
    Dim col候補1 As Long, col候補2 As Long, col候補3 As Long
    Dim col採用候補 As Long, col手入力製品名 As Long, col元材料名 As Long
    Dim msg As String
    Dim srcKey As String
    Dim importedMap As Object
    Dim candidateMap As Object
    Dim candidateArr As Variant
    Dim cKey As String
    Dim transformState As String
    Dim originalMaterial As String
    Dim kubun As String

    On Error GoTo ErrHandler

    If Not SheetExists(INPUT_SHEET) Then
        MsgBox "一括入力フォームが見つかりません。", vbCritical
        Exit Sub
    End If

    If Not SheetExists(CONV_SHEET) Then
        MsgBox "「" & CONV_SHEET & "」シートが見つかりません。" & vbCrLf & _
               "PQの「★変換済みデータ」クエリをシート出力に変更してください。", vbCritical
        Exit Sub
    End If

    Call RefreshPowerQuery

    Set wsInput = ThisWorkbook.Sheets(INPUT_SHEET)
    Set wsConv = ThisWorkbook.Sheets(CONV_SHEET)
    Set importedMap = LoadImportedKeyMap()
    Set candidateMap = LoadCandidateMap()

    On Error Resume Next
    Set tbl = wsInput.ListObjects(TABLE_NAME)
    On Error GoTo ErrHandler

    If tbl Is Nothing Then
        MsgBox "テーブル「" & TABLE_NAME & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    colProduct = GetTableColumnIndex(tbl, "製品名")
    colQty = GetTableColumnIndex(tbl, "数量")
    colResult = GetTableColumnIndex(tbl, "処理結果（自動）")
    colSourceKey = EnsureSourceKeyColumn(tbl)
    Call EnsureCandidateColumns(tbl, col候補1, col候補2, col候補3, col採用候補, col手入力製品名, col元材料名)

    If colProduct = 0 Or colQty = 0 Then
        MsgBox "テーブルに「製品名」または「数量」列が見つかりません。", vbCritical
        Exit Sub
    End If

    targetName = Trim$(CStr(wsInput.Range("C3").Value))
    If targetName = "" Then
        MsgBox "担当者を選択してください。", vbExclamation
        Exit Sub
    End If

    lastRow = wsConv.Cells(wsConv.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "★変換済みデータにデータがありません。" & vbCrLf & _
               "先にマクロでデータを取り込んでください。", vbInformation
        Exit Sub
    End If

    If wsConv.Cells(1, CONV_COL_入力フォーム表示名).Value <> "入力フォーム表示名" _
       Or wsConv.Cells(1, CONV_COL_変換状態).Value <> "変換状態" Then
        MsgBox "★変換済みデータの列構成が想定と違います。" & vbCrLf & _
               "想定：A～S列まで存在し、R列=入力フォーム表示名、S列=変換状態", vbCritical
        Exit Sub
    End If

    dataArr = wsConv.Range("A2:S" & lastRow).Value

    hitCount = 0
    skipCount = 0

    For i = 1 To UBound(dataArr, 1)
        If Trim$(CStr(dataArr(i, CONV_COL_担当者))) = targetName Then
            srcKey = BuildSourceKeyFromConv(dataArr, i)
            If importedMap.exists(srcKey) Then
                skipCount = skipCount + 1
            Else
                hitCount = hitCount + 1
            End If
        End If
    Next i

    If hitCount = 0 And skipCount > 0 Then
        MsgBox targetName & " さんのデータはすべて取込済みです。（スキップ：" & skipCount & "件）", vbInformation
        Exit Sub
    End If

    If hitCount = 0 Then
        MsgBox targetName & " さんのデータは見つかりませんでした。", vbInformation
        Exit Sub
    End If

    Call DeleteAllTableRows(tbl)

    For i = 1 To UBound(dataArr, 1)

        If Trim$(CStr(dataArr(i, CONV_COL_担当者))) = targetName Then

            srcKey = BuildSourceKeyFromConv(dataArr, i)

            If Not importedMap.exists(srcKey) Then

                convName = Trim$(CStr(dataArr(i, CONV_COL_入力フォーム表示名)))
                transformState = Trim$(CStr(dataArr(i, CONV_COL_変換状態)))
                originalMaterial = Trim$(CStr(dataArr(i, CONV_COL_材料)))
                kubun = Trim$(CStr(dataArr(i, CONV_COL_区分)))

                If convName <> "" Then
                    Set newRow = tbl.ListRows.Add

                    newRow.Range.Cells(1, colProduct).Value = convName
                    newRow.Range.Cells(1, colQty).Value = dataArr(i, CONV_COL_数量)
                    newRow.Range.Cells(1, colSourceKey).Value = srcKey
                    newRow.Range.Cells(1, col元材料名).Value = originalMaterial

                    If transformState = "未変換" Then

                        If colResult > 0 Then
                            newRow.Range.Cells(1, colResult).Value = "未変換：候補選択"
                        End If

                        cKey = CandidateKey(kubun, originalMaterial)

                        If candidateMap.exists(cKey) Then
                            candidateArr = candidateMap(cKey)
                            newRow.Range.Cells(1, col候補1).Value = candidateArr(0)
                            newRow.Range.Cells(1, col候補2).Value = candidateArr(1)
                            newRow.Range.Cells(1, col候補3).Value = candidateArr(2)
                        Else
                            If colResult > 0 Then
                                newRow.Range.Cells(1, colResult).Value = "未変換：候補なし"
                            End If
                        End If

                    Else
                        If colResult > 0 Then
                            newRow.Range.Cells(1, colResult).ClearContents
                        End If
                    End If
                End If
            End If
        End If
    Next i

    msg = hitCount & "件の抽出が完了しました。"
    If skipCount > 0 Then msg = msg & vbCrLf & "（取込済みスキップ：" & skipCount & "件）"
    MsgBox msg, vbInformation, "抽出完了"

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました：" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub



' =========================================================
' 【機能追加】候補を製品名へ反映する
' =========================================================
Sub 候補を製品名へ反映する()

    Dim wsInput As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim rowRng As Range
    Dim colProduct As Long, colResult As Long
    Dim col候補1 As Long, col候補2 As Long, col候補3 As Long
    Dim col採用候補 As Long, col手入力製品名 As Long, col元材料名 As Long
    Dim choice As String
    Dim newName As String
    Dim changedCount As Long
    Dim skipCount As Long

    On Error GoTo ErrHandler

    If Not SheetExists(INPUT_SHEET) Then
        MsgBox "一括入力フォームが見つかりません。", vbCritical
        Exit Sub
    End If

    Set wsInput = ThisWorkbook.Sheets(INPUT_SHEET)

    On Error Resume Next
    Set tbl = wsInput.ListObjects(TABLE_NAME)
    On Error GoTo ErrHandler

    If tbl Is Nothing Then
        MsgBox "テーブル「" & TABLE_NAME & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    colProduct = GetTableColumnIndex(tbl, "製品名")
    colResult = GetTableColumnIndex(tbl, "処理結果（自動）")
    Call EnsureCandidateColumns(tbl, col候補1, col候補2, col候補3, col採用候補, col手入力製品名, col元材料名)

    If colProduct = 0 Then
        MsgBox "入力フォームに「製品名」列がありません。", vbCritical
        Exit Sub
    End If

    If tbl.DataBodyRange Is Nothing Then
        MsgBox "反映する行がありません。", vbInformation
        Exit Sub
    End If

    For i = 1 To tbl.ListRows.Count
        Set rowRng = tbl.ListRows(i).Range

        choice = Trim$(CStr(rowRng.Cells(1, col採用候補).Value))
        newName = ""

        Select Case choice
            Case "候補1"
                newName = Trim$(CStr(rowRng.Cells(1, col候補1).Value))
            Case "候補2"
                newName = Trim$(CStr(rowRng.Cells(1, col候補2).Value))
            Case "候補3"
                newName = Trim$(CStr(rowRng.Cells(1, col候補3).Value))
            Case "手入力"
                newName = Trim$(CStr(rowRng.Cells(1, col手入力製品名).Value))
            Case "見送り", ""
                skipCount = skipCount + 1
            Case Else
                skipCount = skipCount + 1
        End Select

        If newName <> "" Then
            rowRng.Cells(1, colProduct).Value = newName
            If colResult > 0 Then
                rowRng.Cells(1, colResult).Value = "候補反映済"
            End If
            changedCount = changedCount + 1
        End If
    Next i

    MsgBox "候補反映が完了しました。" & vbCrLf & _
           "反映：" & changedCount & "件" & vbCrLf & _
           "未処理・見送り：" & skipCount & "件", vbInformation

    Exit Sub

ErrHandler:
    MsgBox "候補反映中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub


' =========================================================
' 【機能6】入力フォームのデータを全件クリアする
' ※ ボタンに割り当てて使用
' =========================================================
Sub 入力フォームをクリアする()

    Dim wsInput As Worksheet
    Dim tbl As ListObject

    On Error GoTo ErrHandler

    If Not SheetExists(INPUT_SHEET) Then
        MsgBox "一括入力フォームが見つかりません。", vbCritical
        Exit Sub
    End If

    Set wsInput = ThisWorkbook.Sheets(INPUT_SHEET)

    On Error Resume Next
    Set tbl = wsInput.ListObjects(TABLE_NAME)
    On Error GoTo ErrHandler

    If tbl Is Nothing Then
        MsgBox "テーブル「" & TABLE_NAME & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    If tbl.DataBodyRange Is Nothing Then
        MsgBox "クリアするデータがありません。", vbInformation
        Exit Sub
    End If

    If MsgBox("入力フォームのデータをすべて削除しますか？" & vbCrLf & _
              "※この操作は元に戻せません。", _
              vbExclamation + vbYesNo, "クリア確認") = vbNo Then
        Exit Sub
    End If

    Call DeleteAllTableRows(tbl)

    MsgBox "入力フォームをクリアしました。", vbInformation, "完了"
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました：" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub



'====================================================
' 採用分を変換リストへ追加 UR対応版
'
' 変換候補一覧：
' A列：元の材料名
' B列：区分
' C列：候補順位
' D列：候補製品名
' E列：変換前材料名
' F列：UR
' G列：メーカー
' H列：類似対象
' I列：判定根拠
' J列：UR加点
' K列：類似度スコア
' L列：判定
' M列：採用
'
' 変換リスト：
' A列：変換前（材料名）
' B列：変換後（製品名）
' C列：UR
' D列：メーカー
'====================================================
Sub 採用分を変換リストへ追加_UR対応()

    Dim wsCand As Worksheet
    Dim wsList As Worksheet

    Dim lastCandRow As Long
    Dim lastListRow As Long
    Dim i As Long
    Dim addCount As Long

    Dim srcName As String
    Dim productName As String
    Dim urValue As String
    Dim makerValue As String
    Dim adoptValue As String

    Dim existsAlready As Boolean

    On Error GoTo ErrHandler

    Set wsCand = ThisWorkbook.Worksheets("変換候補一覧")
    Set wsList = ThisWorkbook.Worksheets("変換リスト")

    lastCandRow = wsCand.Cells(wsCand.Rows.Count, "A").End(xlUp).Row

    If lastCandRow < 5 Then
        MsgBox "変換候補一覧にデータがありません。", vbExclamation
        Exit Sub
    End If

    addCount = 0

    Application.ScreenUpdating = False

    For i = 5 To lastCandRow

        adoptValue = Trim(CStr(wsCand.Cells(i, "M").Value))

        If adoptValue = "採用" Then

            srcName = Trim(CStr(wsCand.Cells(i, "A").Value))
            productName = Trim(CStr(wsCand.Cells(i, "D").Value))
            urValue = Trim(CStr(wsCand.Cells(i, "F").Value))
            makerValue = Trim(CStr(wsCand.Cells(i, "G").Value))

            If srcName <> "" And productName <> "" Then

                existsAlready = 変換リストに存在するか_UR対応(wsList, srcName, productName, urValue)

                If existsAlready = False Then

                    lastListRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row + 1

                    wsList.Cells(lastListRow, "A").Value = srcName
                    wsList.Cells(lastListRow, "B").Value = productName
                    wsList.Cells(lastListRow, "C").Value = urValue
                    wsList.Cells(lastListRow, "D").Value = makerValue

                    wsCand.Cells(i, "M").Value = "登録済"

                    addCount = addCount + 1

                Else
                    wsCand.Cells(i, "M").Value = "登録済または重複"
                End If

            End If

        End If

    Next i

    Application.ScreenUpdating = True

    MsgBox addCount & "件を変換リストへ追加しました。", vbInformation

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True

    MsgBox "採用分の登録中にエラーが発生しました。" & vbCrLf & _
           "エラー番号：" & Err.Number & vbCrLf & _
           "内容：" & Err.Description, vbCritical

End Sub


'====================================================
' 変換リスト重複チェック UR対応版
'
' 変換前材料名・変換後製品名・UR の3点で重複判定
'====================================================
Private Function 変換リストに存在するか_UR対応( _
    ByVal ws As Worksheet, _
    ByVal srcName As String, _
    ByVal productName As String, _
    ByVal urValue As String) As Boolean

    Dim lastRow As Long
    Dim i As Long

    Dim aVal As String
    Dim bVal As String
    Dim cVal As String

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    変換リストに存在するか_UR対応 = False

    For i = 2 To lastRow

        aVal = Trim(CStr(ws.Cells(i, "A").Value))
        bVal = Trim(CStr(ws.Cells(i, "B").Value))
        cVal = Trim(CStr(ws.Cells(i, "C").Value))

        If aVal = srcName And bVal = productName And cVal = urValue Then
            変換リストに存在するか_UR対応 = True
            Exit Function
        End If

    Next i

End Function


