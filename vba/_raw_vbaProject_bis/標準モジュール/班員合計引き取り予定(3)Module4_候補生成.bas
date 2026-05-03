- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'==============================================================
' Module4: 変換候補生成
'
' Python in Excel の PY()セル（変換候補一覧!A5）を代替する。
' SequenceMatcher.ratio() を bigram 重複率で近似実装。
'
' 対応: python/班員合計引き取り予定(3)_python_script_1.py
' 呼び出し: 「変換候補を生成」ボタン → 変換候補一覧を生成する()
'==============================================================
Option Explicit

' ---------- シート名定数 ----------
Private Const CAND_SHEET     As String = "変換候補一覧"
Private Const CONV_DATA_SHEET As String = "★変換済みデータ"
Private Const SETTING_SHEET  As String = "プログラム設定"

' ---------- プログラム設定 セル位置 ----------
Private Const USE_PY_FLAG_ROW As Long = 3
Private Const USE_PY_FLAG_COL As Long = 1

' ---------- 変換候補一覧 行番号 ----------
Private Const HEADER_ROW     As Long = 4
Private Const DATA_START_ROW As Long = 5

' ---------- ★変換済みデータ 列番号（Module2 CONV_COL_* と同値）----------
Private Const CD_区分         As Long = 5
Private Const CD_材料         As Long = 6
Private Const CD_変換状態     As Long = 19
Private Const CD_TOTAL_COLS   As Long = 19

' ---------- 変換リストテーブル名 ----------
Private Const CONV_TABLE_NAME As String = "テーブル1"


'==============================================================
' 公開: 変換候補一覧を生成する
'
' 1. USE_PYTHON_PATH フラグ確認
' 2. Power Query 更新
' 3. ★変換済みデータ から未変換行抽出（区分・材料のユニークペア）
' 4. 変換リスト（テーブル1）読み込み
' 5. bigram 類似度スコア計算 + URボーナス
' 6. 変換候補一覧シートへ書き込み
'==============================================================
Public Sub 変換候補一覧を生成する()

    Dim wsConv  As Worksheet
    Dim wsCand  As Worksheet
    Dim wsSet   As Worksheet

    Dim convDataArr  As Variant
    Dim convListArr  As Variant
    Dim lastConvRow  As Long
    Dim i            As Long

    ' ---------- 0. USE_PYTHON_PATH フラグ確認 ----------
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets(SETTING_SHEET)
    On Error GoTo 0

    If Not wsSet Is Nothing Then
        Dim flagVal As Variant
        flagVal = wsSet.Cells(USE_PY_FLAG_ROW, USE_PY_FLAG_COL).Value
        If UCase$(Trim$(CStr(flagVal))) = "TRUE" Then
            MsgBox "プログラム設定（A3）の USE_PYTHON_PATH が TRUE のため" & vbCrLf & _
                   "このマクロはスキップします。" & vbCrLf & vbCrLf & _
                   "VBA経路に切り替えるには A3 を FALSE または空白にしてください。", _
                   vbInformation, "Python経路が有効"
            Exit Sub
        End If
    End If

    ' ---------- 1. Power Query 更新 ----------
    Application.StatusBar = "Power Query を更新中..."
    On Error Resume Next
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.OLEDBConnection.BackgroundQuery = False
        conn.Refresh
    Next conn
    On Error GoTo 0
    Application.StatusBar = False

    ' ---------- 2. シート存在確認 ----------
    If Not SheetExistsM4(CONV_DATA_SHEET) Then
        MsgBox "「" & CONV_DATA_SHEET & "」シートが見つかりません。" & vbCrLf & vbCrLf & _
               "【対処】Power Queryの「★変換済みデータ」クエリを" & vbCrLf & _
               "「クエリと接続」→ 右クリック → 読み込み先 →" & vbCrLf & _
               "「テーブル」として新規シートに出力してください。", vbCritical, "シートなし"
        Exit Sub
    End If

    Set wsConv = ThisWorkbook.Sheets(CONV_DATA_SHEET)
    lastConvRow = wsConv.Cells(wsConv.Rows.Count, 1).End(xlUp).Row

    If lastConvRow < 2 Then
        MsgBox "★変換済みデータにデータがありません。" & vbCrLf & _
               "先にマクロでデータを取り込み、PQを更新してください。", vbInformation
        Exit Sub
    End If

    ' ---------- 3. 列構成チェック ----------
    If CStr(wsConv.Cells(1, CD_変換状態).Value) <> "変換状態" Then
        MsgBox "★変換済みデータの列構成が想定と異なります。" & vbCrLf & _
               "S列（19列目）が「変換状態」であることを確認してください。" & vbCrLf & vbCrLf & _
               "現在の S1 の値: " & CStr(wsConv.Cells(1, CD_変換状態).Value), vbCritical, "列構成エラー"
        Exit Sub
    End If

    ' ---------- 4. ★変換済みデータ配列化 ----------
    convDataArr = wsConv.Range(wsConv.Cells(2, 1), wsConv.Cells(lastConvRow, CD_TOTAL_COLS)).Value

    ' ---------- 5. 未変換の (区分, 材料) ユニークペアを収集 ----------
    Dim matDict As Object
    Set matDict = CreateObject("Scripting.Dictionary")
    matDict.CompareMode = vbTextCompare

    Dim 区分v As String, 材料v As String, mKey As String

    For i = 1 To UBound(convDataArr, 1)
        If Trim$(CStr(convDataArr(i, CD_変換状態))) = "未変換" Then
            区分v = Trim$(CStr(convDataArr(i, CD_区分)))
            材料v = Trim$(CStr(convDataArr(i, CD_材料)))
            If 材料v <> "" Then
                mKey = 区分v & Chr(1) & 材料v
                If Not matDict.exists(mKey) Then
                    matDict.Add mKey, Array(区分v, 材料v)
                End If
            End If
        End If
    Next i

    If matDict.Count = 0 Then
        MsgBox "未変換の材料が見つかりませんでした。" & vbCrLf & _
               "PQを更新してから再実行してください。", vbInformation
        Exit Sub
    End If

    ' ---------- 6. 変換リスト（テーブル1）読み込み ----------
    convListArr = 変換リスト配列取得_M4()

    If IsEmpty(convListArr) Then
        MsgBox "変換リスト（テーブル1）が読み込めませんでした。" & vbCrLf & _
               "「変換リスト」シートに「テーブル1」テーブルが存在するか確認してください。", _
               vbCritical, "変換リスト読み込みエラー"
        Exit Sub
    End If

    ' ---------- 7. 候補シートへ書き込み ----------
    If Not SheetExistsM4(CAND_SHEET) Then
        MsgBox "「" & CAND_SHEET & "」シートが見つかりません。", vbCritical, "シートなし"
        Exit Sub
    End If

    Set wsCand = ThisWorkbook.Sheets(CAND_SHEET)
    Application.ScreenUpdating = False

    ' データ行（A〜H列のみ）をクリア（I列以降のユーザー追記列は保護）
    Dim lastCandRow As Long
    lastCandRow = wsCand.Cells(wsCand.Rows.Count, "A").End(xlUp).Row
    If lastCandRow >= DATA_START_ROW Then
        wsCand.Range("A" & DATA_START_ROW & ":H" & lastCandRow).ClearContents
    End If

    ' ヘッダ書き込み
    With wsCand.Rows(HEADER_ROW)
        .Cells(1, 1).Value = "元の材料名"
        .Cells(1, 2).Value = "区分"
        .Cells(1, 3).Value = "候補1"
        .Cells(1, 4).Value = "スコア1"
        .Cells(1, 5).Value = "候補2"
        .Cells(1, 6).Value = "スコア2"
        .Cells(1, 7).Value = "候補3"
        .Cells(1, 8).Value = "スコア3"
    End With

    ' 各未変換材料について上位3候補を計算して書き込む
    Dim writeRow As Long
    writeRow = DATA_START_ROW

    Dim mKeyVar As Variant
    Dim matEntry As Variant
    Dim top3 As Variant
    Dim ci As Long, colBase As Long

    For Each mKeyVar In matDict.Keys
        matEntry = matDict(mKeyVar)
        top3 = Top3Candidates_M4(CStr(matEntry(1)), CStr(matEntry(0)), convListArr)

        wsCand.Cells(writeRow, 1).Value = CStr(matEntry(1))  ' 元の材料名
        wsCand.Cells(writeRow, 2).Value = CStr(matEntry(0))  ' 区分

        For ci = 0 To 2
            colBase = 3 + ci * 2
            wsCand.Cells(writeRow, colBase).Value = top3(ci, 0)
            If CDbl(top3(ci, 1)) > 0 Then
                wsCand.Cells(writeRow, colBase + 1).Value = top3(ci, 1)
            End If
        Next ci

        writeRow = writeRow + 1
    Next mKeyVar

    Application.ScreenUpdating = True

    MsgBox "変換候補一覧を生成しました（" & matDict.Count & " 件）。" & vbCrLf & _
           "「採用候補」列を確認し、採用するものを選択してください。", _
           vbInformation, "完了"

End Sub


'==============================================================
' 内部: 変換リスト（テーブル1）を2次元配列で返す
' 列: 1=変換前（材料名）, 2=変換後（製品名）, 3=UR, 4=メーカー
'==============================================================
Private Function 変換リスト配列取得_M4() As Variant

    Dim ws  As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Dim t As ListObject
        For Each t In ws.ListObjects
            If t.Name = CONV_TABLE_NAME Then
                Set tbl = t
                Exit For
            End If
        Next t
        If Not tbl Is Nothing Then Exit For
    Next ws
    On Error GoTo 0

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    変換リスト配列取得_M4 = tbl.DataBodyRange.Value

End Function


'==============================================================
' 内部: 上位3候補を取得
'
' 引数:
'   材料    - 未変換の材料名
'   区分    - "UR" or "通常"
'   convArr - 変換リスト配列 (変換前, 変換後, UR, メーカー)
'
' 戻り値: Variant(0 To 2, 0 To 1)
'   (候補インデックス, 0) = 製品名
'   (候補インデックス, 1) = スコア（0なら空白扱い）
'==============================================================
Private Function Top3Candidates_M4(ByVal 材料 As String, ByVal 区分 As String, _
                                    ByVal convArr As Variant) As Variant

    Dim isUrCase As Boolean
    isUrCase = (UCase$(Trim$(区分)) = "UR")

    Dim n As Long
    n = UBound(convArr, 1)

    Dim scores() As Double
    Dim names()  As String
    ReDim scores(1 To n)
    ReDim names(1 To n)

    Dim i As Long
    Dim 変換前 As String, 変換後 As String, urFlag As String
    Dim s1 As Double, s2 As Double, baseScore As Double, finalScore As Double

    For i = 1 To n
        変換前 = Trim$(CStr(convArr(i, 1)))
        変換後 = Trim$(CStr(convArr(i, 2)))
        urFlag = Trim$(CStr(convArr(i, 3)))

        s1 = SimilarityScore_M4(材料, 変換前)
        s2 = SimilarityScore_M4(材料, 変換後)
        baseScore = IIf(s1 > s2, s1, s2)

        finalScore = baseScore + IIf(isUrCase And IsUrProduct_M4(urFlag), 10, 0)
        If finalScore > 100 Then finalScore = 100

        scores(i) = finalScore
        names(i) = 変換後
    Next i

    ' 上位3を選択ソート（3回）
    Dim used()      As Boolean
    ReDim used(1 To n)

    Dim result(0 To 2, 0 To 1) As Variant
    Dim r As Long, bestIdx As Long
    Dim bestScore As Double

    For r = 0 To 2
        bestIdx = 0
        bestScore = -1
        For i = 1 To n
            If Not used(i) Then
                If scores(i) > bestScore Then
                    bestScore = scores(i)
                    bestIdx = i
                End If
            End If
        Next i
        If bestIdx > 0 And bestScore > 0 Then
            result(r, 0) = names(bestIdx)
            result(r, 1) = Round(bestScore, 1)
            used(bestIdx) = True
        Else
            result(r, 0) = ""
            result(r, 1) = 0
        End If
    Next r

    Top3Candidates_M4 = result

End Function


'==============================================================
' 内部: 類似度スコア（0〜100）
'
' Python SequenceMatcher(None, a, b).ratio() * 100 の近似実装。
' bigram重複率（Dice係数）でスコアを算出する。
'
' 完全一致 → 100、部分一致 → 92、それ以外 → bigram率×100
'==============================================================
Private Function SimilarityScore_M4(ByVal a As String, ByVal b As String) As Double

    Dim na As String, nb As String
    na = NormalizeForMatch_M4(a)
    nb = NormalizeForMatch_M4(b)

    If na = "" Or nb = "" Then SimilarityScore_M4 = 0 : Exit Function
    If na = nb Then SimilarityScore_M4 = 100 : Exit Function
    If InStr(1, na, nb, vbBinaryCompare) > 0 Or _
       InStr(1, nb, na, vbBinaryCompare) > 0 Then
        SimilarityScore_M4 = 92
        Exit Function
    End If

    Dim la As Long : la = Len(na)
    Dim lb As Long : lb = Len(nb)

    If la < 2 Or lb < 2 Then
        SimilarityScore_M4 = 0
        Exit Function
    End If

    ' Bigram（2文字n-gram）の重複カウント（Dice係数）
    Dim totalBigrams As Long
    totalBigrams = (la - 1) + (lb - 1)

    Dim used() As Boolean
    ReDim used(1 To lb - 1)
    Dim matchCount As Long : matchCount = 0

    Dim i As Long, j As Long
    Dim bg As String

    For i = 1 To la - 1
        bg = Mid$(na, i, 2)
        For j = 1 To lb - 1
            If Not used(j) Then
                If Mid$(nb, j, 2) = bg Then
                    matchCount = matchCount + 1
                    used(j) = True
                    Exit For
                End If
            End If
        Next j
    Next i

    SimilarityScore_M4 = (2# * matchCount / totalBigrams) * 100

End Function


'==============================================================
' 内部: 文字列正規化
'
' Python 正規化() に対応。
' NFKC相当（vbNarrow）+ 大文字化 + 記号除去 + 単位変換。
'==============================================================
Private Function NormalizeForMatch_M4(ByVal s As Variant) As String

    If IsNull(s) Or IsEmpty(s) Then NormalizeForMatch_M4 = "" : Exit Function

    Dim r As String
    r = CStr(s)
    If Trim$(r) = "" Then NormalizeForMatch_M4 = "" : Exit Function

    On Error Resume Next
    r = StrConv(r, vbNarrow)
    On Error GoTo 0

    r = UCase$(r)

    r = Replace(r, " ",     "")
    r = Replace(r, "　",    "")
    r = Replace(r, Chr(160), "")
    r = Replace(r, vbTab,   "")

    r = Replace(r, "－", "-")
    r = Replace(r, "―", "-")
    r = Replace(r, "‐", "-")
    r = Replace(r, "／", "/")
    r = Replace(r, "（", "(")
    r = Replace(r, "）", ")")
    r = Replace(r, "［", "[")
    r = Replace(r, "］", "]")
    r = Replace(r, "｛", "{")
    r = Replace(r, "｝", "}")
    r = Replace(r, ".",  "")
    r = Replace(r, "．", "")
    r = Replace(r, "・", "")
    r = Replace(r, "【", "")
    r = Replace(r, "】", "")
    r = Replace(r, "「", "")
    r = Replace(r, "」", "")
    r = Replace(r, "『", "")
    r = Replace(r, "』", "")
    r = Replace(r, ",",  "")
    r = Replace(r, "，", "")
    r = Replace(r, "_",  "")
    r = Replace(r, "＿", "")

    ' 単位変換（Python 正規化() と同等）
    r = Replace(r, "㎡",   "M2")
    r = Replace(r, "M²",  "M2")
    r = Replace(r, "㎜",   "MM")
    r = Replace(r, "ミリ", "MM")
    r = Replace(r, "㎖",   "ML")
    r = Replace(r, "ML",   "ML")

    NormalizeForMatch_M4 = Trim$(r)

End Function


'==============================================================
' 内部: UR製品かどうかの判定
'
' Python UR品か() に対応。
'==============================================================
Private Function IsUrProduct_M4(ByVal v As Variant) As Boolean

    If IsNull(v) Or IsEmpty(v) Then IsUrProduct_M4 = False : Exit Function

    Dim s As String
    On Error Resume Next
    s = UCase$(Trim$(StrConv(CStr(v), vbNarrow)))
    On Error GoTo 0

    Select Case s
        Case "UR", "○", "〇", "1", "TRUE", "YES", "対象"
            IsUrProduct_M4 = True
        Case Else
            IsUrProduct_M4 = False
    End Select

End Function


'==============================================================
' 内部: シート存在確認（このモジュール専用、Module2と独立）
'==============================================================
Private Function SheetExistsM4(ByVal sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sName)
    SheetExistsM4 = Not ws Is Nothing
    On Error GoTo 0
End Function
