Attribute VB_Name = "Module1"
Option Explicit

Sub TransferDataByProductNumber()

    Dim wsSource As Worksheet
    Dim wsFour As Worksheet
    Dim wsTransferData As Worksheet

    Dim lastRowSource As Long
    Dim lastRowFour As Long
    Dim lastRowTransferData As Long

    Dim targetCol As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim productNumber As String
    Dim targetValue1 As Variant
    Dim targetValue2 As Variant

    ' ���݂�Excel�t�@�C�����㏑���ۑ�
    ThisWorkbook.Save

    ' �g�p����V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("�S���ҕ� �����\��\")
    Set wsFour = ThisWorkbook.Sheets("��s��(�l�H�i)")
    Set wsTransferData = ThisWorkbook.Sheets("��s��(�d���i)")

    ' ===== �@ �]�ʑO�Ɋ����f�[�^���N���A =====
    Dim pageRangesFour As Variant
    Dim pageRangesProc As Variant

    ' �l�H�i���̃N���A�͈�
    pageRangesFour = Array( _
        "N18:U71", _
        "N91:U144", _
        "N164:U217", _
        "N237:U290", _
        "N310:U363" _
    )

    ' �d���i���̃N���A�͈�
    pageRangesProc = Array( _
        "N15:U24", _
        "N27:U68", _
        "N87:U142", _
        "N159:U214", _
        "N224:U287", _
        "N297:U357" _
    )

    ' �d���i�� pageRangesProc ����ɃN���A����
    ' �l�H�i�� pageRangesFour �͔͈͐������Ȃ����߁A���݂���͈͂����N���A����
    For k = LBound(pageRangesProc) To UBound(pageRangesProc)

        If k <= UBound(pageRangesFour) Then
            wsFour.Range(pageRangesFour(k)).ClearContents
        End If

        wsTransferData.Range(pageRangesProc(k)).ClearContents

    Next k

    ' ===== ���f�[�^�s���m�F =====
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastRowFour = wsFour.Cells(wsFour.Rows.Count, 2).End(xlUp).Row
    lastRowTransferData = wsTransferData.Cells(wsTransferData.Rows.Count, 2).End(xlUp).Row

    ' �u�S���ҕ� �����\��\�v�́u���v�v�̗������
    targetCol = 0

    For i = 1 To wsSource.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column

        If wsSource.Cells(2, i).Value = "���v" Then
            targetCol = i
            Exit For
        End If

    Next i

    If targetCol = 0 Then
        MsgBox "�u���v�v�ƈ�v����񂪌�����܂���ł����B", vbExclamation
        Exit Sub
    End If

    ' �]�ʑΏۗ�
    Dim sourceCol1 As Long
    Dim sourceCol2 As Long

    sourceCol1 = targetCol
    sourceCol2 = targetCol + 1

    ' �]�ʐ��
    Dim destCol1 As Long
    Dim destCol2 As Long

    destCol1 = 14 ' N��
    destCol2 = 18 ' R��

    ' ===== �A �f�[�^�]�� =====
    For i = 3 To lastRowSource

        productNumber = Trim(wsSource.Cells(i, 1).Value) ' ���i�ԍ� A��

        ' ���i�ԍ�����̏ꍇ�̓X�L�b�v
        If productNumber = "" Then
            GoTo SkipRow
        End If

        targetValue1 = wsSource.Cells(i, sourceCol1).Value
        targetValue2 = wsSource.Cells(i, sourceCol2).Value

        ' �l�H�i
        For j = 2 To lastRowFour

            If wsFour.Cells(j, 2).Value = productNumber Then
                wsFour.Cells(j, destCol1).Value = targetValue1
                wsFour.Cells(j, destCol2).Value = targetValue2
                Exit For
            End If

        Next j

        ' �d���i
        For j = 2 To lastRowTransferData

            If wsTransferData.Cells(j, 2).Value = productNumber Then
                wsTransferData.Cells(j, destCol1).Value = targetValue1
                wsTransferData.Cells(j, destCol2).Value = targetValue2
                Exit For
            End If

        Next j

SkipRow:
    Next i

    ' �N����]��
    wsFour.Range("B2").Value = wsSource.Range("M1").Value

    ' ===== �B N1�Z���󔒎��� Left(..., Len(...) - 1) �G���[�h�~ =====
    If Len(wsSource.Range("N1").Value) > 0 Then
        wsFour.Range("G2").Value = Left(wsSource.Range("N1").Value, Len(wsSource.Range("N1").Value) - 1)
    Else
        wsFour.Range("G2").Value = ""
    End If

    wsTransferData.Range("C2").Value = wsFour.Range("G2").Value
    wsFour.Range("B4").Value = Format(Date, "yyyy/m/d")
    wsTransferData.Range("B4").Value = Format(Date, "yyyy/m/d")

    MsgBox "�f�[�^�̓]�ʂ��������܂����B", vbInformation

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

    ' �g�p����V�[�g��ݒ�
    Set wsFour = ThisWorkbook.Sheets("��s��(�l�H�i)")
    Set wsProcurement = ThisWorkbook.Sheets("��s��(�d���i)")
    Set wsSettings = ThisWorkbook.Sheets("�v���O�����ݒ�")

    ' ���g���ۑ�����Ă���t�H���_�̃p�X���擾���āu�v���O�����ݒ�v�̃Z��A1�ɏo��
    currentFolderPath = ThisWorkbook.Path
    wsSettings.Range("A1").Value = currentFolderPath

    ' �ۑ���̃t�H���_�p�X���擾
    saveFolder = wsSettings.Range("A2").Value

    ' �ۑ���t�H���_�����݂��Ȃ��ꍇ�͍쐬
    If Dir(saveFolder, vbDirectory) = "" Then
        MkDir saveFolder
    End If

    ' �t�@�C�������擾
    fileNameFour = wsSettings.Range("A4").Value
    fileNameProcurement = wsSettings.Range("A5").Value

    ' �u��s��(�l�H�i)�v�V�[�g��V�����t�@�C���Ƃ��ĕۑ�
    wsFour.Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs Filename:=saveFolder & "\" & fileNameFour, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False

    ' �u��s��(�d���i)�v�V�[�g��V�����t�@�C���Ƃ��ĕۑ�
    wsProcurement.Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs Filename:=saveFolder & "\" & fileNameProcurement, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False

    Application.ScreenUpdating = True

    MsgBox "�t�@�C���̕ۑ����������܂����B", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical

End Sub

Sub ClearData()

End Sub


