Attribute VB_Name = "PrintBtn"
Option Explicit

'�I�����ꂽNo.�̒l����\�t�D�V�[�g�����
Sub PrintOut()
    Dim cellValues As Variant
    Dim i As Long
    Dim currentSheetName As String
    currentSheetName = ActiveSheet.Name
    
    cellValues = GetSelectedCellsValue()

    If Not IsEmpty(cellValues) Then
        PrintSheetValues currentSheetName, cellValues
    End If
End Sub

'�I�����ꂽ�Z���l�̎擾
Function GetSelectedCellsValue() As Variant
    Dim selectedRange As Range
    'Range�O�̂��̂��I������Ă����ꍇ�͏����𒆎~
    If TypeName(Selection) <> "Range" Then
        MsgBox "�uNo.�v���ڂ̃Z����I�����Ă��������B", vbCritical, "�Z���̌��o�s��"
        Exit Function
    End If
    Set selectedRange = Selection                '���ݑI������Ă���Z�����擾

    
    Dim targetRange As Range
    Set targetRange = Range("B5:B304")           '����͈̔͂̃Z��
    
    Dim cell As Range
    Dim isOutOfRange As Boolean
    isOutOfRange = False
    
    'NO.��ȊO�̃Z�����I������Ă����ꍇisOutOfRange�t���O��TRUE
    For Each cell In selectedRange
        If Intersect(cell, targetRange) Is Nothing Then
            isOutOfRange = True
            Exit For
        End If
    Next cell
    If isOutOfRange Then
        MsgBox "�uNo.�v���ڂ̃Z����I�����Ă��������B", vbCritical, "�͈͊O�̃Z�������o"
        Exit Function
    End If
    
    ' �I�����ꂽ�Z���̌���20��葽���ꍇ�A�����𒆎~����
    If selectedRange.Cells.Count > 20 Then
        MsgBox "��x�ɑI���ł���Z���̐���20�܂łł��B", vbCritical, "������𒴉�"
        Exit Function
    End If
    
    Dim cellValues() As String
    ReDim cellValues(1 To selectedRange.Cells.Count)
    Dim i As Long
    i = 1
    For Each cell In selectedRange
        cellValues(i) = cell.value
        i = i + 1
    Next cell
    
    GetSelectedCellsValue = cellValues
End Function

'�������
Sub PrintSheetValues(sheetName As String, cellValues As Variant)
    Dim printSheet As Worksheet
    Dim i As Long

    Set printSheet = ThisWorkbook.Worksheets("�\�t�D")
    printSheet.Range("B1").value = sheetName
    
    ' �Z���̒l��E1�Z���ɐݒ肵�A�V�[�g�����
    For i = LBound(cellValues) To UBound(cellValues)
        printSheet.Range("E1").value = cellValues(i)
        ' �Z���ɒl���ݒ肳���̂�҂��߂�2�b�̃E�F�C�g���� TODO �s�v�H�{�Ԋ��ŗv����
        'Application.Wait (Now + TimeValue("0:00:02"))
        printSheet.PrintOut
        '�v���s���[��\�����Ď蓮�ň������ꍇ
        'printSheet.PrintPreview
    Next i
    'MsgBox "����������I�����܂����B"
End Sub

