Attribute VB_Name = "PrintBtn"
Option Explicit

'選択されたNo.の値から貼付札シートを印刷
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

'選択されたセル値の取得
Function GetSelectedCellsValue() As Variant
    Dim selectedRange As Range
    'Range外のものが選択されていた場合は処理を中止
    If TypeName(Selection) <> "Range" Then
        MsgBox "「No.」項目のセルを選択してください。", vbCritical, "セルの検出不可"
        Exit Function
    End If
    Set selectedRange = Selection                '現在選択されているセルを取得

    
    Dim targetRange As Range
    Set targetRange = Range("B5:B304")           '特定の範囲のセル
    
    Dim cell As Range
    Dim isOutOfRange As Boolean
    isOutOfRange = False
    
    'NO.列以外のセルが選択されていた場合isOutOfRangeフラグをTRUE
    For Each cell In selectedRange
        If Intersect(cell, targetRange) Is Nothing Then
            isOutOfRange = True
            Exit For
        End If
    Next cell
    If isOutOfRange Then
        MsgBox "「No.」項目のセルを選択してください。", vbCritical, "範囲外のセルを検出"
        Exit Function
    End If
    
    ' 選択されたセルの個数が20より多い場合、処理を中止する
    If selectedRange.Cells.Count > 20 Then
        MsgBox "一度に選択できるセルの数は20までです。", vbCritical, "上限数を超過"
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

'印刷処理
Sub PrintSheetValues(sheetName As String, cellValues As Variant)
    Dim printSheet As Worksheet
    Dim i As Long

    Set printSheet = ThisWorkbook.Worksheets("貼付札")
    printSheet.Range("B1").value = sheetName
    
    ' セルの値をE1セルに設定し、シートを印刷
    For i = LBound(cellValues) To UBound(cellValues)
        printSheet.Range("E1").value = cellValues(i)
        ' セルに値が設定されるのを待つための2秒のウェイト時間 TODO 不要？本番環境で要検証
        'Application.Wait (Now + TimeValue("0:00:02"))
        printSheet.PrintOut
        'プレピューを表示して手動で印刷する場合
        'printSheet.PrintPreview
    Next i
    'MsgBox "印刷処理が終了しました。"
End Sub

