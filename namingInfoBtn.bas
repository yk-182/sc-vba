Attribute VB_Name = "namingInfoBtn"
Option Explicit

Sub namingInfoBtnClick()
    'ファイル名取得
    Dim fileName As String: fileName = ThisWorkbook.Name
    'アクティブシート名取得
    Dim activeSheetName As String: activeSheetName = ActiveSheet.Name
    
    ' テキストファイルを開く
    Workbooks.OpenText fileName:="C:\RRDRFT\SOUGI-01.TXT", Origin:=932, StartRow:=1, _
                       DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                       Tab:=True, Semicolon:=False, Comma:=True, Space:=False, Other:=False, _
                       FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), _
                                        Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), _
                                        Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), _
                                        Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), _
                                        Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), _
                                        Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), _
                                        Array(36, 1), Array(37, 1)), TrailingMinusNumbers:=True

    ' コピー範囲を選択
    Rows("2500:2500").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Copy
    ' 貼り付け先のシートにアクセス
    Windows(fileName).Activate
    Sheets("pasted").Select
    ActiveWindow.WindowState = xlNormal
    ' 貼り付ける
    Rows("1:1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ' 元のシートに戻る
    Sheets(activeSheetName).Select
    ActiveWindow.WindowState = xlMaximized
    ' テキストファイルを閉じる
    Windows("SOUGI-01.TXT").Activate
    ActiveWindow.Close
    
    '読込日時更新
    ActiveSheet.Range("C10").value = Format(Now, "mm/dd hh:mm")

End Sub


