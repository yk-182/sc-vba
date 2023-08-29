VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegisterForm 
   Caption         =   "登録フォーム"
   ClientHeight    =   7546
   ClientLeft      =   110
   ClientTop       =   451
   ClientWidth     =   12760
   OleObjectBlob   =   "RegisterForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'行数修正値
Const DROP_ROWS As Integer = 4
'骨器名入れ有無
Dim namingFlg As Boolean
'届け先配列変数
Dim addressWest As Variant
Dim addressCenter As Variant
Dim addressCentralPref As Variant
Dim addressShonan As Variant
Dim addressEast As Variant
Dim addressWestTokyo As Variant
'データ追加時の各列数
Dim colNumberCeremonyCode As Long
Dim colNumberName As Long
Dim colNumberStatus As Long
Dim colNumberRcptDate As Long
Dim colNumberAddress As Long
Dim colNumberNokanDate As Long
Dim colNumberNaming As Long
Dim colNumberTsuyaDate As Long
Dim colNumberKokubetsushikiDate As Long
Dim colNumberNotes As Long
Dim colNumberItemCode1 As Long
Dim colNumberItemCode2 As Long
Dim colNumberItemCode3 As Long
Dim colNumberItemCode4 As Long
Dim colNumberItemCode5 As Long
Dim colNumberItemCode6 As Long
Dim colNumberItemCode7 As Long
Dim colNumberItemCode8 As Long
Dim colNumberItemCode9 As Long
Dim colNumberItemCode10 As Long
Dim colNumberItemCode11 As Long
Dim colNumberItemCode12 As Long
Dim colNumberItemCode13 As Long
Dim colNumberItemCode14 As Long
Dim colNumberItemCode15 As Long
Dim colNumberItemQty1 As Long
Dim colNumberItemQty2 As Long
Dim colNumberItemQty3 As Long
Dim colNumberItemQty4 As Long
Dim colNumberItemQty5 As Long
Dim colNumberItemQty6 As Long
Dim colNumberItemQty7 As Long
Dim colNumberItemQty8 As Long
Dim colNumberItemQty9 As Long
Dim colNumberItemQty10 As Long
Dim colNumberItemQty11 As Long
Dim colNumberItemQty12 As Long
Dim colNumberItemQty13 As Long
Dim colNumberItemQty14 As Long
Dim colNumberItemQty15 As Long

' フォーム初期化時に呼び出し
Private Sub UserForm_Initialize()
    Dim areaData As Variant
    areaData = Array("西", "中央", "県央", "南", "東", "西東京")
    With DivisionCbo
        .List() = areaData
        .Style = fmStyleDropDownList
    End With
    '初期値設定
    Call SetInitialValue
End Sub

'初期値入力
Public Sub SetInitialValue()
    '通常オプションボタンをON
    TypeNormalOpt.value = True
    '年の初期値を設定
    YearTxt1.value = Year(Date)
    YearTxt2.value = Year(Date)
    YearTxt3.value = Year(Date)
    'カラム列数を設定
    colNumberCeremonyCode = 3
    colNumberName = 4
    colNumberStatus = 6
    colNumberRcptDate = 8
    colNumberAddress = 9
    colNumberNokanDate = 10
    colNumberNaming = 11
    colNumberTsuyaDate = 12
    colNumberKokubetsushikiDate = 13
    colNumberNotes = 14
    colNumberItemCode1 = 15
    colNumberItemCode2 = 18
    colNumberItemCode3 = 21
    colNumberItemCode4 = 24
    colNumberItemCode5 = 27
    colNumberItemCode6 = 30
    colNumberItemCode7 = 33
    colNumberItemCode8 = 36
    colNumberItemCode9 = 39
    colNumberItemCode10 = 42
    colNumberItemCode11 = 45
    colNumberItemCode12 = 48
    colNumberItemCode13 = 51
    colNumberItemCode14 = 54
    colNumberItemCode15 = 57
    colNumberItemQty1 = 17
    colNumberItemQty2 = 20
    colNumberItemQty3 = 23
    colNumberItemQty4 = 26
    colNumberItemQty5 = 29
    colNumberItemQty6 = 32
    colNumberItemQty7 = 35
    colNumberItemQty8 = 38
    colNumberItemQty9 = 41
    colNumberItemQty10 = 44
    colNumberItemQty11 = 47
    colNumberItemQty12 = 50
    colNumberItemQty13 = 53
    colNumberItemQty14 = 56
    colNumberItemQty15 = 59
End Sub

'登録ボタンクリック時の処理
Private Sub RegisterBtn_Click()
    '登録可否フラグ
    Dim isEnabled As Boolean: isEnabled = True
    If TypeNormalOpt.value = True Then
        '通常オプションボタン選択時のバリデーション
        Call checkNormalTypeValues(isEnabled)
        If isEnabled = False Then
            Exit Sub
        End If
    Else
        '規定在庫オプションボタン選択時のバリデーション
        Call checkStockTypeValues(isEnabled)
        If isEnabled = False Then
            Exit Sub
        End If
    End If
    
    '確認メッセージ表示
    Dim Msg As String, title As String, res As Integer
    Msg = "登録処理を行います。よろしいですか？"
    title = "新規登録確認"
    res = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, title)
    '「いいえ」の場合は処理中止
    If res = vbNo Then Exit Sub
    '登録処理
    Call RegistInfo
    
End Sub

'通常タイプのバリデーション
Public Sub checkNormalTypeValues(isEnabled)
    If CeremonyCodeTxt.value = "" Then
        MsgBox "施行コードを入力してください。", vbCritical
        CeremonyCodeTxt.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If NameTxt.value = "" Then
        MsgBox "葬家名を入力してください。", vbCritical
        NameTxt.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If DivisionCbo.value = "" Then
        MsgBox "事業部を選択してください。", vbCritical
        DivisionCbo.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If AddressCbo.value = "" Then
        MsgBox "届け先を選択してください。", vbCritical
        AddressCbo.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If YearTxt3.value = "" Or DateTxt3.value = "" Or TimesTxt3.value = "" Then
        MsgBox "告別式日時を入力してください。", vbCritical
        DateTxt3.SetFocus
        isEnabled = False
        Exit Sub
    End If
    '施行コードが入力されている場合は重複チェック
    If CeremonyCodeTxt.value <> "" Then
        '事業部
        Dim division As String: division = DivisionCbo.value
        '施行コード
        Dim ceremonyCode As String: ceremonyCode = CeremonyCodeTxt.value
        '施行コード列の最終行を取得
        Dim lastRow As Long: lastRow = TableLastRow()
        If WorksheetFunction.CountIf(Range(Worksheets(division).Cells(5, 3), Worksheets(division).Cells(lastRow + 4, 3)), ceremonyCode) >= 1 Then
            MsgBox "同一の施行コードが既に登録されています。", vbCritical
            CeremonyCodeTxt.SetFocus
            isEnabled = False
            Exit Sub
        End If
    End If
    '商品数量の入力チェック
    If checkItemQtyValue() = False Then
        MsgBox "商品数量を入力してください。", vbCritical
        isEnabled = False
        Exit Sub
    End If
End Sub

'規定在庫タイプのバリデーション
Public Sub checkStockTypeValues(isEnabled)
    If NameTxt.value = "" Then
        MsgBox "葬家名を入力してください。", vbCritical
        NameTxt.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If DivisionCbo.value = "" Then
        MsgBox "事業部を選択してください。", vbCritical
        DivisionCbo.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If AddressCbo.value = "" Then
        MsgBox "届け先を選択してください。", vbCritical
        AddressCbo.SetFocus
        isEnabled = False
        Exit Sub
    End If
    '商品数量の入力チェック
    If checkItemQtyValue() = False Then
        MsgBox "商品数量を入力してください。", vbCritical
        isEnabled = False
        Exit Sub
    End If
End Sub

'テーブルのデータ最終行を取得する
Function TableLastRow() As Long
    '対象テーブル
    Dim targetTable As ListObject
    Set targetTable = Worksheets(DivisionCbo.value).ListObjects(1)
    Dim codeColumn As Long: codeColumn = targetTable.ListColumns("施行コード").index
    'テーブルの下から順にデータの入っている行を探す
    Dim i As Long
    With targetTable.DataBodyRange
        For i = .Rows.Count To 1 Step -1
            If .Cells(i, codeColumn).value <> "" Then
                TableLastRow = i
                Exit Function
            End If
        Next
    End With
End Function

'商品数量の入力チェック
Function checkItemQtyValue() As Boolean
    Dim i As Integer
    Dim itemCode As Control, itemQty As Control
    checkItemQtyValue = True
    For i = 1 To 15
        Set itemCode = Me.Controls("itemCode" & i)
        Set itemQty = Me.Controls("itemQty" & i)

        If itemCode.value <> "" And itemQty.value = "" Then
            checkItemQtyValue = False
            Exit For
        End If
    Next i
End Function


'名入れ情報取得処理
Private Sub NamingBtn_Click()
    
    Dim serchValue As String                     '検索値
    Dim serchRange As Range                      '検索範囲
    Dim kotsukiName As String                    '骨器名称
    Dim namingValue As Long                      '名入れ判定値
    '入力チェック
    If CeremonyCodeTxt.value = "" Or DivisionCbo.value = "" Then
        MsgBox "施行コードと事業部を入力してください。", vbCritical
        Exit Sub
    End If
    '検索値を設定
    serchValue = DivisionCbo.value & Right(CeremonyCodeTxt.value, 4)
    '検索範囲を設定
    Set serchRange = Worksheets("list").Range("F2:H900")

    On Error Resume Next                         'エラー無視に設定
    '検索値からlistシートの一致する骨器名と名入れ判別値を取得
    kotsukiName = WorksheetFunction.VLookup(serchValue, serchRange, 2, False)
    If Err.Number > 0 Then
        MsgBox "名入れ情報を取得できません。先にデータを取り込んでください。", vbCritical
        On Error GoTo 0                          'エラー発生に戻す
        Exit Sub
    End If
    namingValue = WorksheetFunction.VLookup(serchValue, serchRange, 3, False)
    On Error GoTo 0                              'エラー発生に戻す
    
    If namingValue < 0 Then
        '名入れ有り
        NamingExistOpt.value = True
        MsgBox "【名入れ有り】" & vbCrLf & "骨器種類：" & kotsukiName, vbInformation
    Else
        '名入れ無し
        NamingNoneOpt.value = True
        MsgBox "【名入れなし】" & vbCrLf & "骨器種類：" & kotsukiName, vbInformation
    End If
End Sub

'情報登録処理
Public Sub RegistInfo()
    '名入れ有無設定
    Dim namingInfo As String
    If namingFlg = True Then
        namingInfo = "有り"
    Else
        namingInfo = "-"
    End If
    '施行コードの最終行の次の行を取得
    Dim insertRow As Long: insertRow = TableLastRow() + 5
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '通常オプションボタン選択時
    If TypeNormalOpt.value = True Then
        With Worksheets(DivisionCbo.value)
            .Cells(insertRow, colNumberCeremonyCode).value = CeremonyCodeTxt.value
            .Cells(insertRow, colNumberName).value = NameTxt.value
            .Cells(insertRow, colNumberStatus).value = "受付"
            .Cells(insertRow, colNumberRcptDate).value = Date
            .Cells(insertRow, colNumberAddress).value = AddressCbo.value
            .Cells(insertRow, colNumberNaming).value = namingInfo
            .Cells(insertRow, colNumberNokanDate).value = DateTxt1.value & " " & TimesTxt1.value
            .Cells(insertRow, colNumberTsuyaDate).value = DateTxt2.value & " " & TimesTxt2.value
            .Cells(insertRow, colNumberKokubetsushikiDate).value = DateTxt3.value & " " & TimesTxt3.value
            .Cells(insertRow, colNumberNotes).value = NotesTxt.value
            '葬消耗品
            Call registItems(insertRow)
        End With
    '規定在庫オプションボタン選択時
    ElseIf TypeStockOpt.value = True Then
        With Worksheets(DivisionCbo.value)
            .Cells(insertRow, colNumberCeremonyCode).value = CeremonyCodeTxt.value
            .Cells(insertRow, colNumberName).value = NameTxt.value
            .Cells(insertRow, colNumberStatus).value = "受付"
            .Cells(insertRow, colNumberRcptDate).value = Date
            .Cells(insertRow, colNumberAddress).value = AddressCbo.value
            .Cells(insertRow, colNumberNaming).value = namingInfo
            .Cells(insertRow, colNumberNokanDate).value = DateTxt1.value & " " & TimesTxt1.value
            .Cells(insertRow, colNumberTsuyaDate).value = DateTxt2.value & " " & TimesTxt2.value
            .Cells(insertRow, colNumberKokubetsushikiDate).value = DateTxt3.value & " " & TimesTxt3.value
            .Cells(insertRow, colNumberNotes).value = NotesTxt.value
            '葬消耗品
            Call registItems(insertRow)
        End With
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Dim Msg As String, title As String, res As Integer
    Msg = "登録が完了しました。続けて登録しますか？"
    title = "新規登録完了"
    res = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, title)
    '「いいえ」の場合は処理中止
    If res = vbNo Then Unload RegisterForm
    '「はい」の場合は入力値を消去する
    Call CtrlsClear(Me.Controls, False)
End Sub

'葬消耗品の登録処理
Function registItems(ByVal insertRow As Long)
    '商品コード登録
    With Worksheets(DivisionCbo.value)
        If itemCode1.value <> "" Then .Cells(insertRow, colNumberItemCode1).value = itemCode1.value
        If itemCode2.value <> "" Then .Cells(insertRow, colNumberItemCode2).value = itemCode2.value
        If itemCode3.value <> "" Then .Cells(insertRow, colNumberItemCode3).value = itemCode3.value
        If itemCode4.value <> "" Then .Cells(insertRow, colNumberItemCode4).value = itemCode4.value
        If itemCode5.value <> "" Then .Cells(insertRow, colNumberItemCode5).value = itemCode5.value
        If itemCode6.value <> "" Then .Cells(insertRow, colNumberItemCode6).value = itemCode6.value
        If itemCode7.value <> "" Then .Cells(insertRow, colNumberItemCode7).value = itemCode7.value
        If itemCode8.value <> "" Then .Cells(insertRow, colNumberItemCode8).value = itemCode8.value
        If itemCode9.value <> "" Then .Cells(insertRow, colNumberItemCode9).value = itemCode9.value
        If itemCode10.value <> "" Then .Cells(insertRow, colNumberItemCode10).value = itemCode10.value
        If itemCode11.value <> "" Then .Cells(insertRow, colNumberItemCode11).value = itemCode11.value
        If itemCode12.value <> "" Then .Cells(insertRow, colNumberItemCode12).value = itemCode12.value
        If itemCode13.value <> "" Then .Cells(insertRow, colNumberItemCode13).value = itemCode13.value
        If itemCode14.value <> "" Then .Cells(insertRow, colNumberItemCode14).value = itemCode14.value
        If itemCode15.value <> "" Then .Cells(insertRow, colNumberItemCode15).value = itemCode15.value
        '商品数量登録
        If itemCode1.value <> "" Then .Cells(insertRow, colNumberItemQty1).value = itemQty1.value
        If itemCode2.value <> "" Then .Cells(insertRow, colNumberItemQty2).value = itemQty2.value
        If itemCode3.value <> "" Then .Cells(insertRow, colNumberItemQty3).value = itemQty3.value
        If itemCode4.value <> "" Then .Cells(insertRow, colNumberItemQty4).value = itemQty4.value
        If itemCode5.value <> "" Then .Cells(insertRow, colNumberItemQty5).value = itemQty5.value
        If itemCode6.value <> "" Then .Cells(insertRow, colNumberItemQty6).value = itemQty6.value
        If itemCode7.value <> "" Then .Cells(insertRow, colNumberItemQty7).value = itemQty7.value
        If itemCode8.value <> "" Then .Cells(insertRow, colNumberItemQty8).value = itemQty8.value
        If itemCode9.value <> "" Then .Cells(insertRow, colNumberItemQty9).value = itemQty9.value
        If itemCode10.value <> "" Then .Cells(insertRow, colNumberItemQty10).value = itemQty10.value
        If itemCode11.value <> "" Then .Cells(insertRow, colNumberItemQty11).value = itemQty11.value
        If itemCode12.value <> "" Then .Cells(insertRow, colNumberItemQty12).value = itemQty12.value
        If itemCode13.value <> "" Then .Cells(insertRow, colNumberItemQty13).value = itemQty13.value
        If itemCode14.value <> "" Then .Cells(insertRow, colNumberItemQty14).value = itemQty14.value
        If itemCode15.value <> "" Then .Cells(insertRow, colNumberItemQty15).value = itemQty15.value
    End With
End Function

'通常オプションボタンを選択時
Private Sub TypeNormalOpt_Click()
    Call CheckEnable(Me.Controls)
End Sub

'規定在庫オプションボタンを選択時
Private Sub TypeStockOpt_Click()
    Call CheckEnable(Me.Controls)
End Sub

'名入れ有りオプションボタン押下時
Private Sub NamingExistOpt_Click()
    namingFlg = True
End Sub

'名入れ無しオプションボタン押下時
Private Sub NamingNoneOpt_Click()
    namingFlg = False
End Sub

'フォームの各コントロールの有効・無効を切り替え
Public Sub CheckEnable(ctrls As Controls)
    Dim ctrl As Control
    
    '通常ボタン選択時
    If TypeNormalOpt.value = True Then
        CeremonyCodeTxt.Locked = False
        NameTxt.Locked = False
        For Each ctrl In ctrls
            ctrl.Enabled = True
        Next
        CeremonyCodeTxt.value = ""
        NameTxt.value = ""
        
    '規定在庫ボタン選択時
    ElseIf TypeStockOpt.value = True Then
        CeremonyCodeTxt.value = GenerateUniqueNumber()
        CeremonyCodeTxt.Locked = True
        NameTxt.value = "規定在庫"
        For Each ctrl In ctrls
            If ctrl.Name = "CeremonyCodeLbl" Or _
               ctrl.Name = "CeremonyCodeTxt" Or _
               ctrl.Name = "UrnNamingLbl" Or _
               ctrl.Name = "NamingExistOpt" Or _
               ctrl.Name = "NamingNoneOpt" Then
                ctrl.Enabled = False
            Else
                ctrl.Enabled = True
            End If
        Next
    End If
End Sub

'規定在庫用の施行コード作成
Function GenerateUniqueNumber() As String
    Dim currentDateTime As Double
    Dim dayValue As Integer
    Dim hourValue As Integer
    Dim minuteValue As Integer
    Dim secondValue As Integer
    currentDateTime = Now
    dayValue = day(currentDateTime)
    hourValue = hour(currentDateTime)
    minuteValue = minute(currentDateTime)
    secondValue = second(currentDateTime)
    GenerateUniqueNumber = "9" & Format(dayValue, "00") & Format(hourValue, "00") & Format(minuteValue, "00") & Format(secondValue, "00")
End Function

' キャンセルボタンクリック時にフォームを閉じる
Private Sub CancelBtn_Click()
    Unload RegisterForm
End Sub

' クリアボタンのクリックイベント。クリアプロシージャを呼び出す
Private Sub ClearBtn_Click()
    Call CtrlsClear(Me.Controls, False)
End Sub

'ユーザーフォームの入力項目をすべてクリア
'blListClear：Trueの場合はコンボボックスとリストボックスの候補もクリア
Public Sub CtrlsClear(ctrls As Controls, Optional blListClear As Boolean = False)
    'コントロールコレクション(Controls)から1つずつ取り出してクリア
    Dim ctrl As Control
    For Each ctrl In ctrls
        Select Case TypeName(ctrl)
        Case "TextBox", "RefEdit"
            ctrl.value = ""
        Case "CheckBox", "OptionButton", "ToggleButton"
            ctrl.value = False
        Case "ComboBox", "ListBox"
            ctrl.value = ""
            If blListClear Then ctrl.Clear
        End Select
    Next
    '商品名ラベルの値をクリア
    Dim i As Integer
    Dim targetName As String
    For i = 1 To 15
        targetName = "itemName" & i
        RegisterForm.Controls(targetName).Caption = ""
    Next i
    '初期値を再設定
    Call SetInitialValue
End Sub

' 事業部コンボボックスの値が変更されたとき
Private Sub DivisionCbo_Change()
    '届け先の値をクリア
    AddressCbo.value = ""
    Dim division As String
    division = DivisionCbo.value
    Select Case division
    Case "西"
        addressWest = Array("CBO", "久野", "小田会", "鴨宮", "岩原", "栢山", "大井", "二宮", "小田原セレモ", "成田セレモ", "その他")
        AddressCbo.List() = addressWest
    Case "中央"
        addressCenter = Array("CBH", "平会", "真土", "田村", "追分", "金目", "国府", "秦野", "秦野EP", "渋沢", "伊勢原", "鶴巻", "愛甲", "その他")
        AddressCbo.List() = addressCenter
    Case "県央"
        addressCentralPref = Array("本厚木", "入谷", "広野台", "桜ヶ丘", "その他")
        AddressCbo.List() = addressCentralPref
    Case "南"
        addressShonan = Array("茅ヶ崎", "寒川", "西久保", "赤松", "南湖", "松浪", "その他")
        AddressCbo.List() = addressShonan
    Case "東"
        addressEast = Array("CBF", "辻堂会堂", "本鵠沼", "辻堂元町", "片瀬鵠沼", "手広", "由比ヶ浜", "西富", "六会", "長後", "秋葉台", "その他")
        AddressCbo.List() = addressEast
    Case "西東京"
        addressWestTokyo = Array("町田", "相模大野", "木曽", "淵野辺", "相模原", "その他")
        AddressCbo.List() = addressWestTokyo
    End Select
End Sub

' 施行コードの入力制限
Private Sub CeremonyCodeTxt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

' 納棺日付テキストボックスの入力制限
Private Sub DateTxt1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like "/" Then KeyAscii = 0
End Sub

' 通夜日付テキストボックスの入力制限
Private Sub DateTxt2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like "/" Then KeyAscii = 0
End Sub

'告別式日付テキストボックスの入力制限
Private Sub DateTxt3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like "/" Then KeyAscii = 0
End Sub

' 納棺時刻テキストボックス
Private Sub TimesTxt1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like ":" Then KeyAscii = 0
End Sub

' 通夜時刻テキストボックスの入力制限
Private Sub TimesTxt2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like ":" Then KeyAscii = 0
End Sub

' 告別式時刻テキストボックス 入力時に呼び出し
Private Sub TimesTxt3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like ":" Then KeyAscii = 0
End Sub

' 施行コード 他のコントロールに移動する直前
Private Sub CeremonyCodeTxt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(CeremonyCodeTxt.Text) = 0 Then
        Exit Sub
    ElseIf Len(CeremonyCodeTxt.Text) <> 9 Then
        MsgBox "施行コードは9桁で入力してください。"
        Cancel = True
    End If
End Sub

' 日付のバリデーション
Private Sub ValidateDateInput(DateTxt As MSForms.TextBox, ByVal Cancel As MSForms.ReturnBoolean)
    Dim isValid As Boolean
    If Len(DateTxt.value) = 0 Then Exit Sub
    ' 日付のバリデーション
    isValid = IsDate(DateTxt.value)
    If isValid Then
        Exit Sub
    Else
        MsgBox "正しい日付を入力してください。"
        DateTxt.value = ""
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub DateTxt1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateDateInput DateTxt1, Cancel
End Sub

Private Sub DateTxt2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateDateInput DateTxt2, Cancel
End Sub

Private Sub DateTxt3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateDateInput DateTxt3, Cancel
End Sub

' 時刻のバリデーション
Private Sub ValidateTimeInput(TimesTxt As MSForms.TextBox, ByVal Cancel As MSForms.ReturnBoolean)
    Dim isValid As Boolean
    If Len(TimesTxt.value) = 0 Then Exit Sub
    ' 時刻のバリデーション
    isValid = IsTime(TimesTxt.value)
    If Len(TimesTxt.value) <= 3 Or Not isValid Then
        MsgBox "正しい時刻を入力してください。"
        TimesTxt.value = ""
        Cancel = True
    ElseIf isValid Then
        Exit Sub
    End If
End Sub

'バリデーション処理
Function IsTime(timeString As String) As Boolean
    On Error Resume Next
    Dim tempTime As Date
    tempTime = TimeValue(timeString)
    IsTime = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub TimesTxt1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateTimeInput TimesTxt1, Cancel
End Sub

Private Sub TimesTxt2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateTimeInput TimesTxt2, Cancel
End Sub

Private Sub TimesTxt3_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateTimeInput TimesTxt3, Cancel
End Sub

' 数量コンボボックスを数字のみ入力可にする
Private Sub itemQty1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQt12y_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

Private Sub itemQty15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

'商品の名称と数量のクリア処理
Private Sub clearFields(index As Integer)
    Controls("itemName" & index).Caption = ""
    Controls("itemQty" & index).value = ""
End Sub

Private Sub itemCode1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 1
End Sub

Private Sub itemCode2_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 2
End Sub

Private Sub itemCode3_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 3
End Sub

Private Sub itemCode4_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 4
End Sub

Private Sub itemCode5_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 5
End Sub

Private Sub itemCode6_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 6
End Sub

Private Sub itemCode7_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 7
End Sub

Private Sub itemCode8_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 8
End Sub

Private Sub itemCode9_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 9
End Sub

Private Sub itemCode10_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 10
End Sub

Private Sub itemCode11_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 11
End Sub

Private Sub itemCode12_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 12
End Sub

Private Sub itemCode13_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 13
End Sub

Private Sub itemCode14_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 14
End Sub

Private Sub itemCode15_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    clearFields 15
End Sub

'商品コードのバリデーションと商品名の表示処理
Private Sub ValidateAndUpdate(index As Integer, ByVal Cancel As MSForms.ReturnBoolean)
    Dim itemCode As String
    Dim itemName As String
    itemCode = Controls("itemCode" & index).Text
    itemName = GetItemName(itemCode)

    If Len(itemCode) <> 5 Or itemName = "" Then
        If Len(itemCode) <> 0 Then
            MsgBox "正しい商品コードを入力してください。"
            Controls("itemCode" & index).Text = ""
            Cancel = True
        End If
        Exit Sub
    End If
    Controls("itemName" & index).Caption = itemName
    Controls("itemQty" & index).value = 1
End Sub

'商品名を取得
Private Function GetItemName(itemCode As String) As String
    Dim targetSheet As String: targetSheet = "マスタ"
    Dim itemCodeColumn As Integer: itemCodeColumn = 6
    Dim itemNameColumn As Integer: itemNameColumn = 7
    Dim itemCodeLastRow As Integer
    Dim i As Integer
    itemCodeLastRow = Worksheets(targetSheet).Cells(Rows.Count, itemCodeColumn).End(xlUp).row
    For i = DROP_ROWS To itemCodeLastRow
        If itemCode = GetCellValue(targetSheet, i, itemCodeColumn) Then
            GetItemName = GetCellValue(targetSheet, i, itemNameColumn)
            Exit Function
        End If
    Next i
    GetItemName = ""
End Function

'任意のセル値を取得
Private Function GetCellValue(sheetName As String, row As Integer, col As Integer) As Variant
    GetCellValue = Worksheets(sheetName).Cells(row, col).value
End Function

Private Sub itemCode1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 1, Cancel
End Sub

Private Sub itemCode2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 2, Cancel
End Sub

Private Sub itemCode3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 3, Cancel
End Sub

Private Sub itemCode4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 4, Cancel
End Sub

Private Sub itemCode5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 5, Cancel
End Sub

Private Sub itemCode6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 6, Cancel
End Sub

Private Sub itemCode7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 7, Cancel
End Sub

Private Sub itemCode8_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 8, Cancel
End Sub

Private Sub itemCode9_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 9, Cancel
End Sub

Private Sub itemCode10_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 10, Cancel
End Sub

Private Sub itemCode11_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 11, Cancel
End Sub

Private Sub itemCode12_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 12, Cancel
End Sub

Private Sub itemCode13_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 13, Cancel
End Sub

Private Sub itemCode14_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 14, Cancel
End Sub

Private Sub itemCode15_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAndUpdate 15, Cancel
End Sub



