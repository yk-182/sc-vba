VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegisterForm 
   Caption         =   "�o�^�t�H�[��"
   ClientHeight    =   7546
   ClientLeft      =   110
   ClientTop       =   451
   ClientWidth     =   12760
   OleObjectBlob   =   "RegisterForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "RegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�s���C���l
Const DROP_ROWS As Integer = 4
'���햼����L��
Dim namingFlg As Boolean
'�͂���z��ϐ�
Dim addressWest As Variant
Dim addressCenter As Variant
Dim addressCentralPref As Variant
Dim addressShonan As Variant
Dim addressEast As Variant
Dim addressWestTokyo As Variant
'�f�[�^�ǉ����̊e��
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

' �t�H�[�����������ɌĂяo��
Private Sub UserForm_Initialize()
    Dim areaData As Variant
    areaData = Array("��", "����", "����", "��", "��", "������")
    With DivisionCbo
        .List() = areaData
        .Style = fmStyleDropDownList
    End With
    '�����l�ݒ�
    Call SetInitialValue
End Sub

'�����l����
Public Sub SetInitialValue()
    '�ʏ�I�v�V�����{�^����ON
    TypeNormalOpt.value = True
    '�N�̏����l��ݒ�
    YearTxt1.value = Year(Date)
    YearTxt2.value = Year(Date)
    YearTxt3.value = Year(Date)
    '�J�����񐔂�ݒ�
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

'�o�^�{�^���N���b�N���̏���
Private Sub RegisterBtn_Click()
    '�o�^�ۃt���O
    Dim isEnabled As Boolean: isEnabled = True
    If TypeNormalOpt.value = True Then
        '�ʏ�I�v�V�����{�^���I�����̃o���f�[�V����
        Call checkNormalTypeValues(isEnabled)
        If isEnabled = False Then
            Exit Sub
        End If
    Else
        '�K��݌ɃI�v�V�����{�^���I�����̃o���f�[�V����
        Call checkStockTypeValues(isEnabled)
        If isEnabled = False Then
            Exit Sub
        End If
    End If
    
    '�m�F���b�Z�[�W�\��
    Dim Msg As String, title As String, res As Integer
    Msg = "�o�^�������s���܂��B��낵���ł����H"
    title = "�V�K�o�^�m�F"
    res = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, title)
    '�u�������v�̏ꍇ�͏������~
    If res = vbNo Then Exit Sub
    '�o�^����
    Call RegistInfo
    
End Sub

'�ʏ�^�C�v�̃o���f�[�V����
Public Sub checkNormalTypeValues(isEnabled)
    If CeremonyCodeTxt.value = "" Then
        MsgBox "�{�s�R�[�h����͂��Ă��������B", vbCritical
        CeremonyCodeTxt.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If NameTxt.value = "" Then
        MsgBox "���Ɩ�����͂��Ă��������B", vbCritical
        NameTxt.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If DivisionCbo.value = "" Then
        MsgBox "���ƕ���I�����Ă��������B", vbCritical
        DivisionCbo.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If AddressCbo.value = "" Then
        MsgBox "�͂����I�����Ă��������B", vbCritical
        AddressCbo.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If YearTxt3.value = "" Or DateTxt3.value = "" Or TimesTxt3.value = "" Then
        MsgBox "���ʎ���������͂��Ă��������B", vbCritical
        DateTxt3.SetFocus
        isEnabled = False
        Exit Sub
    End If
    '�{�s�R�[�h�����͂���Ă���ꍇ�͏d���`�F�b�N
    If CeremonyCodeTxt.value <> "" Then
        '���ƕ�
        Dim division As String: division = DivisionCbo.value
        '�{�s�R�[�h
        Dim ceremonyCode As String: ceremonyCode = CeremonyCodeTxt.value
        '�{�s�R�[�h��̍ŏI�s���擾
        Dim lastRow As Long: lastRow = TableLastRow()
        If WorksheetFunction.CountIf(Range(Worksheets(division).Cells(5, 3), Worksheets(division).Cells(lastRow + 4, 3)), ceremonyCode) >= 1 Then
            MsgBox "����̎{�s�R�[�h�����ɓo�^����Ă��܂��B", vbCritical
            CeremonyCodeTxt.SetFocus
            isEnabled = False
            Exit Sub
        End If
    End If
    '���i���ʂ̓��̓`�F�b�N
    If checkItemQtyValue() = False Then
        MsgBox "���i���ʂ���͂��Ă��������B", vbCritical
        isEnabled = False
        Exit Sub
    End If
End Sub

'�K��݌Ƀ^�C�v�̃o���f�[�V����
Public Sub checkStockTypeValues(isEnabled)
    If NameTxt.value = "" Then
        MsgBox "���Ɩ�����͂��Ă��������B", vbCritical
        NameTxt.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If DivisionCbo.value = "" Then
        MsgBox "���ƕ���I�����Ă��������B", vbCritical
        DivisionCbo.SetFocus
        isEnabled = False
        Exit Sub
    End If
    If AddressCbo.value = "" Then
        MsgBox "�͂����I�����Ă��������B", vbCritical
        AddressCbo.SetFocus
        isEnabled = False
        Exit Sub
    End If
    '���i���ʂ̓��̓`�F�b�N
    If checkItemQtyValue() = False Then
        MsgBox "���i���ʂ���͂��Ă��������B", vbCritical
        isEnabled = False
        Exit Sub
    End If
End Sub

'�e�[�u���̃f�[�^�ŏI�s���擾����
Function TableLastRow() As Long
    '�Ώۃe�[�u��
    Dim targetTable As ListObject
    Set targetTable = Worksheets(DivisionCbo.value).ListObjects(1)
    Dim codeColumn As Long: codeColumn = targetTable.ListColumns("�{�s�R�[�h").index
    '�e�[�u���̉����珇�Ƀf�[�^�̓����Ă���s��T��
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

'���i���ʂ̓��̓`�F�b�N
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


'��������擾����
Private Sub NamingBtn_Click()
    
    Dim serchValue As String                     '�����l
    Dim serchRange As Range                      '�����͈�
    Dim kotsukiName As String                    '���햼��
    Dim namingValue As Long                      '�����ꔻ��l
    '���̓`�F�b�N
    If CeremonyCodeTxt.value = "" Or DivisionCbo.value = "" Then
        MsgBox "�{�s�R�[�h�Ǝ��ƕ�����͂��Ă��������B", vbCritical
        Exit Sub
    End If
    '�����l��ݒ�
    serchValue = DivisionCbo.value & Right(CeremonyCodeTxt.value, 4)
    '�����͈͂�ݒ�
    Set serchRange = Worksheets("list").Range("F2:H900")

    On Error Resume Next                         '�G���[�����ɐݒ�
    '�����l����list�V�[�g�̈�v���鍜�햼�Ɩ����ꔻ�ʒl���擾
    kotsukiName = WorksheetFunction.VLookup(serchValue, serchRange, 2, False)
    If Err.Number > 0 Then
        MsgBox "����������擾�ł��܂���B��Ƀf�[�^����荞��ł��������B", vbCritical
        On Error GoTo 0                          '�G���[�����ɖ߂�
        Exit Sub
    End If
    namingValue = WorksheetFunction.VLookup(serchValue, serchRange, 3, False)
    On Error GoTo 0                              '�G���[�����ɖ߂�
    
    If namingValue < 0 Then
        '������L��
        NamingExistOpt.value = True
        MsgBox "�y������L��z" & vbCrLf & "�����ށF" & kotsukiName, vbInformation
    Else
        '�����ꖳ��
        NamingNoneOpt.value = True
        MsgBox "�y������Ȃ��z" & vbCrLf & "�����ށF" & kotsukiName, vbInformation
    End If
End Sub

'���o�^����
Public Sub RegistInfo()
    '������L���ݒ�
    Dim namingInfo As String
    If namingFlg = True Then
        namingInfo = "�L��"
    Else
        namingInfo = "-"
    End If
    '�{�s�R�[�h�̍ŏI�s�̎��̍s���擾
    Dim insertRow As Long: insertRow = TableLastRow() + 5
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '�ʏ�I�v�V�����{�^���I����
    If TypeNormalOpt.value = True Then
        With Worksheets(DivisionCbo.value)
            .Cells(insertRow, colNumberCeremonyCode).value = CeremonyCodeTxt.value
            .Cells(insertRow, colNumberName).value = NameTxt.value
            .Cells(insertRow, colNumberStatus).value = "��t"
            .Cells(insertRow, colNumberRcptDate).value = Date
            .Cells(insertRow, colNumberAddress).value = AddressCbo.value
            .Cells(insertRow, colNumberNaming).value = namingInfo
            .Cells(insertRow, colNumberNokanDate).value = DateTxt1.value & " " & TimesTxt1.value
            .Cells(insertRow, colNumberTsuyaDate).value = DateTxt2.value & " " & TimesTxt2.value
            .Cells(insertRow, colNumberKokubetsushikiDate).value = DateTxt3.value & " " & TimesTxt3.value
            .Cells(insertRow, colNumberNotes).value = NotesTxt.value
            '�����Օi
            Call registItems(insertRow)
        End With
    '�K��݌ɃI�v�V�����{�^���I����
    ElseIf TypeStockOpt.value = True Then
        With Worksheets(DivisionCbo.value)
            .Cells(insertRow, colNumberCeremonyCode).value = CeremonyCodeTxt.value
            .Cells(insertRow, colNumberName).value = NameTxt.value
            .Cells(insertRow, colNumberStatus).value = "��t"
            .Cells(insertRow, colNumberRcptDate).value = Date
            .Cells(insertRow, colNumberAddress).value = AddressCbo.value
            .Cells(insertRow, colNumberNaming).value = namingInfo
            .Cells(insertRow, colNumberNokanDate).value = DateTxt1.value & " " & TimesTxt1.value
            .Cells(insertRow, colNumberTsuyaDate).value = DateTxt2.value & " " & TimesTxt2.value
            .Cells(insertRow, colNumberKokubetsushikiDate).value = DateTxt3.value & " " & TimesTxt3.value
            .Cells(insertRow, colNumberNotes).value = NotesTxt.value
            '�����Օi
            Call registItems(insertRow)
        End With
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Dim Msg As String, title As String, res As Integer
    Msg = "�o�^���������܂����B�����ēo�^���܂����H"
    title = "�V�K�o�^����"
    res = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, title)
    '�u�������v�̏ꍇ�͏������~
    If res = vbNo Then Unload RegisterForm
    '�u�͂��v�̏ꍇ�͓��͒l����������
    Call CtrlsClear(Me.Controls, False)
End Sub

'�����Օi�̓o�^����
Function registItems(ByVal insertRow As Long)
    '���i�R�[�h�o�^
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
        '���i���ʓo�^
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

'�ʏ�I�v�V�����{�^����I����
Private Sub TypeNormalOpt_Click()
    Call CheckEnable(Me.Controls)
End Sub

'�K��݌ɃI�v�V�����{�^����I����
Private Sub TypeStockOpt_Click()
    Call CheckEnable(Me.Controls)
End Sub

'������L��I�v�V�����{�^��������
Private Sub NamingExistOpt_Click()
    namingFlg = True
End Sub

'�����ꖳ���I�v�V�����{�^��������
Private Sub NamingNoneOpt_Click()
    namingFlg = False
End Sub

'�t�H�[���̊e�R���g���[���̗L���E������؂�ւ�
Public Sub CheckEnable(ctrls As Controls)
    Dim ctrl As Control
    
    '�ʏ�{�^���I����
    If TypeNormalOpt.value = True Then
        CeremonyCodeTxt.Locked = False
        NameTxt.Locked = False
        For Each ctrl In ctrls
            ctrl.Enabled = True
        Next
        CeremonyCodeTxt.value = ""
        NameTxt.value = ""
        
    '�K��݌Ƀ{�^���I����
    ElseIf TypeStockOpt.value = True Then
        CeremonyCodeTxt.value = GenerateUniqueNumber()
        CeremonyCodeTxt.Locked = True
        NameTxt.value = "�K��݌�"
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

'�K��݌ɗp�̎{�s�R�[�h�쐬
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

' �L�����Z���{�^���N���b�N���Ƀt�H�[�������
Private Sub CancelBtn_Click()
    Unload RegisterForm
End Sub

' �N���A�{�^���̃N���b�N�C�x���g�B�N���A�v���V�[�W�����Ăяo��
Private Sub ClearBtn_Click()
    Call CtrlsClear(Me.Controls, False)
End Sub

'���[�U�[�t�H�[���̓��͍��ڂ����ׂăN���A
'blListClear�FTrue�̏ꍇ�̓R���{�{�b�N�X�ƃ��X�g�{�b�N�X�̌����N���A
Public Sub CtrlsClear(ctrls As Controls, Optional blListClear As Boolean = False)
    '�R���g���[���R���N�V����(Controls)����1�����o���ăN���A
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
    '���i�����x���̒l���N���A
    Dim i As Integer
    Dim targetName As String
    For i = 1 To 15
        targetName = "itemName" & i
        RegisterForm.Controls(targetName).Caption = ""
    Next i
    '�����l���Đݒ�
    Call SetInitialValue
End Sub

' ���ƕ��R���{�{�b�N�X�̒l���ύX���ꂽ�Ƃ�
Private Sub DivisionCbo_Change()
    '�͂���̒l���N���A
    AddressCbo.value = ""
    Dim division As String
    division = DivisionCbo.value
    Select Case division
    Case "��"
        addressWest = Array("CBO", "�v��", "���c��", "���{", "�⌴", "���R", "���", "��{", "���c���Z����", "���c�Z����", "���̑�")
        AddressCbo.List() = addressWest
    Case "����"
        addressCenter = Array("CBH", "����", "�^�y", "�c��", "�Ǖ�", "����", "���{", "�`��", "�`��EP", "�a��", "�ɐ���", "�ߊ�", "���b", "���̑�")
        AddressCbo.List() = addressCenter
    Case "����"
        addressCentralPref = Array("�{����", "���J", "�L���", "�����u", "���̑�")
        AddressCbo.List() = addressCentralPref
    Case "��"
        addressShonan = Array("������", "����", "���v��", "�ԏ�", "���", "���Q", "���̑�")
        AddressCbo.List() = addressShonan
    Case "��"
        addressEast = Array("CBF", "�ғ��", "�{����", "�ғ�����", "�А�����", "��L", "�R�䃖�l", "���x", "�Z��", "����", "�H�t��", "���̑�")
        AddressCbo.List() = addressEast
    Case "������"
        addressWestTokyo = Array("���c", "���͑��", "�ؑ]", "�����", "���͌�", "���̑�")
        AddressCbo.List() = addressWestTokyo
    End Select
End Sub

' �{�s�R�[�h�̓��͐���
Private Sub CeremonyCodeTxt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" Then KeyAscii = 0
End Sub

' �[�����t�e�L�X�g�{�b�N�X�̓��͐���
Private Sub DateTxt1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like "/" Then KeyAscii = 0
End Sub

' �ʖ���t�e�L�X�g�{�b�N�X�̓��͐���
Private Sub DateTxt2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like "/" Then KeyAscii = 0
End Sub

'���ʎ����t�e�L�X�g�{�b�N�X�̓��͐���
Private Sub DateTxt3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like "/" Then KeyAscii = 0
End Sub

' �[�������e�L�X�g�{�b�N�X
Private Sub TimesTxt1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like ":" Then KeyAscii = 0
End Sub

' �ʖ鎞���e�L�X�g�{�b�N�X�̓��͐���
Private Sub TimesTxt2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like ":" Then KeyAscii = 0
End Sub

' ���ʎ������e�L�X�g�{�b�N�X ���͎��ɌĂяo��
Private Sub TimesTxt3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not Chr(KeyAscii) Like "[0-9]" And Not Chr(KeyAscii) Like ":" Then KeyAscii = 0
End Sub

' �{�s�R�[�h ���̃R���g���[���Ɉړ����钼�O
Private Sub CeremonyCodeTxt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(CeremonyCodeTxt.Text) = 0 Then
        Exit Sub
    ElseIf Len(CeremonyCodeTxt.Text) <> 9 Then
        MsgBox "�{�s�R�[�h��9���œ��͂��Ă��������B"
        Cancel = True
    End If
End Sub

' ���t�̃o���f�[�V����
Private Sub ValidateDateInput(DateTxt As MSForms.TextBox, ByVal Cancel As MSForms.ReturnBoolean)
    Dim isValid As Boolean
    If Len(DateTxt.value) = 0 Then Exit Sub
    ' ���t�̃o���f�[�V����
    isValid = IsDate(DateTxt.value)
    If isValid Then
        Exit Sub
    Else
        MsgBox "���������t����͂��Ă��������B"
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

' �����̃o���f�[�V����
Private Sub ValidateTimeInput(TimesTxt As MSForms.TextBox, ByVal Cancel As MSForms.ReturnBoolean)
    Dim isValid As Boolean
    If Len(TimesTxt.value) = 0 Then Exit Sub
    ' �����̃o���f�[�V����
    isValid = IsTime(TimesTxt.value)
    If Len(TimesTxt.value) <= 3 Or Not isValid Then
        MsgBox "��������������͂��Ă��������B"
        TimesTxt.value = ""
        Cancel = True
    ElseIf isValid Then
        Exit Sub
    End If
End Sub

'�o���f�[�V��������
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

' ���ʃR���{�{�b�N�X�𐔎��̂ݓ��͉ɂ���
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

'���i�̖��̂Ɛ��ʂ̃N���A����
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

'���i�R�[�h�̃o���f�[�V�����Ə��i���̕\������
Private Sub ValidateAndUpdate(index As Integer, ByVal Cancel As MSForms.ReturnBoolean)
    Dim itemCode As String
    Dim itemName As String
    itemCode = Controls("itemCode" & index).Text
    itemName = GetItemName(itemCode)

    If Len(itemCode) <> 5 Or itemName = "" Then
        If Len(itemCode) <> 0 Then
            MsgBox "���������i�R�[�h����͂��Ă��������B"
            Controls("itemCode" & index).Text = ""
            Cancel = True
        End If
        Exit Sub
    End If
    Controls("itemName" & index).Caption = itemName
    Controls("itemQty" & index).value = 1
End Sub

'���i�����擾
Private Function GetItemName(itemCode As String) As String
    Dim targetSheet As String: targetSheet = "�}�X�^"
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

'�C�ӂ̃Z���l���擾
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



