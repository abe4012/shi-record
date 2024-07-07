VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Create_RepNo 
   Caption         =   "�Ӓ菑�ԍ����s���"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   OleObjectBlob   =   "Create_RepNo.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Create_RepNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label_hakkonengappi_Click()

End Sub

' Form�ǂݍ��ݎ�����������
Private Sub UserForm_Initialize()
    
    Call ShipInspectRecord_Search
    Call ShipInspectRecord_ShowAll

    ' ���s�N�����̏�����
    ' ���݂̔N�A���A�����擾
    Dim CurrentYear As Integer, CurrentMonth As Integer, CurrentDay As Integer
    CurrentYear = Year(Date)
    CurrentMonth = Month(Date)
    CurrentDay = Day(Date)
    ' txtBox_CurrentY�Ɍ��݂̔N��ݒ�
    txtBox_CurrentY.Value = CurrentYear
    
    ' cmbBox_CurrentM�Ɍ��̑I������ǉ��i1��?12���j
    Dim months As Integer
    For months = 1 To 12
        cmbBox_CurrentM.AddItem months
    Next months
    
    ' cmbBox_CurrentD�ɓ��̑I������ǉ��i1��?31���j
    Dim Days As Integer
    For Days = 1 To 31
        cmbBox_CurrentD.AddItem Days
    Next Days
    
    ' ComboBox�̌��݂̌��Ɠ���I��
    cmbBox_CurrentM.Value = CurrentMonth
    cmbBox_CurrentD.Value = CurrentDay
    
    
    ' �w�b�_�[���X�g�\���̏�����
    Call InitializeCmbBox(Me.cmbBox_Header, GetNamedRange("HeaderList"))
    
    ' �N�x�ϊ�
    Dim currentFiscalY As Integer
    If CurrentMonth <= 3 Then
        currentFiscalY = CurrentYear - 1
    Else
        currentFiscalY = CurrentYear
    End If
    
    cmbBox_Header.Value = SearchFiscalYheader(currentFiscalY)
    

End Sub

' "�Č�����"�{�^������������
Sub Button_Search_Click()

    Dim FiscalY As String, RefNum As String
    ' Dim FiscalY_RefNum As String
    
    FiscalY = TrimAndHalfDigit(Me.txtBox_FiscalY)
    RefNum = TrimAndHalfDigit(Me.txtBox_RefNum)
    
    If FiscalY = "0" Or RefNum = "0" Then
        MsgBox "��t�N�x �������� ��tNo. �ɒl�����͂���Ă��܂���B", , "�G���["
        Exit Sub
    Else
    End If
    
    ' FiscalY_RefNum = FiscalY & RefNum
    
    Dim SearchInspectRec As Object
    Set SearchInspectRec = SearchRec_RefID("", FiscalY, RefNum)
    ' Set SearchInspectRec = SearchRec_RefID(FiscalY_RefNum, "", "")
    
    ' �������ʂ�Label�ɏ�������
    Label_FStat = SearchInspectRec("stat")
    Label_FFiscalY = SearchInspectRec("year")
    Label_FRefNum = SearchInspectRec("refNum")
    Label_FReceiptDate = SearchInspectRec("receiptDate")
    Label_FShipNMandType = SearchInspectRec("shipName") & "�@�@(" & SearchInspectRec("shipType") & ")"
    Label_FOwner = SearchInspectRec("owner")
    Label_FDelegateAndStaff = SearchInspectRec("delegater") & "�@" & SearchInspectRec("delegateStaff")
    Label_FInspectType = SearchInspectRec("inspectType")
    Label_FClause = SearchInspectRec("clause")
    Label_FConcurrentInspection = SearchInspectRec("concurrentInspection")
    Label_FShipyard = SearchInspectRec("shipyard")
    Label_FInspectDate = SearchInspectRec("inspectDate")
    Label_FKmsStaff = SearchInspectRec("kmsStaff")
    Label_FRepNoCreateDate = SearchInspectRec("repNoCreateDate")
    Label_FRepNo = SearchInspectRec("repNo")
    
    If SearchInspectRec("repNoCreateDate") <> "" Or SearchInspectRec("repNoCreateDate") <> "" Then
        MsgBox "���s���������͊Ӓ菑�ԍ������ɔ��s����Ă��܂��B" & Chr(13) & "���̂܂ܔԍ��𔭍s����ꍇ�́A�Ċm�F�̏�A�s���Ă��������B", , "�x��"
    Else
    End If
    
End Sub

' "�Ӓ菑�ԍ����s"�{�^������������
Private Sub Button_CreateRepNo_Click()

    Dim FiscalY As String, RefNum As String
    ' Dim FiscalY_RefNum As String
    Dim RepNoManual As Boolean
    RepNoManual = False
    
    Dim CreateDate As String
    CreateDate = TrimAndHalfDigit(Me.txtBox_CurrentY) & "/" & TrimAndHalfDigit(Me.cmbBox_CurrentM) & "/" & TrimAndHalfDigit(Me.cmbBox_CurrentD)
    ' Debug.Print (CreateDate)
    
    FiscalY = TrimAndHalfDigit(Me.txtBox_FiscalY)
    RefNum = TrimAndHalfDigit(Me.txtBox_RefNum)
    
    If FiscalY = "0" Or RefNum = "0" Then
        MsgBox "��t�N�x �������� ��tNo. �ɒl�����͂���Ă��܂���B", , "�G���["
        Exit Sub
    ElseIf FiscalY <> TrimAndHalfDigit(Me.Label_FFiscalY) Or RefNum <> TrimAndHalfDigit(Me.Label_FRefNum) Then
        MsgBox "���͂���Ă��� ��t�N�x, ��tNo. ���\������Ă���Č����̂��̂ƈقȂ�܂��B" & Chr(13) & "��x�u�Č������v�{�^���������ĈČ���񂪐��������m�F���Ă��������B", , "�G���["
        Exit Sub
    ElseIf Not CheckCorrectDate(CreateDate) Then
        MsgBox "�w�肳�ꂽ���s�N�����͓��t�Ƃ��đ��݂��܂���B", , "�G���["
        Exit Sub
    ElseIf chk_RepNoManual.Value = True Then
        If txtBox_RepNoManual.Value <> "" Then
            RepNoManual = True
        Else
            MsgBox "�Ӓ菑�ԍ�����͂��Ă��������B" & Chr(13) & "(�������͂���ꍇ�́u�Ӓ菑�ԍ�����́v��" & Chr(13) & "�`�F�b�N���O���Ă��������B)", , "�G���["
            Exit Sub
        End If
    ElseIf Me.cmbBox_Header = "" Then
        MsgBox "�w�b�_�[��I�����Ă��������B", , "�G���["
        Exit Sub
    Else
    End If
    
    ' �Ӓ菑�ԍ��擾
    Dim repNo As String
    If RepNoManual = True Then
        repNo = txtBox_RepNoManual.Value
    Else
        repNo = cmbBox_Header & GetNewRepNo(cmbBox_Header, 1)
    End If
    
    ' FiscalY_RefNum = FiscalY & RefNum
    
    Dim SearchInspectRec As Object, SearchInspectRecAddr As Object
    Set SearchInspectRec = SearchRec_RefID("", FiscalY, RefNum)
    ' Set SearchInspectRec = SearchRec_RefID(FiscalY_RefNum, "", "")
    Set SearchInspectRecAddr = SearchRec_RefID("", FiscalY, RefNum, "Address")
    ' Set SearchInspectRecAddr = SearchRec_RefID(FiscalY_RefNum, "", "", "Address")

    ' Debug.Print (repNo)
    
    Dim CheckCreateRepNo As Integer
    CheckCreateRepNo = MsgBox("�ȉ��̓��e�ŊӒ菑�ԍ��𔭍s���܂��B" & Chr(13) & "�m�F�̏�A��薳����΁u�͂�(Yes)�v�������ĉ������B" & Chr(13) & Chr(13) & "�N�x�F" & SearchInspectRec("year") & Chr(13) & "��tNo.�F" & SearchInspectRec("refNum") & Chr(13) & "��t���F" & SearchInspectRec("receiptDate") & Chr(13) & "�D��(��)�F" & SearchInspectRec("shipName") & " (" & SearchInspectRec("shipType") & ")" & Chr(13) & "�S���ҁF" & SearchInspectRec("kmsStaff") & Chr(13) & Chr(13) & "���s���F" & CreateDate & Chr(13) & "�Ӓ菑�ԍ��F" & repNo, vbQuestion + vbYesNo + vbDefaultButton2, "�Ӓ菑�ԍ����s�O�m�F")
    If CheckCreateRepNo = vbYes Then
    
        Dim CreateRepNo As String, CreateRepNoCreateDate As String
        ' Debug.Print RewriteCaseItems("", FiscalY, RefNum, "repNo", repNo)
        ' Debug.Print RewriteCaseItems("", FiscalY, RefNum, "repNoCreateDate", CreateDate)
        CreateRepNo = RewriteCaseItems("", FiscalY, RefNum, "repNo", repNo)
        CreateRepNoCreateDate = RewriteCaseItems("", FiscalY, RefNum, "repNoCreateDate", CreateDate)
        
        Dim CheckExitForm As Integer
        CheckExitForm = MsgBox("�Ӓ菑�𔭍s���܂����B" & Chr(13) & Chr(13) & "�Ӓ菑���s��ʂ���܂����H", vbQuestion + vbYesNo + vbDefaultButton2, "�t�H�[���I���m�F")
        If CheckExitForm = vbYes Then
            Unload Me
        Else
        End If
    
    Else
    End If

End Sub


' "����"�{�^������
Private Sub Button_Exit_Click()
    Unload Me
End Sub

