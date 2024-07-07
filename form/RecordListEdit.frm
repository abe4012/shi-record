VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RecordListEdit 
   Caption         =   "�D�������L�^�\�ҏW"
   ClientHeight    =   9390.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14925
   OleObjectBlob   =   "RecordListEdit.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "RecordListEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbBox_PrevInspectType_Change()

End Sub

' Form�ǂݍ��ݎ�����������
Private Sub UserForm_Initialize()
    
    Call ShipInspectRecord_Search
    Call ShipInspectRecord_ShowAll
    
    ' ����cmbBox ������
    ' ���̑I������ǉ��i1��?12���j
    Dim months As Integer
    For months = 1 To 12
        cmbBox_ReceiptM.AddItem months
        cmbBox_RepNoCreateM.AddItem months
        cmbBox_InspectM.AddItem months
        cmbBox_UnDockingM.AddItem months
        cmbBox_PrevUnDockingM.AddItem months
    Next months
    
    ' ���̑I������ǉ��i1��?31���j
    Dim Days As Integer
    For Days = 1 To 31
        cmbBox_ReceiptD.AddItem Days
        cmbBox_RepNoCreateD.AddItem Days
        cmbBox_InspectD.AddItem Days
        cmbBox_UnDockingD.AddItem Days
        cmbBox_PrevUnDockingD.AddItem Days
    Next Days
    
    ' ���̑�cmbBox
    ' cmbBox�\���̏�����
    ' ���ӁF "���s" �� "���s" �͓��̓~�X���N���₷��
    Call InitializeCmbBox(Me.cmbBox_Stat, GetNamedRange("��"))
    Call InitializeCmbBox(Me.cmbBox_KmsStaff, GetNamedRange("�S����"))
    Call InitializeCmbBox(Me.cmbBox_Location, GetNamedRange("���_"))
    Call InitializeCmbBox(Me.cmbBox_ShipType, GetNamedRange("�D�����"))
    Call InitializeCmbBox(Me.cmbBox_InspectType, GetNamedRange("�������"))
    Call InitializeCmbBox(Me.cmbBox_Clause, GetNamedRange("��"))
    Call InitializeCmbBox(Me.cmbBox_ConcurrentInspection, GetNamedRange("���s����"))
    Call InitializeCmbBox(Me.cmbBox_ShipYard, GetNamedRange("���D��"))
    Call InitializeCmbBox(Me.cmbBox_PrevInspectType, GetNamedRange("���s����"))
    Call InitializeCmbBox(Me.cmbBox_PropellerNum, GetNamedRange("����"))
    Call InitializeCmbBox(Me.cmbBox_PropellerMaterial, GetNamedRange("�ގ�"))
    Call InitializeCmbBox(Me.cmbBox_MarineAccidentReport, GetNamedRange("�C��񍐏�"))
    
End Sub


' �D�������L�^����t�H�[���ɍ��ڂ���͂���
Private Function InputFromInspectRec(RefID As String, FiscalY As String, RefNum As String)

    Dim load As Object
    Set load = SearchRec_RefID(RefID, FiscalY, RefNum)
    
    ' �������ʂ��e���ڂɓ���
    Me.cmbBox_Stat = load("stat")
    Me.cmbBox_KmsStaff = load("kmsStaff")
    Me.cmbBox_Location = load("location")
    Me.txtBox_ReceiptY = Year(load("receiptDate"))
    Me.cmbBox_ReceiptM = Month(load("receiptDate"))
    Me.cmbBox_ReceiptD = Day(load("receiptDate"))
    Me.txtBox_RepNoCreateY = Year(load("repNoCreateDate"))
    Me.cmbBox_RepNoCreateM = Month(load("repNoCreateDate"))
    Me.cmbBox_RepNoCreateD = Day(load("repNoCreateDate"))
    Me.txtBox_RepNo = load("repNo")
    Me.txtBox_ShipName = load("shipName")
    Me.cmbBox_ShipType = load("shipType")
    Me.txtBox_Owner = load("owner")
    Me.txtBox_CaptainName = load("captainName")
    Me.txtBox_Delegater = load("delegater")
    Me.txtBox_DelegateStaff = load("delegateStaff")
    Me.cmbBox_InspectType = load("inspectType")
    Me.cmbBox_Clause = load("clause")
    Me.cmbBox_ConcurrentInspection = load("concurrentInspection")
    Me.cmbBox_ShipYard = load("shipyard")
    Me.txtBox_InspectY = Year(load("inspectDate"))
    Me.cmbBox_InspectM = Month(load("inspectDate"))
    Me.cmbBox_InspectD = Day(load("inspectDate"))
    Me.txtBox_UndockingY = Year(load("unDocking"))
    Me.cmbBox_UnDockingM = Month(load("unDocking"))
    Me.cmbBox_UnDockingD = Day(load("unDocking"))
    Me.txtBox_PrevUnDockingY = Year(load("prevUndocking"))
    Me.cmbBox_PrevUnDockingM = Month(load("prevUndocking"))
    Me.cmbBox_PrevUnDockingD = Day(load("prevUndocking"))
    Me.cmbBox_PrevInspectType = load("prevInspection")
    Me.txtBox_PrevRepNo = load("prevRepNo")
    Me.txtBox_GrossT = load("grossT")
    Me.txtBox_Length = load("length")
    Me.txtBox_Breadth = load("breadth")
    Me.txtBox_Depth = load("depth")
    Me.txtBox_ShaftDia = load("shaftDia")
    Me.cmbBox_PropellerNum = load("propellerNum")
    Me.cmbBox_PropellerMaterial = load("propellerMaterial")
    Me.txtBox_PropellerDia = load("propellerDia")
    Me.txtBox_PropellerPitch = load("propellerPitch")
    Me.txtBox_AccidentDetail = load("accidentDetail")
    Me.cmbBox_MarineAccidentReport = load("marineAccidentReport")
    Me.txtBox_RepairAmount = Format(load("repairAmount"), "#,###")
    Me.txtBox_inspectDateOther = load("inspectDateOther")
    Me.txtBox_Remark = load("remark")
    
    Set InputFromInspectRec = load

End Function

' �t�H�[���̓��͓��e��D�������L�^�ɕۑ�����
Private Function SaveToInspectRec(RefID As String, FiscalY As String, RefNum As String)

    ' �e���ړ��e��D�������L�^�ɏ㏑��
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "stat", Me.cmbBox_Stat.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "kmsStaff", Me.cmbBox_KmsStaff.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "location", Me.cmbBox_Location.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "receiptDate", TrimAndHalfDigit(Me.txtBox_ReceiptY.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_ReceiptM.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_ReceiptD.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "repNoCreateDate", TrimAndHalfDigit(Me.txtBox_RepNoCreateY.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_RepNoCreateM.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_RepNoCreateD.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "repNo", Me.txtBox_RepNo.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "shipName", Me.txtBox_ShipName.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "shipType", Me.cmbBox_ShipType.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "owner", Me.txtBox_Owner.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "captainName", Me.txtBox_CaptainName.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "delegater", Me.txtBox_Delegater.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "delegateStaff", Me.txtBox_DelegateStaff.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "inspectType", Me.cmbBox_InspectType.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "clause", Me.cmbBox_Clause.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "concurrentInspection", Me.cmbBox_ConcurrentInspection.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "shipyard", Me.cmbBox_ShipYard.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "inspectDate", TrimAndHalfDigit(Me.txtBox_InspectY.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_InspectM.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_InspectD.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "unDocking", TrimAndHalfDigit(Me.txtBox_UndockingY.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_UnDockingM.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_UnDockingD.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "prevUndocking", TrimAndHalfDigit(Me.txtBox_PrevUnDockingY.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_PrevUnDockingM.Value) & "/" & TrimAndHalfDigit(Me.cmbBox_PrevUnDockingD.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "prevInspection", Me.cmbBox_PrevInspectType.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "prevRepNo", Me.txtBox_PrevRepNo.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "grossT", TrimAndHalfDigit(Me.txtBox_GrossT.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "length", TrimAndHalfDigit(Me.txtBox_Length.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "breadth", TrimAndHalfDigit(Me.txtBox_Breadth.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "depth", TrimAndHalfDigit(Me.txtBox_Depth.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "shaftDia", TrimAndHalfDigit(Me.txtBox_ShaftDia.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "propellerNum", Me.cmbBox_PropellerNum.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "propellerMaterial", Me.cmbBox_PropellerMaterial.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "propellerDia", TrimAndHalfDigit(Me.txtBox_PropellerDia.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "propellerPitch", TrimAndHalfDigit(Me.txtBox_PropellerPitch.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "accidentDetail", Me.txtBox_AccidentDetail.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "marineAccidentReport", Me.cmbBox_MarineAccidentReport.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "repairAmount", TrimAndHalfDigit(Me.txtBox_RepairAmount.Value))
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "inspectDateOther", Me.txtBox_inspectDateOther.Value)
    Call RewriteCaseItems(RefID, FiscalY, RefNum, "remark", Me.txtBox_Remark.Value)


End Function

' "�Č�����" �{�^������
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
    Set SearchInspectRec = InputFromInspectRec("", FiscalY, RefNum)
    ' Set SearchInspectRec = InputFromInspectRec(FiscalY_RefNum, "", "")
    
    If SearchInspectRec("repNoCreateDate") <> "" Or SearchInspectRec("repNoCreateDate") <> "" Then
        MsgBox "���s���������͊Ӓ菑�ԍ������ɔ��s����Ă��܂��B" & Chr(13) & "�ҏW�𑱂���ꍇ�͍ēx�m�F�̏�s���ĉ������B", , "�x��"
    Else
    End If

End Sub

' "����D���Č�����" �{�^������
Private Sub Button_SearchCaseSameShipName_Click()
    
    Dim ShipName As String
    ShipName = Me.txtBox_ShipName
    Dim SearchPrevCase As Variant
    SearchPrevCase = SearchRec_ShipName(ShipName)
    
    Dim PrevCase As Object
    ' MsgBox (SearchPrevCase(UBound(SearchPrevCase)))
    Set PrevCase = SearchRec_RefID(Val(SearchPrevCase), "", "")
    
    Me.cmbBox_ShipType = PrevCase("shipType")
    Me.txtBox_Owner = PrevCase("owner")
    Me.txtBox_CaptainName = PrevCase("captainName")
    Me.txtBox_Delegater = PrevCase("delegater")
    Me.cmbBox_Clause = PrevCase("clause")
    Me.txtBox_PrevUnDockingY = Year(PrevCase("unDocking"))
    Me.cmbBox_PrevUnDockingM = Month(PrevCase("unDocking"))
    Me.cmbBox_PrevUnDockingD = Day(PrevCase("unDocking"))
    Me.cmbBox_PrevInspectType = PrevCase("concurrentInspection")
    Me.txtBox_PrevRepNo = PrevCase("repNo")
    Me.txtBox_GrossT = PrevCase("grossT")
    Me.txtBox_Length = PrevCase("length")
    Me.txtBox_Breadth = PrevCase("breadth")
    Me.txtBox_Depth = PrevCase("depth")
    Me.txtBox_ShaftDia = PrevCase("shaftDia")
    Me.cmbBox_PropellerNum = PrevCase("propellerNum")
    Me.cmbBox_PropellerMaterial = PrevCase("propellerMaterial")
    Me.txtBox_PropellerDia = PrevCase("propellerDia")
    Me.txtBox_PropellerPitch = PrevCase("propellerPitch")
    
    ' �����܂�
    

End Sub

' "�Ӓ菑�ԍ����s" �{�^������
Private Sub Button_CreateRepNo_Click()

    Call Button_Save_Click
    
    Dim FiscalY As String, RefNum As String
    
    FiscalY = TrimAndHalfDigit(Me.txtBox_FiscalY)
    RefNum = TrimAndHalfDigit(Me.txtBox_RefNum)
    
    If FiscalY = "0" Or RefNum = "0" Then
        MsgBox "��t�N�x �������� ��tNo. �ɒl�����͂���Ă��܂���B", , "�G���["
        Exit Sub
    Else
    End If
    
    Unload RecordListEdit
    Call Create_RepNo.Button_Search_Click

End Sub

' "�V�K��t���" �{�^������
Private Sub Button_NewReception_Click()

    Call Button_Save_Click
    
    Unload RecordListEdit
    NewReception.Show

End Sub

' "��ʃ��Z�b�g" �{�^������
Private Sub Button_Reset_Click()

    Dim CheckResetForm As Integer
    CheckResetForm = MsgBox("��ʂ̓��͓��e�����Z�b�g���܂����H", vbQuestion + vbYesNo + vbDefaultButton2, "���Z�b�g�m�F")
    If CheckResetForm = vbYes Then
        Unload RecordListEdit
        RecordListEdit.Show
        MsgBox "��ʂ����Z�b�g���܂����B", , "����"
    Else
    End If

End Sub

' "�ύX��ۑ�" �{�^������
Private Sub Button_Save_Click()

    Dim CheckSaveForm As Integer
    CheckSaveForm = MsgBox("���͓��e��ۑ����܂����H", vbQuestion + vbYesNo + vbDefaultButton2, "�ۑ��m�F")
    If CheckSaveForm = vbYes Then
        Dim FiscalY As String, RefNum As String
        ' Dim FiscalY_RefNum As String
    
        FiscalY = TrimAndHalfDigit(Me.txtBox_FiscalY)
        RefNum = TrimAndHalfDigit(Me.txtBox_RefNum)
    
        If FiscalY = "0" Or RefNum = "0" Then
            MsgBox "��t�N�x �������� ��tNo. �ɒl�����͂���Ă��܂���B", , "�G���["
            Exit Sub
        Else
            Call SaveToInspectRec("", FiscalY, RefNum)
            MsgBox "�ۑ����܂����B", , "����"
        End If
        
    Else
    End If

End Sub

' "�ۑ����Ĉ��" �{�^������
Private Sub Button_SaveAndPrint_Click()
    Call Button_Save_Click

    Dim sheetName As String
    Dim ws As Worksheet
    sheetName = "test"
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Range("AY7").Value = Me.txtBox_FiscalY
    ws.Range("AZ7").Value = Me.txtBox_RefNum
    
    Unload RecordListEdit
    ws.PrintPreview
    
End Sub

' "����" �{�^������
Private Sub Button_Exit_Click()

    Call Button_Save_Click
    Unload RecordListEdit

End Sub
