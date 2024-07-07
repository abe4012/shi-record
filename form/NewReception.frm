VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewReception 
   Caption         =   "�V�K��t"
   ClientHeight    =   2355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3045
   OleObjectBlob   =   "NewReception.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "NewReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form�ǂݍ��ݎ�����������
Private Sub UserForm_Initialize()

    Call ShipInspectRecord_Search
    Call ShipInspectRecord_ShowAll
    
    ' ���݂̔N�A���A�����擾
    Dim CurrentYear As Integer, CurrentMonth As Integer, CurrentDay As Integer
    CurrentYear = Year(Date)
    CurrentMonth = Month(Date)
    CurrentDay = Day(Date)
    
    ' �N�x�ϊ�
    Dim currentFiscalY As Integer
    If CurrentMonth <= 3 Then
        currentFiscalY = CurrentYear - 1
    Else
        currentFiscalY = CurrentYear
    End If
    
    txtBox_FiscalY.Value = currentFiscalY

End Sub

' "��t�� " �{�^��������
Private Sub Button_Create_Click()
    
    
    Dim FiscalY As String, RefNum As String
    FiscalY = Me.txtBox_FiscalY
    Dim CheckCreateForm As String
    CheckCreateForm = MsgBox(FiscalY & "�N�x�Ŏ�t����V�K���s���܂����H", vbQuestion + vbYesNo + vbDefaultButton2, "���s�m�F")
    If CheckCreateForm = vbYes Then
        
        Dim ws As Worksheet
        Dim sheetName As String
        sheetName = "�D�������L�^"
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        Dim categoryRange As Range
        Set categoryRange = ws.Range("B8:AP8")
        
        Dim yearCell As Range
        Set yearCell = categoryRange.Find(What:=InspectRecCategories("year"), LookIn:=xlValues, LookAt:=xlWhole)
        
        Dim yearRange As Range
        ' 9�s�ڂ���30000�s�ڂ܂ł�\�͈̔͂Ƃ��Đݒ�
        Set yearRange = ws.Range(ws.Cells(9, yearCell.Column), ws.Cells(30000, yearCell.Column))
        
        Dim foundCell As Range
        Set foundCell = yearRange.Find(What:=FiscalY, After:=yearRange.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        
        If Not foundCell Is Nothing Then
            ' ��v����l�����������ꍇ�A����1�s���̃Z���̃A�h���X��Ԃ�
            foundCell.Offset(1, 0).Value = FiscalY
        
            Dim PrevRefNum As Integer
            PrevRefNum = foundCell.Offset(0, 1).Value
            RefNum = PrevRefNum + 1
            
            foundCell.Offset(1, 1).Value = RefNum
            Debug.Print "If"
        
        Else
            ' ��v����l��������Ȃ��ꍇ�A�͈͓��̍ŉ����̋󔒂ł͂Ȃ��Z����T��
            Dim lastNonEmptyCell As Range
            Dim cell As Range
            For Each cell In yearRange
                If cell.Value <> "" Then Set lastNonEmptyCell = cell
            Next cell
            If Not lastNonEmptyCell Is Nothing Then
                ' �ł����̋󔒂ł͂Ȃ��Z����10�s���ɔN�x�Ǝ�t�����L������
                lastNonEmptyCell.Offset(10, 0).Value = FiscalY
                lastNonEmptyCell.Offset(10, 1).Value = 1
            Else
                Debug.Print "No non-empty cells found"
            End If
            Debug.Print "Else"
        End If
        
        MsgBox "��t���𔭍s���܂����B" & Chr(13) & Chr(13) & "��t�N�x: " & FiscalY & Chr(13) & "��t��: " & RefNum & Chr(13) & Chr(13) & "��L�̔N�x�Ǝ�t�����A���̃t�H�[����" & Chr(13) & "����ɓ��͂��Ă��������B", , "��t����"
        
        Unload NewReception
        RecordListEdit.Show
    Else
    End If
    
End Sub

' "����" �{�^��������
Private Sub Button_Exit_Click()

    Unload NewReception

End Sub
