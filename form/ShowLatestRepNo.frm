VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShowLatestRepNo 
   Caption         =   "�ŐV�̊Ӓ菑�ԍ�"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4110
   OleObjectBlob   =   "ShowLatestRepNo.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ShowLatestRepNo"
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


Private Sub Button_Search_Click()

    If Me.cmbBox_Header = "" Then
        MsgBox "�w�b�_�[��I�����Ă��������B", , "�G���["
        Exit Sub
    End If
    
    Label_LatestRepNo = cmbBox_Header & GetNewRepNo(cmbBox_Header)
    
End Sub


Private Sub Button_Exit_Click()

    Unload ShowLatestRepNo

End Sub

