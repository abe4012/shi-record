VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "MENU"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4125
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form�ǂݍ��ݎ�����������
Private Sub UserForm_Initialize()

    Call ShipInspectRecord_Search
    Call ShipInspectRecord_ShowAll

End Sub

' �V�K��t��ʂ��J��
Private Sub Button_NewReception_Click()

    Unload Me
    NewReception.Show
    
End Sub

' �D�������L�^�\�ҏW��ʂ��J��
Private Sub Button_RecordListEdit_Click()

    Unload Me
    RecordListEdit.Show
    
End Sub

' �Ӓ菑�ԍ����s��ʂ��J��
Private Sub Button_CreateRepNo_Click()

    Unload Me
    Create_RepNo.Show
    
End Sub

' "����"�{�^������
Private Sub Button_Exit_Click()

    Unload Me

End Sub
