VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "MENU"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4125
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form読み込み時初期化処理
Private Sub UserForm_Initialize()

    Call ShipInspectRecord_Search
    Call ShipInspectRecord_ShowAll

End Sub

' 新規受付画面を開く
Private Sub Button_NewReception_Click()

    Unload Me
    NewReception.Show
    
End Sub

' 船舶検査記録表編集画面を開く
Private Sub Button_RecordListEdit_Click()

    Unload Me
    RecordListEdit.Show
    
End Sub

' 鑑定書番号発行画面を開く
Private Sub Button_CreateRepNo_Click()

    Unload Me
    Create_RepNo.Show
    
End Sub

' "閉じる"ボタン押下
Private Sub Button_Exit_Click()

    Unload Me

End Sub
