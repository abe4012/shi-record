Attribute VB_Name = "ShowForm"
Option Explicit

' MENU を開く
Sub Show_Menu()

    Menu.Show
    
End Sub

' 新規受付画面を開く
Sub Show_NewReception()

    NewReception.Show
    
End Sub


' 鑑定書番号発行画面を開く
Sub Show_Create_RepNo()

    Create_RepNo.Show
    
End Sub


' 船舶検査記録編集画面 を開く
Sub Show_RecordListEdit()

    RecordListEdit.Show
    
End Sub


' 最新の鑑定書番号表示画面を開く
Sub Show_ShowLatestRepNo()

    ShowLatestRepNo.Show

End Sub
