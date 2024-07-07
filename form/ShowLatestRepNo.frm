VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShowLatestRepNo 
   Caption         =   "最新の鑑定書番号"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4110
   OleObjectBlob   =   "ShowLatestRepNo.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ShowLatestRepNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form読み込み時初期化処理
Private Sub UserForm_Initialize()

    Call ShipInspectRecord_Search
    Call ShipInspectRecord_ShowAll

    ' 現在の年、月、日を取得
    Dim CurrentYear As Integer, CurrentMonth As Integer, CurrentDay As Integer
    CurrentYear = Year(Date)
    CurrentMonth = Month(Date)
    CurrentDay = Day(Date)
    
    ' ヘッダーリスト表示の初期化
    Call InitializeCmbBox(Me.cmbBox_Header, GetNamedRange("HeaderList"))
    
    ' 年度変換
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
        MsgBox "ヘッダーを選択してください。", , "エラー"
        Exit Sub
    End If
    
    Label_LatestRepNo = cmbBox_Header & GetNewRepNo(cmbBox_Header)
    
End Sub


Private Sub Button_Exit_Click()

    Unload ShowLatestRepNo

End Sub

