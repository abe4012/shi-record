VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewReception 
   Caption         =   "新規受付"
   ClientHeight    =   2355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3045
   OleObjectBlob   =   "NewReception.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "NewReception"
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
    
    ' 年度変換
    Dim currentFiscalY As Integer
    If CurrentMonth <= 3 Then
        currentFiscalY = CurrentYear - 1
    Else
        currentFiscalY = CurrentYear
    End If
    
    txtBox_FiscalY.Value = currentFiscalY

End Sub

' "受付№ " ボタン押下時
Private Sub Button_Create_Click()
    
    
    Dim FiscalY As String, RefNum As String
    FiscalY = Me.txtBox_FiscalY
    Dim CheckCreateForm As String
    CheckCreateForm = MsgBox(FiscalY & "年度で受付№を新規発行しますか？", vbQuestion + vbYesNo + vbDefaultButton2, "発行確認")
    If CheckCreateForm = vbYes Then
        
        Dim ws As Worksheet
        Dim sheetName As String
        sheetName = "船舶検査記録"
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        Dim categoryRange As Range
        Set categoryRange = ws.Range("B8:AP8")
        
        Dim yearCell As Range
        Set yearCell = categoryRange.Find(What:=InspectRecCategories("year"), LookIn:=xlValues, LookAt:=xlWhole)
        
        Dim yearRange As Range
        ' 9行目から30000行目までを表の範囲として設定
        Set yearRange = ws.Range(ws.Cells(9, yearCell.Column), ws.Cells(30000, yearCell.Column))
        
        Dim foundCell As Range
        Set foundCell = yearRange.Find(What:=FiscalY, After:=yearRange.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        
        If Not foundCell Is Nothing Then
            ' 一致する値が見つかった場合、その1行下のセルのアドレスを返す
            foundCell.Offset(1, 0).Value = FiscalY
        
            Dim PrevRefNum As Integer
            PrevRefNum = foundCell.Offset(0, 1).Value
            RefNum = PrevRefNum + 1
            
            foundCell.Offset(1, 1).Value = RefNum
            Debug.Print "If"
        
        Else
            ' 一致する値が見つからない場合、範囲内の最下部の空白ではないセルを探す
            Dim lastNonEmptyCell As Range
            Dim cell As Range
            For Each cell In yearRange
                If cell.Value <> "" Then Set lastNonEmptyCell = cell
            Next cell
            If Not lastNonEmptyCell Is Nothing Then
                ' 最も下の空白ではないセルの10行下に年度と受付№を記入する
                lastNonEmptyCell.Offset(10, 0).Value = FiscalY
                lastNonEmptyCell.Offset(10, 1).Value = 1
            Else
                Debug.Print "No non-empty cells found"
            End If
            Debug.Print "Else"
        End If
        
        MsgBox "受付№を発行しました。" & Chr(13) & Chr(13) & "受付年度: " & FiscalY & Chr(13) & "受付№: " & RefNum & Chr(13) & Chr(13) & "上記の年度と受付№を、次のフォームの" & Chr(13) & "左上に入力してください。", , "受付完了"
        
        Unload NewReception
        RecordListEdit.Show
    Else
    End If
    
End Sub

' "閉じる" ボタン押下時
Private Sub Button_Exit_Click()

    Unload NewReception

End Sub
