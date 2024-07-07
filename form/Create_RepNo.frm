VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Create_RepNo 
   Caption         =   "鑑定書番号発行画面"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   OleObjectBlob   =   "Create_RepNo.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Create_RepNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label_hakkonengappi_Click()

End Sub

' Form読み込み時初期化処理
Private Sub UserForm_Initialize()
    
    Call ShipInspectRecord_Search
    Call ShipInspectRecord_ShowAll

    ' 発行年月日の初期化
    ' 現在の年、月、日を取得
    Dim CurrentYear As Integer, CurrentMonth As Integer, CurrentDay As Integer
    CurrentYear = Year(Date)
    CurrentMonth = Month(Date)
    CurrentDay = Day(Date)
    ' txtBox_CurrentYに現在の年を設定
    txtBox_CurrentY.Value = CurrentYear
    
    ' cmbBox_CurrentMに月の選択肢を追加（1月?12月）
    Dim months As Integer
    For months = 1 To 12
        cmbBox_CurrentM.AddItem months
    Next months
    
    ' cmbBox_CurrentDに日の選択肢を追加（1日?31日）
    Dim Days As Integer
    For Days = 1 To 31
        cmbBox_CurrentD.AddItem Days
    Next Days
    
    ' ComboBoxの現在の月と日を選択
    cmbBox_CurrentM.Value = CurrentMonth
    cmbBox_CurrentD.Value = CurrentDay
    
    
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

' "案件検索"ボタン押下時処理
Sub Button_Search_Click()

    Dim FiscalY As String, RefNum As String
    ' Dim FiscalY_RefNum As String
    
    FiscalY = TrimAndHalfDigit(Me.txtBox_FiscalY)
    RefNum = TrimAndHalfDigit(Me.txtBox_RefNum)
    
    If FiscalY = "0" Or RefNum = "0" Then
        MsgBox "受付年度 もしくは 受付No. に値が入力されていません。", , "エラー"
        Exit Sub
    Else
    End If
    
    ' FiscalY_RefNum = FiscalY & RefNum
    
    Dim SearchInspectRec As Object
    Set SearchInspectRec = SearchRec_RefID("", FiscalY, RefNum)
    ' Set SearchInspectRec = SearchRec_RefID(FiscalY_RefNum, "", "")
    
    ' 検索結果をLabelに書き込み
    Label_FStat = SearchInspectRec("stat")
    Label_FFiscalY = SearchInspectRec("year")
    Label_FRefNum = SearchInspectRec("refNum")
    Label_FReceiptDate = SearchInspectRec("receiptDate")
    Label_FShipNMandType = SearchInspectRec("shipName") & "　　(" & SearchInspectRec("shipType") & ")"
    Label_FOwner = SearchInspectRec("owner")
    Label_FDelegateAndStaff = SearchInspectRec("delegater") & "　" & SearchInspectRec("delegateStaff")
    Label_FInspectType = SearchInspectRec("inspectType")
    Label_FClause = SearchInspectRec("clause")
    Label_FConcurrentInspection = SearchInspectRec("concurrentInspection")
    Label_FShipyard = SearchInspectRec("shipyard")
    Label_FInspectDate = SearchInspectRec("inspectDate")
    Label_FKmsStaff = SearchInspectRec("kmsStaff")
    Label_FRepNoCreateDate = SearchInspectRec("repNoCreateDate")
    Label_FRepNo = SearchInspectRec("repNo")
    
    If SearchInspectRec("repNoCreateDate") <> "" Or SearchInspectRec("repNoCreateDate") <> "" Then
        MsgBox "発行日もしくは鑑定書番号が既に発行されています。" & Chr(13) & "このまま番号を発行する場合は、再確認の上、行ってください。", , "警告"
    Else
    End If
    
End Sub

' "鑑定書番号発行"ボタン押下時処理
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
        MsgBox "受付年度 もしくは 受付No. に値が入力されていません。", , "エラー"
        Exit Sub
    ElseIf FiscalY <> TrimAndHalfDigit(Me.Label_FFiscalY) Or RefNum <> TrimAndHalfDigit(Me.Label_FRefNum) Then
        MsgBox "入力されている 受付年度, 受付No. が表示されている案件情報のものと異なります。" & Chr(13) & "一度「案件検索」ボタンを押して案件情報が正しいか確認してください。", , "エラー"
        Exit Sub
    ElseIf Not CheckCorrectDate(CreateDate) Then
        MsgBox "指定された発行年月日は日付として存在しません。", , "エラー"
        Exit Sub
    ElseIf chk_RepNoManual.Value = True Then
        If txtBox_RepNoManual.Value <> "" Then
            RepNoManual = True
        Else
            MsgBox "鑑定書番号を入力してください。" & Chr(13) & "(自動入力する場合は「鑑定書番号手入力」の" & Chr(13) & "チェックを外してください。)", , "エラー"
            Exit Sub
        End If
    ElseIf Me.cmbBox_Header = "" Then
        MsgBox "ヘッダーを選択してください。", , "エラー"
        Exit Sub
    Else
    End If
    
    ' 鑑定書番号取得
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
    CheckCreateRepNo = MsgBox("以下の内容で鑑定書番号を発行します。" & Chr(13) & "確認の上、問題無ければ「はい(Yes)」を押して下さい。" & Chr(13) & Chr(13) & "年度：" & SearchInspectRec("year") & Chr(13) & "受付No.：" & SearchInspectRec("refNum") & Chr(13) & "受付日：" & SearchInspectRec("receiptDate") & Chr(13) & "船名(種)：" & SearchInspectRec("shipName") & " (" & SearchInspectRec("shipType") & ")" & Chr(13) & "担当者：" & SearchInspectRec("kmsStaff") & Chr(13) & Chr(13) & "発行日：" & CreateDate & Chr(13) & "鑑定書番号：" & repNo, vbQuestion + vbYesNo + vbDefaultButton2, "鑑定書番号発行前確認")
    If CheckCreateRepNo = vbYes Then
    
        Dim CreateRepNo As String, CreateRepNoCreateDate As String
        ' Debug.Print RewriteCaseItems("", FiscalY, RefNum, "repNo", repNo)
        ' Debug.Print RewriteCaseItems("", FiscalY, RefNum, "repNoCreateDate", CreateDate)
        CreateRepNo = RewriteCaseItems("", FiscalY, RefNum, "repNo", repNo)
        CreateRepNoCreateDate = RewriteCaseItems("", FiscalY, RefNum, "repNoCreateDate", CreateDate)
        
        Dim CheckExitForm As Integer
        CheckExitForm = MsgBox("鑑定書を発行しました。" & Chr(13) & Chr(13) & "鑑定書発行画面を閉じますか？", vbQuestion + vbYesNo + vbDefaultButton2, "フォーム終了確認")
        If CheckExitForm = vbYes Then
            Unload Me
        Else
        End If
    
    Else
    End If

End Sub


' "閉じる"ボタン押下
Private Sub Button_Exit_Click()
    Unload Me
End Sub

