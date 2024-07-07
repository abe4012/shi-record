Attribute VB_Name = "Functions"
Option Explicit

' 前後の半角,全角スペースを取り除き、数値を半角に変換
Function TrimAndHalfDigit(s As String) As String
    Dim result As String
    result = s
    
    ' 先頭全角スペース削除
    Do While Left(result, 1) = ChrW(&H3000)
        result = Mid(result, 2)
    Loop
    ' 末尾全角スペース削除
    Do While Right(result, 1) = ChrW(&H3000)
        result = Left(result, Len(result) - 1)
    Loop
    
    TrimAndHalfDigit = Val(Trim(StrConv(result, vbNarrow)))

End Function


' 入力された値が存在する年月日であるか検証して True or False で返す
Function CheckCorrectDate(inputDate As String)

    If IsDate(inputDate) Then
        CheckCorrectDate = True
    Else
        CheckCorrectDate = False
    End If
    
End Function


' 月(1～12),日(1～31)の連番を配列で返す
' 引数: "Month" or "Days" (String)
Function MonthsDaysArray(DateType As String)
    If DateType = "Month" Then
        Dim result(1 To 12) As Integer
        Dim i As Integer
        
        For i = 1 To 12
            result(i) = i
        Next i
    ElseIf DateType = "Days" Then
        Dim result(1 To 31) As Integer
        Dim i As Integer
        
        For i = 1 To 31
            result(i) = i
        Next i
    Else
    End If
    
    ' 配列を返す
    MonthDaysArray = result

End Function


' 入力された年度から現在の年度のヘッダーを取得する
Function SearchFiscalYheader(ByVal searchValue As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Rep.No.-header")
    
    Dim searchRange As Range
    Dim foundCell As Range
    
    ' 検索範囲を設定
    Set searchRange = ws.Range("A6:A305")
    
    ' A列で検索値を検索
    Set foundCell = searchRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    
    ' 検索値が見つかった場合、同じ行のB列の値を返す
    Dim NotFoundMsg As String
    NotFoundMsg = "今年度のヘッダーが見つかりませんでした。" & Chr(13) & "下部からシート「Rep.No.-header」を開き、今年度のヘッダーを追加してください。"
    If foundCell Is Nothing Then
        MsgBox NotFoundMsg, , "エラー"
        Exit Function
    ElseIf ws.Cells(foundCell.Row, "B").Value = 0 Then
        MsgBox NotFoundMsg, , "エラー"
        Exit Function
    Else
        SearchFiscalYheader = ws.Cells(foundCell.Row, "B").Value
    End If
    
    

End Function

' 指定した名前付きリストの値を配列で返す
Function GetNamedRange(namedRange As String) As Variant
    Dim rng As Range
    Dim cell As Range
    Dim coll As New Collection
    Dim arr() As Variant
    Dim i As Long

    ' 名前付きリストを取得
    Set rng = ThisWorkbook.Names(namedRange).RefersToRange

    ' 名前付きリストの全てのセルをループで回す
    For Each cell In rng
        ' セルの値が空白でない場合、その値をCollectionに追加
        If cell.Value <> "" Then
            coll.Add cell.Value
        End If
    Next cell

    ' Collectionの要素を配列に変換
    Dim namedRangeArray As Variant
    ReDim namedRangeArray(1 To coll.Count)
    For i = 1 To coll.Count
        namedRangeArray(i) = coll(i)
    Next i

    ' 配列を返す
    GetNamedRange = namedRangeArray
    
End Function


' 指定したヘッダーでの新規鑑定書番号を取得する
Function GetNewRepNo(header As String, Optional addNum As Integer) As Variant
    Dim searchTxt As String
    searchTxt = "鑑定書番号"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("船舶検査記録")
    
    Dim searchRange As Range
    Dim foundColumn As Range
    Dim cell As Range
    Dim maxValue As Long
    Dim tempValue As Long
    Dim tempStr As String
    
    ' B8:AP8 の範囲で searchTxt に一致する列を検索
    Set searchRange = ws.Range("B8:AP8").Find(What:=searchTxt, LookIn:=xlValues, LookAt:=xlWhole)
    
    
    If Not searchRange Is Nothing Then
        ' 合致した列の範囲で (header) で始まる文字を検索し、数値を取得
        For Each cell In ws.Range(ws.Cells(9, searchRange.Column), ws.Cells(30000, searchRange.Column))
            If Left(cell.Value, Len(header)) = header Then
                tempStr = Replace(cell.Value, header, "")
                If IsNumeric(tempStr) Then
                    tempValue = Val(tempStr)
                    If tempValue > maxValue Then maxValue = tempValue
                End If
            End If
        Next cell
        
        ' 最大値に addNum を加えた値を返す(0埋め4桁 String)
        GetNewRepNo = Format(maxValue + addNum, "0000")
    Else
        ' searchTxt に一致する列が見つからない場合
        GetNewRepNo = "Error: Column not found"
    End If

End Function



' 指定したRefID(RefNum,FiscalY)とカテゴリ名,内容で船舶検査記録を書き換え
Function RewriteCaseItems(RefID As String, FiscalY As String, RefNum As String, Category As String, Content As String)

    ' RefID(2024001)もしくはFiscalY(2024)&RefNum(001)の入力別の分岐
    Dim searchValue As String
    If RefID = "" Then
        searchValue = Val(FiscalY & RefNum)
    Else
        searchValue = Val(RefID)
    End If

    Dim SearchInspectRecAddr As Object
    Set SearchInspectRecAddr = SearchRec_RefID(searchValue, "", "", "Address")


    ' RefID に対応する Category 及び Content を用いて書き換え
    If Category <> "0" And Content <> "0" Then
        If SearchInspectRecAddr.Exists(Category) Then
            Dim celladdress As String
            celladdress = SearchInspectRecAddr(Category)
            ' DebugPoint (celladdress)
            Range(celladdress).Value = Content
            
            ' Debug.Print celladdress & "の内容が" & Content & "に更新されました。"
            RewriteCaseItems = "RewriteCaseItems: true"
        Else
            MsgBox Category & "は連想配列に存在しません。"
            Exit Function
        End If
    Else
        RewriteCaseItems = "エラー： 値が入力されていません。"
        ' debug.Print "RewriteCaseItems: エラー： 値が入力されていません。"
        Exit Function
    End If
    
End Function




