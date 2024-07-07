Attribute VB_Name = "SearchInspectRec"
Option Explicit

' 船舶検査記録内の受付No.や年度、項目等を検索するプロシージャ

' 年度&受付№で船舶検査記録を検索して、値(セル内容orアドレス)を返す

Function SearchRec_RefID(RefID As String, FiscalY As String, RefNum As String, Optional AddressOrOther As String)

    ' 各変数等の宣言
    Dim searchCategories As Object
    Set searchCategories = InspectRecCategories()
    
    Dim resultCategories As Object
    Set resultCategories = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "船舶検査記録"
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    Dim rangeRefID, rangeCategory, cellCategory As Range
    Dim matchedRow, matchedColumn As Long
    
    
    ' RefID(2024001)もしくはFiscalY(2024)&RefNum(001)の入力別の分岐
    Dim searchValue As String
    If RefID = "" Then
        searchValue = Val(FiscalY & RefNum)
    Else
        searchValue = Val(RefID)
    End If
    

    ' A列(RefID)での検索値の検索
    Set rangeRefID = ws.Range("A9:A30000").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    If rangeRefID Is Nothing Then
        If RefID = "" Then
            SearchRec_RefID = "エラー: " & FiscalY & "年 もしくは No." & RefNum & " のどちらかの数値は存在しません。"
        Else
            SearchRec_RefID = "エラー: " & RefID & " というIDは存在しません。"
        End If
        Exit Function
    Else
        matchedRow = rangeRefID.Row
    End If
    
    
    
    ' カテゴリ名(船名,所有者,受付日 等)での検索&配列に追記
    Dim key As Variant
    Set rangeCategory = ws.Range("B8:AP8")
    
    For Each key In searchCategories.keys
        Set cellCategory = rangeCategory.Find(What:=searchCategories(key), LookIn:=xlValues, LookAt:=xlPart)
        If cellCategory Is Nothing Then
            SearchRec_RefID = "エラー: " & searchCategories(key) & " というカテゴリ名は存在しません。"
            Exit Function
        Else
            matchedColumn = cellCategory.Column
            If AddressOrOther = "Address" Then
                resultCategories.Add key, sheetName & "!" & ws.Cells(matchedRow, matchedColumn).Address
            Else
                resultCategories.Add key, ws.Cells(matchedRow, matchedColumn).Value
            End If
        End If
    Next key
    
    
    
    ' 結果を表示する (デバッグ用)
    ' For Each key In resultCategories.keys
    '    Debug.Print "Key: " & key & ", Result: " & resultCategories(key)
    ' Next key

    Set SearchRec_RefID = resultCategories

End Function


' 項目が多いため使用するとかなり重くなります。
' カテゴリ名(shipName(船名),shipType(船舶種類) 等)と内容 例) 検定丸,貨物船 等) で船舶検査記録を検索してRefIDを配列で返す
Function SearchRec_Category(Category As String, Content As String, Optional RefIDarray As Variant)
    
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "船舶検査記録"
    Set ws = ThisWorkbook.Sheets(sheetName)
        
    ' Functionのオプションに配列が無い場合は船舶検査記録のA列入力済全値を検索対象にする
    If IsMissing(RefIDarray) Then
        
        Dim cell As Range
        Dim tempList As Collection
        Set tempList = New Collection
        
        ' 範囲内の空白でないセルの値をコレクションに追加
        For Each cell In ws.Range("A9:A30000").Cells
            If cell.Value <> "" Then
                tempList.Add cell.Value
            End If
        Next cell
    
        ' コレクションから配列へ変換
        ReDim RefIDarray(1 To tempList.Count)
    
        Dim RefIDarrayCount As Integer
        For RefIDarrayCount = 1 To tempList.Count
            RefIDarray(RefIDarrayCount) = tempList.Item(RefIDarrayCount)
        Next RefIDarrayCount
    
    Else
    End If


    Dim findColumn As Range
    Set findColumn = ws.Range("B8:AP8").Find(What:=InspectRecCategories(Category), LookIn:=xlValues, LookAt:=xlWhole)

    Dim rangeRefID As Range
    Dim tmpColl As New Collection
    Dim result As Variant
    Dim rangeCount As Long
    
    For rangeCount = LBound(RefIDarray) To UBound(RefIDarray)
        Set rangeRefID = ws.Range("A9:A30000").Find(What:=RefIDarray(rangeCount), LookIn:=xlValues, LookAt:=xlWhole)
        If ws.Cells(rangeRefID.Row, findColumn.Column).Value = Content Then
            tmpColl.Add RefIDarray(rangeCount)
        Else
            ' Debug.Print RefIDarray(rangeCount) & "は一致しませんでした。"
        End If
            
        ' If rangeRefID Is Nothing Then
            ' Debug.Print RefIDarray(rangeCount) & " は指定範囲に存在しません。"
        ' Else
        '
        '     tmpColl.Add RefIDarray(rangeCount)
        ' End If
        
    Next rangeCount
    
    ' 1件も一致しなかった場合の処理
    If tmpColl.Count = "0" Then
        Dim emptyArray As Variant
        emptyArray = Array()
        SearchRec_Category = emptyArray
        Exit Function
    Else
    End If
    
    ReDim result(1 To tmpColl.Count)
    Dim resultCount As Long
    For resultCount = 1 To tmpColl.Count
        result(resultCount) = tmpColl(resultCount)
    Next resultCount
    
    SearchRec_Category = result

End Function


' SearchRec_Category が、項目が多くて重いため、ページ最下部(最新)から1項目のみマッチ(暫定利用)
' カテゴリ名(shipName(船名),shipType(船舶種類) 等)と内容 例) 検定丸,貨物船 等) で船舶検査記録を検索してRefIDを配列で返す
Function SearchRec_Category2(Category As String, Content As String, Optional RefIDarray As Variant)
    
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "船舶検査記録"
    Set ws = ThisWorkbook.Sheets(sheetName)
        
    ' Functionのオプションに配列が無い場合は船舶検査記録のA列入力済全値を検索対象にする
    If IsMissing(RefIDarray) Then
        
        Dim cell As Range
        Dim tempList As Collection
        Set tempList = New Collection
        
        ' 範囲内の空白でないセルの値をコレクションに追加
        For Each cell In ws.Range("A9:A30000").Cells
            If cell.Value <> "" Then
                tempList.Add cell.Value
            End If
        Next cell
    
        ' コレクションから配列へ変換
        ReDim RefIDarray(1 To tempList.Count)
    
        Dim RefIDarrayCount As Integer
        For RefIDarrayCount = 1 To tempList.Count
            RefIDarray(RefIDarrayCount) = tempList.Item(RefIDarrayCount)
        Next RefIDarrayCount
    
    Else
    End If


    Dim findColumn As Range
    Set findColumn = ws.Range("B8:AP8").Find(What:=InspectRecCategories(Category), LookIn:=xlValues, LookAt:=xlWhole)

    Dim rangeRefID As Range
    Dim tmpColl As New Collection
    Dim result As Variant
    Dim rangeCount As Long
    
    For rangeCount = LBound(RefIDarray) To UBound(RefIDarray)
        ' MsgBox rangeCount
        Set rangeRefID = ws.Range("A9:A30000").Find(What:=RefIDarray(rangeCount), LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, After:=ws.Range("A30000"))
        If ws.Cells(rangeRefID.Row, findColumn.Column).Value = Content Then
            tmpColl.Add RefIDarray(rangeCount)
            Exit For
        Else
            ' Debug.Print RefIDarray(rangeCount) & "は一致しませんでした。"
        End If
            
        ' If rangeRefID Is Nothing Then
            ' Debug.Print RefIDarray(rangeCount) & " は指定範囲に存在しません。"
        ' Else
        '
        '     tmpColl.Add RefIDarray(rangeCount)
        ' End If
        
    Next rangeCount
    
    ' 1件も一致しなかった場合の処理
    If tmpColl.Count = "0" Then
        Dim emptyArray As Variant
        emptyArray = Array()
        SearchRec_Category2 = emptyArray
        Exit Function
    Else
    End If
    
    ReDim result(1 To tmpColl.Count)
    Dim resultCount As Long
    For resultCount = 1 To tmpColl.Count
        result(resultCount) = tmpColl(resultCount)
    Next resultCount
    
    SearchRec_Category2 = result

End Function


' 船名検索で、直近の1件にマッチ

Function SearchRec_ShipName(ShipName As String)

    ' 各変数等の宣言
    Dim searchCategories As Object
    Set searchCategories = InspectRecCategories()
    
    Dim resultCategories As Object
    Set resultCategories = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "船舶検査記録"
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    Dim rangeShipName As Range
    Dim matchedRow As Long

    ' A列(RefID)での検索値の検索
    Set rangeShipName = ws.Range("E9:E30000").Find(What:=ShipName, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, After:=ws.Range("E30000"))
    If rangeShipName Is Nothing Then
        If ShipName = "" Then
            SearchRec_ShipName = "エラー: " & ShipName & " は存在しません。"
        Else
        End If
        Exit Function
    Else
        matchedRow = rangeShipName.Row
    End If
    
    ' 結果を表示する (デバッグ用)
    ' For Each key In resultCategories.keys
    '    Debug.Print "Key: " & key & ", Result: " & resultCategories(key)
    ' Next key

    SearchRec_ShipName = ws.Cells(matchedRow, 1).Value

End Function
