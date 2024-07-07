Attribute VB_Name = "SearchInspectRec"
Option Explicit

' �D�������L�^���̎�tNo.��N�x�A���ړ�����������v���V�[�W��

' �N�x&��t���őD�������L�^���������āA�l(�Z�����eor�A�h���X)��Ԃ�

Function SearchRec_RefID(RefID As String, FiscalY As String, RefNum As String, Optional AddressOrOther As String)

    ' �e�ϐ����̐錾
    Dim searchCategories As Object
    Set searchCategories = InspectRecCategories()
    
    Dim resultCategories As Object
    Set resultCategories = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "�D�������L�^"
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    Dim rangeRefID, rangeCategory, cellCategory As Range
    Dim matchedRow, matchedColumn As Long
    
    
    ' RefID(2024001)��������FiscalY(2024)&RefNum(001)�̓��͕ʂ̕���
    Dim searchValue As String
    If RefID = "" Then
        searchValue = Val(FiscalY & RefNum)
    Else
        searchValue = Val(RefID)
    End If
    

    ' A��(RefID)�ł̌����l�̌���
    Set rangeRefID = ws.Range("A9:A30000").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    If rangeRefID Is Nothing Then
        If RefID = "" Then
            SearchRec_RefID = "�G���[: " & FiscalY & "�N �������� No." & RefNum & " �̂ǂ��炩�̐��l�͑��݂��܂���B"
        Else
            SearchRec_RefID = "�G���[: " & RefID & " �Ƃ���ID�͑��݂��܂���B"
        End If
        Exit Function
    Else
        matchedRow = rangeRefID.Row
    End If
    
    
    
    ' �J�e�S����(�D��,���L��,��t�� ��)�ł̌���&�z��ɒǋL
    Dim key As Variant
    Set rangeCategory = ws.Range("B8:AP8")
    
    For Each key In searchCategories.keys
        Set cellCategory = rangeCategory.Find(What:=searchCategories(key), LookIn:=xlValues, LookAt:=xlPart)
        If cellCategory Is Nothing Then
            SearchRec_RefID = "�G���[: " & searchCategories(key) & " �Ƃ����J�e�S�����͑��݂��܂���B"
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
    
    
    
    ' ���ʂ�\������ (�f�o�b�O�p)
    ' For Each key In resultCategories.keys
    '    Debug.Print "Key: " & key & ", Result: " & resultCategories(key)
    ' Next key

    Set SearchRec_RefID = resultCategories

End Function


' ���ڂ��������ߎg�p����Ƃ��Ȃ�d���Ȃ�܂��B
' �J�e�S����(shipName(�D��),shipType(�D�����) ��)�Ɠ��e ��) �����,�ݕ��D ��) �őD�������L�^����������RefID��z��ŕԂ�
Function SearchRec_Category(Category As String, Content As String, Optional RefIDarray As Variant)
    
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "�D�������L�^"
    Set ws = ThisWorkbook.Sheets(sheetName)
        
    ' Function�̃I�v�V�����ɔz�񂪖����ꍇ�͑D�������L�^��A����͍ϑS�l�������Ώۂɂ���
    If IsMissing(RefIDarray) Then
        
        Dim cell As Range
        Dim tempList As Collection
        Set tempList = New Collection
        
        ' �͈͓��̋󔒂łȂ��Z���̒l���R���N�V�����ɒǉ�
        For Each cell In ws.Range("A9:A30000").Cells
            If cell.Value <> "" Then
                tempList.Add cell.Value
            End If
        Next cell
    
        ' �R���N�V��������z��֕ϊ�
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
            ' Debug.Print RefIDarray(rangeCount) & "�͈�v���܂���ł����B"
        End If
            
        ' If rangeRefID Is Nothing Then
            ' Debug.Print RefIDarray(rangeCount) & " �͎w��͈͂ɑ��݂��܂���B"
        ' Else
        '
        '     tmpColl.Add RefIDarray(rangeCount)
        ' End If
        
    Next rangeCount
    
    ' 1������v���Ȃ������ꍇ�̏���
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


' SearchRec_Category ���A���ڂ������ďd�����߁A�y�[�W�ŉ���(�ŐV)����1���ڂ̂݃}�b�`(�b�藘�p)
' �J�e�S����(shipName(�D��),shipType(�D�����) ��)�Ɠ��e ��) �����,�ݕ��D ��) �őD�������L�^����������RefID��z��ŕԂ�
Function SearchRec_Category2(Category As String, Content As String, Optional RefIDarray As Variant)
    
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "�D�������L�^"
    Set ws = ThisWorkbook.Sheets(sheetName)
        
    ' Function�̃I�v�V�����ɔz�񂪖����ꍇ�͑D�������L�^��A����͍ϑS�l�������Ώۂɂ���
    If IsMissing(RefIDarray) Then
        
        Dim cell As Range
        Dim tempList As Collection
        Set tempList = New Collection
        
        ' �͈͓��̋󔒂łȂ��Z���̒l���R���N�V�����ɒǉ�
        For Each cell In ws.Range("A9:A30000").Cells
            If cell.Value <> "" Then
                tempList.Add cell.Value
            End If
        Next cell
    
        ' �R���N�V��������z��֕ϊ�
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
            ' Debug.Print RefIDarray(rangeCount) & "�͈�v���܂���ł����B"
        End If
            
        ' If rangeRefID Is Nothing Then
            ' Debug.Print RefIDarray(rangeCount) & " �͎w��͈͂ɑ��݂��܂���B"
        ' Else
        '
        '     tmpColl.Add RefIDarray(rangeCount)
        ' End If
        
    Next rangeCount
    
    ' 1������v���Ȃ������ꍇ�̏���
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


' �D�������ŁA���߂�1���Ƀ}�b�`

Function SearchRec_ShipName(ShipName As String)

    ' �e�ϐ����̐錾
    Dim searchCategories As Object
    Set searchCategories = InspectRecCategories()
    
    Dim resultCategories As Object
    Set resultCategories = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "�D�������L�^"
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    Dim rangeShipName As Range
    Dim matchedRow As Long

    ' A��(RefID)�ł̌����l�̌���
    Set rangeShipName = ws.Range("E9:E30000").Find(What:=ShipName, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, After:=ws.Range("E30000"))
    If rangeShipName Is Nothing Then
        If ShipName = "" Then
            SearchRec_ShipName = "�G���[: " & ShipName & " �͑��݂��܂���B"
        Else
        End If
        Exit Function
    Else
        matchedRow = rangeShipName.Row
    End If
    
    ' ���ʂ�\������ (�f�o�b�O�p)
    ' For Each key In resultCategories.keys
    '    Debug.Print "Key: " & key & ", Result: " & resultCategories(key)
    ' Next key

    SearchRec_ShipName = ws.Cells(matchedRow, 1).Value

End Function
