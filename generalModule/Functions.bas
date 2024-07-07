Attribute VB_Name = "Functions"
Option Explicit

' �O��̔��p,�S�p�X�y�[�X����菜���A���l�𔼊p�ɕϊ�
Function TrimAndHalfDigit(s As String) As String
    Dim result As String
    result = s
    
    ' �擪�S�p�X�y�[�X�폜
    Do While Left(result, 1) = ChrW(&H3000)
        result = Mid(result, 2)
    Loop
    ' �����S�p�X�y�[�X�폜
    Do While Right(result, 1) = ChrW(&H3000)
        result = Left(result, Len(result) - 1)
    Loop
    
    TrimAndHalfDigit = Val(Trim(StrConv(result, vbNarrow)))

End Function


' ���͂��ꂽ�l�����݂���N�����ł��邩���؂��� True or False �ŕԂ�
Function CheckCorrectDate(inputDate As String)

    If IsDate(inputDate) Then
        CheckCorrectDate = True
    Else
        CheckCorrectDate = False
    End If
    
End Function


' ��(1�`12),��(1�`31)�̘A�Ԃ�z��ŕԂ�
' ����: "Month" or "Days" (String)
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
    
    ' �z���Ԃ�
    MonthDaysArray = result

End Function


' ���͂��ꂽ�N�x���猻�݂̔N�x�̃w�b�_�[���擾����
Function SearchFiscalYheader(ByVal searchValue As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Rep.No.-header")
    
    Dim searchRange As Range
    Dim foundCell As Range
    
    ' �����͈͂�ݒ�
    Set searchRange = ws.Range("A6:A305")
    
    ' A��Ō����l������
    Set foundCell = searchRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    
    ' �����l�����������ꍇ�A�����s��B��̒l��Ԃ�
    Dim NotFoundMsg As String
    NotFoundMsg = "���N�x�̃w�b�_�[��������܂���ł����B" & Chr(13) & "��������V�[�g�uRep.No.-header�v���J���A���N�x�̃w�b�_�[��ǉ����Ă��������B"
    If foundCell Is Nothing Then
        MsgBox NotFoundMsg, , "�G���["
        Exit Function
    ElseIf ws.Cells(foundCell.Row, "B").Value = 0 Then
        MsgBox NotFoundMsg, , "�G���["
        Exit Function
    Else
        SearchFiscalYheader = ws.Cells(foundCell.Row, "B").Value
    End If
    
    

End Function

' �w�肵�����O�t�����X�g�̒l��z��ŕԂ�
Function GetNamedRange(namedRange As String) As Variant
    Dim rng As Range
    Dim cell As Range
    Dim coll As New Collection
    Dim arr() As Variant
    Dim i As Long

    ' ���O�t�����X�g���擾
    Set rng = ThisWorkbook.Names(namedRange).RefersToRange

    ' ���O�t�����X�g�̑S�ẴZ�������[�v�ŉ�
    For Each cell In rng
        ' �Z���̒l���󔒂łȂ��ꍇ�A���̒l��Collection�ɒǉ�
        If cell.Value <> "" Then
            coll.Add cell.Value
        End If
    Next cell

    ' Collection�̗v�f��z��ɕϊ�
    Dim namedRangeArray As Variant
    ReDim namedRangeArray(1 To coll.Count)
    For i = 1 To coll.Count
        namedRangeArray(i) = coll(i)
    Next i

    ' �z���Ԃ�
    GetNamedRange = namedRangeArray
    
End Function


' �w�肵���w�b�_�[�ł̐V�K�Ӓ菑�ԍ����擾����
Function GetNewRepNo(header As String, Optional addNum As Integer) As Variant
    Dim searchTxt As String
    searchTxt = "�Ӓ菑�ԍ�"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("�D�������L�^")
    
    Dim searchRange As Range
    Dim foundColumn As Range
    Dim cell As Range
    Dim maxValue As Long
    Dim tempValue As Long
    Dim tempStr As String
    
    ' B8:AP8 �͈̔͂� searchTxt �Ɉ�v����������
    Set searchRange = ws.Range("B8:AP8").Find(What:=searchTxt, LookIn:=xlValues, LookAt:=xlWhole)
    
    
    If Not searchRange Is Nothing Then
        ' ���v������͈̔͂� (header) �Ŏn�܂镶�����������A���l���擾
        For Each cell In ws.Range(ws.Cells(9, searchRange.Column), ws.Cells(30000, searchRange.Column))
            If Left(cell.Value, Len(header)) = header Then
                tempStr = Replace(cell.Value, header, "")
                If IsNumeric(tempStr) Then
                    tempValue = Val(tempStr)
                    If tempValue > maxValue Then maxValue = tempValue
                End If
            End If
        Next cell
        
        ' �ő�l�� addNum ���������l��Ԃ�(0����4�� String)
        GetNewRepNo = Format(maxValue + addNum, "0000")
    Else
        ' searchTxt �Ɉ�v����񂪌�����Ȃ��ꍇ
        GetNewRepNo = "Error: Column not found"
    End If

End Function



' �w�肵��RefID(RefNum,FiscalY)�ƃJ�e�S����,���e�őD�������L�^����������
Function RewriteCaseItems(RefID As String, FiscalY As String, RefNum As String, Category As String, Content As String)

    ' RefID(2024001)��������FiscalY(2024)&RefNum(001)�̓��͕ʂ̕���
    Dim searchValue As String
    If RefID = "" Then
        searchValue = Val(FiscalY & RefNum)
    Else
        searchValue = Val(RefID)
    End If

    Dim SearchInspectRecAddr As Object
    Set SearchInspectRecAddr = SearchRec_RefID(searchValue, "", "", "Address")


    ' RefID �ɑΉ����� Category �y�� Content ��p���ď�������
    If Category <> "0" And Content <> "0" Then
        If SearchInspectRecAddr.Exists(Category) Then
            Dim celladdress As String
            celladdress = SearchInspectRecAddr(Category)
            ' DebugPoint (celladdress)
            Range(celladdress).Value = Content
            
            ' Debug.Print celladdress & "�̓��e��" & Content & "�ɍX�V����܂����B"
            RewriteCaseItems = "RewriteCaseItems: true"
        Else
            MsgBox Category & "�͘A�z�z��ɑ��݂��܂���B"
            Exit Function
        End If
    Else
        RewriteCaseItems = "�G���[�F �l�����͂���Ă��܂���B"
        ' debug.Print "RewriteCaseItems: �G���[�F �l�����͂���Ă��܂���B"
        Exit Function
    End If
    
End Function




