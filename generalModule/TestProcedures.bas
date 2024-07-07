Attribute VB_Name = "TestProcedures"
Option Explicit

Sub CheckDate()

    Dim inputDate As String
    
    inputDate = "202/01/01"

If CheckCorrectDate(inputDate) Then
    MsgBox "日付です"
Else
    MsgBox "日付でない"
End If

End Sub



Sub getDate()
    Dim inputDate As String
    Dim yearPart As Integer, monthPart As Integer, dayPart As Integer
    
    inputDate = "2024/03/27"
    
    MsgBox "Year:" & Year(inputDate)
    MsgBox "Month:" & Month(inputDate)
    MsgBox "Day:" & Day(inputDate)
    
End Sub


Sub testSearchRec()

    Dim search1 As Variant, result As Variant
    search1 = SearchRec_Category("year", "2024")
    result = SearchRec_Category("delegater", "デリゲーター", search1)
    
    Dim i As Long
    For i = LBound(result) To UBound(result)
        Debug.Print result(i)
    Next i

End Sub


Sub testNamedRange()
    Dim test As Variant
    test = GetNamedRange("状況")
    Dim i As Integer
    For i = 1 To 6
        MsgBox test(i)
    Next i
End Sub


Sub test2()
    MsgBox InspectRecCategories("year")
End Sub
