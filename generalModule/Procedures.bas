Attribute VB_Name = "Procedures"
Option Explicit

' �w�肵��ComboBox�̒l���A�w�肵��Array�ŏ���������
Sub InitializeCmbBox(cmbBoxName As ComboBox, contentArray As Variant)

    Dim i As Long

        With cmbBoxName
            .Clear ' �����̍��ڂ��폜
            For i = LBound(contentArray) To UBound(contentArray)
                .AddItem contentArray(i) ' ���ڂ�ǉ�
            Next i
        End With
End Sub
