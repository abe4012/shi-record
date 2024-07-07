Attribute VB_Name = "Procedures"
Option Explicit

' 指定したComboBoxの値を、指定したArrayで初期化する
Sub InitializeCmbBox(cmbBoxName As ComboBox, contentArray As Variant)

    Dim i As Long

        With cmbBoxName
            .Clear ' 既存の項目を削除
            For i = LBound(contentArray) To UBound(contentArray)
                .AddItem contentArray(i) ' 項目を追加
            Next i
        End With
End Sub
