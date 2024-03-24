Sub allLower()
    ' Делает все буквы в выделенном диапазоне строчными.
    For Each x In Range(Selection.Address)
        x.Value = LCase(x.Value)
    Next
End Sub

Sub allUpper()
    ' Делает все буквы в выделенном диапазоне прописными.
    For Each x In Range(Selection.Address)
        x.Value = UCase(x.Value)
    Next
End Sub

Sub convertDates_fromIntegers()
    ' Преобразует числа вида 44885 в даты вида 20.11.2022.
    For Each x In Range(Selection.Address)
        If x.Value <> "" And x.Value <> "None" Then x.Value = "'" & CDate(x.Value)
    Next
End Sub
