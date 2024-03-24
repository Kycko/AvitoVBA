Sub eveningCC()
    ' eveningCC Макрос
    ' Обработка вечерних отчётов КЦ.
    Range("$A$1:$ZZ$1048576").AutoFilter Field:=14, Criteria1:=Array( _
        "Дорого", "Другая категория", "Нарушение правил ASD", "Не вышли на контактное лицо по заявке", _
        "Не настроен на диалог", "Не прошел по бюджету/тратам", "Недостаточный ассортимент/найм", "Нерегулярный найм/Сезонность продаж", "Нецелевой клиент", _
        "Низкая потребность", "Не настроен на диалог, Согласие", "Работает менеджер", "Сложное возражение", "Частный клиент", _
        "="), Operator:=xlFilterValues
    Rows("2:1048576").Delete Shift:=xlUp
    Range("$A$1:$AP$1048576").AutoFilter Field:=14
    Cells.Replace What:="Поле ввода не заполнено", Replacement:="", LookAt:= _
        xlWhole, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="Ответ не сохранен", Replacement:="", LookAt:=xlWhole _
        , SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("W2").EntireColumn.Insert
    Range("W2:W1048576").FormulaR1C1 = _
        "=IF(RC[1]="""","""",RC[1]&"" | ""&RC[2]&"" | ""&RC[3]&"" | ""&RC[4]&"" | ""&RC[5]&"" | ""&RC[6]&"" | ""&RC[7]&"" | ""&RC[8]&"" | ""&RC[9]&"" | ""&RC[10]&"" | ""&RC[11]&"" | ""&RC[12]&"" | ""&RC[13]&"" | ""&RC[14]&"" | ""&RC[15]&"" | ""&RC[16]&"" | ""&RC[17]&"" | ""&RC[18]&"" | ""&RC[19]&"" | ""&R[-1]C[20])"
    Application.CutCopyMode = False
    Columns("W:W").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("X:AS").Delete Shift:=xlToLeft
    Columns("A:A").Delete Shift:=xlToLeft
    Range("E1:F1").EntireColumn.Delete
    Range("H1:I1").EntireColumn.Delete
    Range("K1:M1").EntireColumn.Delete
    Range("M1:N1").EntireColumn.Delete
    Columns("C:C").Cut
    Columns("M:M").Insert Shift:=xlToRight
    Range("M1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Комментарий"
    Range("M2:M1048576").FormulaR1C1 = _
        "=IF(RC[-12]="""","""",""Дозвонились по номеру ""&RC[-1]&IF(RC[1]="""","""","" | ""&RC[1])&IF(RC[-4]="""","""","" | ""&RC[-4])&IF(RC[-7]="""","""","" | ""&RC[-7])&IF(RC[-6]="""","""","" | ""&RC[-6]))"
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("N:N").Delete Shift:=xlToLeft
    Columns("H:H").Cut
    Columns("N:N").Insert Shift:=xlToRight
    Range("F1:H1").EntireColumn.Delete
    Columns("C:C").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
End Sub
