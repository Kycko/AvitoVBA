Sub BMC_CRMM()
    ' BMC_CRMM Макрос
    ' Ежедневные выгрузки BMC CRMM

    Dim months
    months = Array("Jan", "Feb", "Mrch", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    Cells.Select
    Selection.Copy
    sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Columns("A:A").Delete Shift:=xlToLeft
    Columns("B:C").Delete Shift:=xlToLeft
    Range("B1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Комментарий"
    Range("B2:B1000").FormulaR1C1 = "=IF(RC[1]="""","""",R1C[2]&"" ""&RC[2]&"" | ""&R1C[3]&"" ""&RC[3]&"" | ""&RC[1])"
    Application.CutCopyMode = False
    Columns("B:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:E").Delete Shift:=xlToLeft
    Range("A1").FormulaR1C1 = "Авито-аккаунт"
    Range("A1:M1").EntireColumn.Insert
    
    Range("A1").FormulaR1C1 = "Регион и город"
    Range("B1").FormulaR1C1 = "Категория"
    Range("C1").FormulaR1C1 = "Вертикаль"
    Range("D1").FormulaR1C1 = "Источник"
    Range("E1").FormulaR1C1 = "Ответственный менеджер в сделке"
    Range("F1").FormulaR1C1 = "Название лида"
    Range("G1").FormulaR1C1 = "Наименование проекта"
    Range("H1").FormulaR1C1 = "Название компании"
    Range("I1").FormulaR1C1 = "Имя"
    Range("J1").FormulaR1C1 = "Основной телефон"
    Range("K1").FormulaR1C1 = "Статус"
    Range("L1").FormulaR1C1 = "Ответственный"
    Range("M1").FormulaR1C1 = "Доступен для всех"
    
    Range("A2:A1000").FormulaR1C1 = "=IF(RC14="""","""",""Другие регионы России"")"
    Range("B2:B1000").FormulaR1C1 = "=IF(RC14="""","""",""Вакансии"")"
    Range("C2:C1000").FormulaR1C1 = "=IF(RC14="""","""",""Работа"")"
    Range("D2:D1000").FormulaR1C1 = "=IF(RC14="""","""",""CRM маркетинг"")"
    Range("E2:E1000").FormulaR1C1 = "=IF(RC14="""","""",""Кристина Лебедева"")"
    Range("F2:F1000").FormulaR1C1 = "=IF(RC14="""","""",""Job BMC-CRMM " & Date & """)"
    Range("G2:G1000").FormulaR1C1 = "=IF(RC14="""","""",""Job | BMC-CRMM " & months(Month(Date) - 1) & """)"
    Range("H2:H1000").FormulaR1C1 = "=IF(RC14="""","""",""Unknown"")"
    Range("I2:I1000").FormulaR1C1 = "=IF(RC14="""","""",""Unknown"")"
    Range("J2:J1000").FormulaR1C1 = "=IF(RC14="""","""",""79999999999"")"
    Range("K2:K1000").FormulaR1C1 = "=IF(RC14="""","""",""Новый"")"
    Range("L2:L1000").FormulaR1C1 = "=IF(RC14="""","""",""Квалификаторы"")"
    Range("M2:M1000").FormulaR1C1 = "=IF(RC14="""","""",""Да"")"
    
    Columns("A:M").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Cells.EntireColumn.AutoFit
    
    Range("O1:O11").Select
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10724347
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
