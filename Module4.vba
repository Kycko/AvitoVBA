Sub CRMmrkg()
    ' Обработка еженедельных CRMmrkg.

    Dim resp, tempStr, tempNum, tempNum2, counter, SHlist, searchList, titles, finalOrder, regVar, Lname, Pname, WGpriority, indxList(5), testMSG  ' для indxList указать кол-во из searchList!!
    SHlist = Array("White", "Grey", "WG")
    searchList = Array("external", "name", "email", "phone", "category", "region")    ' все столбцы, номера которых надо запомнить
    titles = Array("Авито-аккаунт", "Название компании", "Рабочий e-mail", "Основной телефон", "Категория", "Регион и город")   ' заголовки
    regVar = Array("region", "city")
    finalOrder = Array("region", "category", "Вертикаль", "Источник", "Направление клиента", "Микрокатегория", "Название лида", "Наименование проекта", "name", "Имя", "phone", "email", "Статус", "Ответственный", "Доступен для всех", "Комментарий", "external")

    resp = MsgBox("1. Должны быть созданы все 7 листов с такими названиями:" & vbNewLine & "White, Grey, WG, cities, cat, log cat, prev." & vbNewLine & vbNewLine & "2. ПРОВЕРИТЬ, что в White/Grey priority нет пустых строк (это ошибка импорта)", vbQuestion + vbYesNo + vbDefaultButton2)
    If resp = vbYes Then
        For Each Sheet In SHlist
            ' ищем столбцы
            For i = 0 To UBound(searchList)
                If searchList(i) = "region" Then
                    For Each reg In regVar
                        If Not sheets(Sheet).Range("A1:Z1").Find(reg) Is Nothing Then tempStr = reg
                      Next
                  Else
                    tempStr = searchList(i)
                  End If

                ' запоминаем номера столбцов
                indxList(i) = sheets(Sheet).Range("A1:Z1").Find(tempStr, LookIn:=xlValues, LookAt:=xlPart).Column
              Next i
            If Sheet = "WG" Then WGpriority = sheets(Sheet).Range("A1:Z1").Find("priority", LookIn:=xlValues, LookAt:=xlPart).Column

            ' создаём временный лист
            sheets.Add After:=sheets(sheets.count)
            ActiveSheet.Name = Sheet & " temp"

            ' выстраиваем столбцы в правильном порядке и меняем заголовки
            counter = 1
            For Each col In finalOrder
                tempNum = Application.Match(col, searchList, False)
                If IsNumeric(tempNum) Then
                    tempNum = tempNum - 1
                    sheets(Sheet).Columns(indxList(tempNum)).Copy ActiveSheet.Columns(counter)
                    ActiveSheet.Cells(1, counter).Value = titles(tempNum)
                  Else
                    ActiveSheet.Cells(1, counter).Value = col
                  End If
                counter = counter + 1
              Next col

            ' удаляем лишние приоритеты в WG
            If Sheet = "WG" Then
                For i = Cells.CurrentRegion.Rows.Count To 2 step -1
                    tempStr = sheets(Sheet).Cells(i, WGpriority).Value
                    If not tempStr = "1" And not tempStr = "2" And not tempStr = "3" And not tempStr = "4" Then
                        ActiveSheet.Rows(i).Delete
                      End If
                  Next i
              End If

            ' добавляем название лида и наименование проекта
            tempNum  = ActiveSheet.Range("A1:Z1").Find("Название лида").Column
            tempNum2 = ActiveSheet.Range("A1:Z1").Find("Наименование проекта").Column
            If Sheet = "White" Then
                Lname = "GE CRMmrkg NWL Hunter " & Date
                Pname = "GE | CRMmrkg | NWL | Hunter"
              ElseIf Sheet = "Grey" Then
                Lname = "GE CRMmrkg B2C Grey Hunter " & Date
                Pname = "GE | CRMmrkg | B2C Grey | Hunter"
              Else
                Lname = "GE CRMmrkg PRIOR Hunter"
                Pname = "GE | CRMmrkg | PRIOR | Hunter"
              End If
            For i = 2 To Cells.CurrentRegion.Rows.Count
                tempStr = ""
                If Sheet = "WG" Then
                    tempStr = sheets(Sheet).Cells(i, WGpriority).Value
                    If tempStr = "1" or tempStr = "2" or tempStr = "3" Then
                        tempStr = "White " & tempStr
                      Else
                        tempStr = "Grey " & tempStr
                      End If
                  End If
                ActiveSheet.Cells(i, tempNum ).Value = Replace(Lname, "PRIOR", tempStr)
                ActiveSheet.Cells(i, tempNum2).Value = Replace(Pname, "PRIOR", tempStr)
              Next i
          Next Sheet

        ' ActiveSheet.Range("$A:$S").RemoveDuplicates Columns:=1, Header:=xlNo    ' удаляем дубликаты
      End If

  End Sub
