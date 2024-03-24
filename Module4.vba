Sub CRMmrkg()
    ' Обработка еженедельных CRMmrkg.

    Dim resp, tempStr, tempNum, counter, SHlist, searchList, titles, finalOrder, regVar, WGpriority, indxList(5), testMSG  ' для indxList указать кол-во из searchList!!
    SHlist = Array("White", "Grey", "WG")
    searchList = Array("external", "name", "email", "phone", "category", "region")    ' все столбцы, номера которых надо запомнить
    titles = Array("Авито-аккаунт", "Название компании", "Рабочий e-mail", "Основной телефон", "Категория", "Регион и город")   ' заголовки
    regVar = Array("region", "city")
    finalOrder = Array("region", "category", "Вертикаль", "Источник", "Направление клиента", "Микрокатегория", "Название лида", "Наименование проекта", "name", "Имя", "phone", "email", "Статус", "Ответственный", "Доступен для всех", "Комментарий", "external")

    resp = MsgBox("Вы создали все 7 листов с такими названиями?" & vbNewLine & "White, Grey, WG, cities, cat, log cat, prev.", vbQuestion + vbYesNo + vbDefaultButton2)
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
        Next Sheet

        ' ActiveSheet.Range("$A:$S").RemoveDuplicates Columns:=1, Header:=xlNo    ' удаляем дубликаты
    End If

End Sub
