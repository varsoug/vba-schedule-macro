Attribute VB_Name = "Module1"
Option Explicit

Private Function NormalizePhone(ByVal v As Variant) As String
    Dim s As String, digits As String
    Dim i As Long, ch As String
    s = CStr(v)

    ' Вытащим только цифры
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i

    ' Приведём к российскому формату
    ' Варианты:
    '  - 10 цифр (без кода страны) -> добавим ведущую 7
    '  - 11 цифр, первая 8 -> заменим на 7
    '  - 11 цифр, первая 7 -> ок
    ' Остальное — возвращаем как было (без форматирования)
    If Len(digits) = 10 Then
        digits = "7" & digits
    ElseIf Len(digits) = 11 Then
        If Left$(digits, 1) = "8" Then
            digits = "7" & Mid$(digits, 2)
        ElseIf Left$(digits, 1) <> "7" Then
            NormalizePhone = s
            Exit Function
        End If
    Else
        NormalizePhone = s
        Exit Function
    End If

    ' Форматируем: +7 (XXX) XXX-XX-XX
    NormalizePhone = "+7 (" & Mid$(digits, 2, 3) & ") " & _
                     Mid$(digits, 5, 3) & "-" & Mid$(digits, 8, 2) & "-" & Mid$(digits, 10, 2)
End Function

Sub Расписание()
    Dim srcSheet As Worksheet, dstSheet As Worksheet, tempSheet As Worksheet
    Dim lastRow As Long, i As Long, dstRow As Long, tempRow As Long
    Dim fio As String
    Dim app As Application
    Dim currentGroup As Variant, previousGroup As Variant
    Dim currentSubGroup As String, previousSubGroup As String
    Dim numInGroup As Long
    Dim dictSectionGroups As Object
    Dim key As Variant
    Dim groupsAlwaysOne As Boolean
    Dim employersPresent As Boolean
    Dim lastColOut As Long
    Dim previousDate As Variant
    Dim rngMerge As Range
    Dim minTestDate As Date, maxTestDate As Date
    Dim hasTestDate As Boolean
    Dim dictExam As Object, exKey As Variant, onlyExam As String
    Dim rowIndex As Long
    Dim examValue As String

    Set app = Application
    app.ScreenUpdating = False

    Set dictExam = CreateObject("Scripting.Dictionary")
    ' Исходный лист — активный
    Set srcSheet = ActiveSheet
    lastRow = srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row

    ' Создаем новую книгу для результата
    Workbooks.Add
    Set dstSheet = ActiveSheet

    ' Создаем временный лист
    Set tempSheet = dstSheet.Parent.Sheets.Add(After:=dstSheet)
    On Error Resume Next
    tempSheet.Name = "TempData"
    On Error GoTo 0

    ' Заголовки временного листа (J = Работодатель)
    tempSheet.Range("A1:J1").Value = Array( _
        "Заявка", "ФИО", "Дата рождения", "Телефон", "Гражданство", _
        "Экзамен", "Аудитория", "Время", "Группа", "Работодатель")

    ' Сбор данных
    tempRow = 2
    For i = 2 To lastRow
        If Trim(srcSheet.Cells(i, 29).Value) = "Активная" Then
            fio = Trim(srcSheet.Cells(i, 1).Value & " " & srcSheet.Cells(i, 2).Value)
            If srcSheet.Cells(i, 3).Value <> "" Then fio = fio & " " & srcSheet.Cells(i, 3).Value

            tempSheet.Cells(tempRow, 1).Value = srcSheet.Cells(i, 27).Value  ' Заявка
            tempSheet.Cells(tempRow, 2).Value = fio                           ' ФИО
            tempSheet.Cells(tempRow, 3).Value = srcSheet.Cells(i, 4).Value    ' ДР

            ' Телефон -> нормализуем к +7 (XXX) XXX-XX-XX
            tempSheet.Cells(tempRow, 4).Value = NormalizePhone(srcSheet.Cells(i, 5).Value)

            tempSheet.Cells(tempRow, 5).Value = srcSheet.Cells(i, 6).Value    ' Гражданство
            tempSheet.Cells(tempRow, 6).Value = srcSheet.Cells(i, 8).Value    ' Экзамен
            tempSheet.Cells(tempRow, 7).Value = srcSheet.Cells(i, 10).Value   ' Аудитория
            tempSheet.Cells(tempRow, 8).Value = srcSheet.Cells(i, 12).Value   ' Время
            tempSheet.Cells(tempRow, 9).Value = srcSheet.Cells(i, 13).Value   ' Группа
            tempSheet.Cells(tempRow, 10).Value = srcSheet.Cells(i, 14).Value  ' Работодатель (столбец N)

            tempRow = tempRow + 1
        End If
    Next i

    If tempRow <= 2 Then
        MsgBox "Нет активных записей для обработки.", vbInformation
        Application.DisplayAlerts = False
        On Error Resume Next: tempSheet.Delete: On Error GoTo 0
        Application.DisplayAlerts = True
        app.ScreenUpdating = True
        Exit Sub
    End If

    ' Есть ли где-либо указанный работодатель?
    employersPresent = Application.WorksheetFunction.CountA(tempSheet.Range("J2:J" & tempRow - 1)) > 0

    ' Сортировка: Экзамен -> Время -> Группа -> ФИО
    With tempSheet.Sort
        .SortFields.Clear
        .SortFields.Add key:=tempSheet.Range("F2:F" & tempRow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=tempSheet.Range("H2:H" & tempRow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=tempSheet.Range("I2:I" & tempRow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=tempSheet.Range("B2:B" & tempRow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange tempSheet.Range("A1:J" & tempRow - 1)
        .Header = xlYes
        .Apply
    End With

    ' Уникальные группы по секциям ("Экзамен + Время")
    Set dictSectionGroups = CreateObject("Scripting.Dictionary")
    For i = 2 To tempRow - 1
        ' Проверим смену даты (подзаголовок по дате)
        Dim currDate As Variant
        currDate = tempSheet.Cells(i, 7).Value ' Дата тестирования в tempSheet
        If CStr(currDate) <> CStr(previousDate) Then
            ' Вставим строку с датой
            
            If lastColOut < 1 Then lastColOut = 1
If dstRow > 0 Then
    With dstSheet.Range(dstSheet.Cells(dstRow, 1), dstSheet.Cells(dstRow, lastColOut))
        If .MergeCells Then .UnMerge
        If IsDate(currDate) Then
            .Cells(1, 1).Value = Format(CDate(currDate), "dd.mm.yyyy")
        Else
            .Cells(1, 1).Value = CStr(currDate)
        End If
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End If
dstRow = dstRow + 1
previousDate = currDate
previousGroup = vbNullString
previousSubGroup = vbNullString
        End If

        key = Trim(CStr(tempSheet.Cells(i, 6).Value)) & " " & Format(tempSheet.Cells(i, 8).Value, "hh:mm")
        currentSubGroup = Trim(CStr(tempSheet.Cells(i, 9).Value))
        If Not dictSectionGroups.Exists(key) Then
            dictSectionGroups.Add key, CreateObject("Scripting.Dictionary")
        End If
        dictSectionGroups(key).Item(currentSubGroup) = 1
    Next i

    ' Во всех ли секциях ровно одна группа?
    groupsAlwaysOne = True
    For Each key In dictSectionGroups.Keys
        If dictSectionGroups(key).Count > 1 Then
            groupsAlwaysOne = False
            Exit For
        End If
    Next key

    ' Заголовки итогового листа
    If groupsAlwaysOne Then
        If employersPresent Then
            dstSheet.Range("A1:J1").Value = Array("№", "Заявка", "ФИО", "Дата рождения", "Телефон", _
                                                  "Гражданство", "Экзамен", "Аудитория", "Время", "Работодатель")
            lastColOut = 10
        Else
            dstSheet.Range("A1:I1").Value = Array("№", "Заявка", "ФИО", "Дата рождения", "Телефон", _
                                                  "Гражданство", "Экзамен", "Аудитория", "Время")
            lastColOut = 9
        End If
    Else
        If employersPresent Then
            dstSheet.Range("A1:K1").Value = Array("№", "Заявка", "ФИО", "Дата рождения", "Телефон", _
                                                  "Гражданство", "Экзамен", "Аудитория", "Время", "Группа", "Работодатель")
            lastColOut = 11
        Else
            dstSheet.Range("A1:J1").Value = Array("№", "Заявка", "ФИО", "Дата рождения", "Телефон", _
                                                  "Гражданство", "Экзамен", "Аудитория", "Время", "Группа")
            lastColOut = 10
        End If
    End If

    ' Перенос данных с заголовками/подзаголовками
    dstRow = 2
    previousGroup = ""
    previousSubGroup = ""
    numInGroup = 1

    For i = 2 To tempRow - 1
        currentGroup = Trim(CStr(tempSheet.Cells(i, 6).Value)) & " " & Format(tempSheet.Cells(i, 8).Value, "hh:mm")
        currentSubGroup = Trim(CStr(tempSheet.Cells(i, 9).Value))

        ' Заголовок секции
        If currentGroup <> previousGroup Then
            dstSheet.Cells(dstRow, 1).Value = UCase(currentGroup)
            
            If lastColOut < 1 Then lastColOut = 1
            dstSheet.Rows(dstRow).UnMerge
            Set rngMerge = dstSheet.Range(dstSheet.Cells(dstRow, 1), dstSheet.Cells(dstRow, lastColOut))
            rngMerge.Merge
            With dstSheet.Cells(dstRow, 1)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.Color = RGB(220, 220, 220)
            End With
            dstRow = dstRow + 1
            previousGroup = currentGroup
            previousSubGroup = ""
            numInGroup = 1
        End If

        ' Подзаголовок группы (если групп > 1)
        If Not groupsAlwaysOne Then
            If currentSubGroup <> previousSubGroup Then
                dstSheet.Cells(dstRow, 1).Value = UCase(currentSubGroup)
                
            If lastColOut < 1 Then lastColOut = 1
            dstSheet.Rows(dstRow).UnMerge
            Set rngMerge = dstSheet.Range(dstSheet.Cells(dstRow, 1), dstSheet.Cells(dstRow, lastColOut))
            rngMerge.Merge
                With dstSheet.Cells(dstRow, 1)
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Interior.Color = RGB(240, 240, 240)
                End With
                dstRow = dstRow + 1
                previousSubGroup = currentSubGroup
                numInGroup = 1
            End If
        End If

        ' Строка студента (с номером)
        dstSheet.Cells(dstRow, 1).Value = numInGroup

        If groupsAlwaysOne Then
            ' Копируем поля Заявка..Время (A..H -> 2..9)
            dstSheet.Range(dstSheet.Cells(dstRow, 2), dstSheet.Cells(dstRow, 9)).Value = _
                tempSheet.Range(tempSheet.Cells(i, 1), tempSheet.Cells(i, 8)).Value
            ' Работодатель, если нужен
            If employersPresent Then dstSheet.Cells(dstRow, 10).Value = tempSheet.Cells(i, 10).Value
        Else
            ' Копируем поля Заявка..Группа (A..I -> 2..10)
            dstSheet.Range(dstSheet.Cells(dstRow, 2), dstSheet.Cells(dstRow, 10)).Value = _
                tempSheet.Range(tempSheet.Cells(i, 1), tempSheet.Cells(i, 9)).Value
            ' Работодатель, если нужен
            If employersPresent Then dstSheet.Cells(dstRow, 11).Value = tempSheet.Cells(i, 10).Value
        End If

        numInGroup = numInGroup + 1
        dstRow = dstRow + 1
    Next i

    ' Удаляем временный лист
    Application.DisplayAlerts = False
    On Error Resume Next
    tempSheet.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Финальное форматирование
    dstSheet.Columns.AutoFit
    dstSheet.Rows(1).Font.Bold = True
    dstSheet.Range(dstSheet.Cells(1, 1), dstSheet.Cells(1, lastColOut)).Interior.Color = RGB(200, 200, 200)
    dstSheet.Rows(1).HorizontalAlignment = xlCenter
    dstSheet.Columns("D").NumberFormat = "dd.mm.yyyy" ' ДР
    dstSheet.Columns("I").NumberFormat = "hh:mm"       ' Время (I в любом варианте)
    dstSheet.Columns("A").HorizontalAlignment = xlCenter ' номера по центру
    dstSheet.Columns("E").NumberFormat = "@"            ' Телефон как текст

    ' Границы по всей итоговой таблице
    With dstSheet.Range(dstSheet.Cells(1, 1), dstSheet.Cells(dstRow - 1, lastColOut)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With

    
    ' Подготовим titleOverride и удалим столбец Экзамен, если он один
    Dim titleOverride As String
    Dim datePart As String
    If hasTestDate Then
        If minTestDate <> maxTestDate Then
            datePart = " НА " & Format(minTestDate, "dd.mm.yyyy") & " - " & Format(maxTestDate, "dd.mm.yyyy")
        Else
            datePart = " НА " & Format(minTestDate, "dd.mm.yyyy")
        End If
    Else
        datePart = ""
    End If

    Dim exColHeader As Long: exColHeader = 0
    Dim c As Long
    For c = 1 To lastColOut
        If Trim$(CStr(dstSheet.Cells(1, c).Value)) = "Экзамен" Then exColHeader = c: Exit For
    Next c
    If exColHeader > 0 Then
        For rowIndex = 2 To dstRow - 1
            examValue = Trim$(CStr(dstSheet.Cells(rowIndex, exColHeader).Value))
            If Len(examValue) > 0 Then
                If Not dictExam.Exists(examValue) Then dictExam.Add examValue, 1
            End If
        Next rowIndex
    End If


    If dictExam.Count = 1 Then
        onlyExam = dictExam.Keys()(0)
        titleOverride = "РАСПИСАНИЕ НА ЭКЗАМЕН " & UCase$(onlyExam) & datePart
        If exColHeader > 0 Then
            dstSheet.Columns(exColHeader).Delete
            lastColOut = lastColOut - 1
        End If
    Else
        titleOverride = "РАСПИСАНИЕ" & datePart
    End If
app.ScreenUpdating = True
    ' Добавлено: шапка расписания
    On Error Resume Next
    Call AddScheduleHeader(dstSheet, lastColOut, dstRow - 1, IIf(hasTestDate, minTestDate, Empty), IIf(hasTestDate, maxTestDate, Empty), titleOverride)
    On Error GoTo 0
    MsgBox "Расписание готово"
End Sub

' === Добавлено: Шапка "РАСПИСАНИЕ НА ДД.ММ.ГГГГ" ===
Private Function DarkenColor(ByVal clr As Long, Optional ByVal factor As Double = 0.85) As Long
    Dim r As Long, g As Long, b As Long
    r = (clr And &HFF)
    g = (clr \ &H100) And &HFF
    b = (clr \ &H10000) And &HFF
    r = CLng(r * factor): If r < 0 Then r = 0
    g = CLng(g * factor): If g < 0 Then g = 0
    b = CLng(b * factor): If b < 0 Then b = 0
    DarkenColor = RGB(r, g, b)
End Function

Private Sub AddScheduleHeader(ByVal dstSheet As Worksheet, ByRef lastColOut As Long, ByVal lastDataRow As Long, Optional ByVal dtMin As Variant, Optional ByVal dtMax As Variant, Optional ByVal titleOverride As String = "")
    On Error GoTo AddHeaderFail
    Dim headerRow As Long: headerRow = 1
    Dim findRow As Long: findRow = 1
    Dim dtCol As Long: dtCol = 0
    Dim exCol As Long: exCol = 0
    Dim c As Long, r As Long
    Dim dtVal As Variant
    Dim minDate As Date, maxDate As Date
    Dim hasDate As Boolean
    Dim dictEx As Object: Set dictEx = CreateObject("Scripting.Dictionary")
    
    
    ' При наличии заранее собранного заголовка — используем его и не сканируем таблицу
    If Len(titleOverride) > 0 Then
        ' Вставим строку сверху
        dstSheet.Rows(headerRow).Insert Shift:=xlDown
        With dstSheet.Range(dstSheet.Cells(headerRow, 1), dstSheet.Cells(headerRow, lastColOut))
            .Merge
            .Value = titleOverride
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .RowHeight = 24
            Dim belowColorQuick As Long
            belowColorQuick = dstSheet.Range(dstSheet.Cells(headerRow + 1, 1), dstSheet.Cells(headerRow + 1, lastColOut)).Interior.Color
            If belowColorQuick = 0 Then belowColorQuick = RGB(200, 200, 200)
            .Interior.Color = DarkenColor(belowColorQuick, 0.85)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End With
        Exit Sub
    End If

' Найдём колонки по заголовкам
    For c = 1 To lastColOut
        If Trim$(CStr(dstSheet.Cells(findRow, c).Value)) = "Дата тестирования" Then dtCol = c
        If Trim$(CStr(dstSheet.Cells(findRow, c).Value)) = "Экзамен" Then exCol = c
    Next c
    If dtCol = 0 Then
        findRow = 2
        For c = 1 To lastColOut
            If Trim$(CStr(dstSheet.Cells(findRow, c).Value)) = "Дата тестирования" Then dtCol = c
            If Trim$(CStr(dstSheet.Cells(findRow, c).Value)) = "Экзамен" Then exCol = c
        Next c
    End If
    
    ' Если даты переданы параметрами — используем их
    If IsDate(dtMin) And IsDate(dtMax) Then
        minDate = CDate(dtMin)
        maxDate = CDate(dtMax)
        hasDate = True
    End If

    ' Соберём диапазон дат и множество экзаменов
    If dtCol > 0 Then
        For r = findRow + 1 To lastDataRow
            dtVal = dstSheet.Cells(r, dtCol).Value
            If IsDate(dtVal) Then
                If Not hasDate Then
                    minDate = CDate(dtVal)
                    maxDate = CDate(dtVal)
                    hasDate = True
                Else
                    If CDate(dtVal) < minDate Then minDate = CDate(dtVal)
                    If CDate(dtVal) > maxDate Then maxDate = CDate(dtVal)
                End If
            End If
            If exCol > 0 Then
                Dim exv As String
                exv = Trim$(CStr(dstSheet.Cells(r, exCol).Value))
                If exv <> "" Then If Not dictEx.Exists(exv) Then dictEx.Add exv, 1
            End If
        Next r
    End If
    
    ' Если экзамен один — удалим столбец "Экзамен"
    Dim titleExamPart As String: titleExamPart = ""
    If dictEx.Count = 1 And exCol > 0 Then
        Dim onlyExam As String
        onlyExam = dictEx.Keys()(0)
        titleExamPart = " НА ЭКЗАМЕН " & UCase$(onlyExam)
        ' Удаляем столбец и корректируем ширину
        dstSheet.Columns(exCol).Delete
        lastColOut = lastColOut - 1
        If dtCol > exCol Then dtCol = dtCol - 1 ' если Дата правее Экзамена, её индекс сместится
    End If
    
    ' Сформируем заголовок: если несколько дат -> диапазон
    Dim titleText As String
    If hasDate Then
        If minDate <> maxDate Then
            titleText = "РАСПИСАНИЕ" & titleExamPart & " НА " & Format(minDate, "dd.mm.yyyy") & " - " & Format(maxDate, "dd.mm.yyyy")
        Else
            titleText = "РАСПИСАНИЕ" & titleExamPart & " НА " & Format(minDate, "dd.mm.yyyy")
        End If
    Else
        titleText = "РАСПИСАНИЕ" & titleExamPart
    End If
    
    ' Вставим строку сверху
    dstSheet.Rows(headerRow).Insert Shift:=xlDown
    ' Объединяем на всю ширину таблицы
    With dstSheet.Range(dstSheet.Cells(headerRow, 1), dstSheet.Cells(headerRow, lastColOut))
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .RowHeight = 24
        ' Цвет: чуть темнее, чем нижняя строка
        Dim belowColor As Long
        belowColor = dstSheet.Range(dstSheet.Cells(headerRow + 1, 1), dstSheet.Cells(headerRow + 1, lastColOut)).Interior.Color
        If belowColor = 0 Then belowColor = RGB(200, 200, 200) ' запасной вариант
        .Interior.Color = DarkenColor(belowColor, 0.85)
        ' Границы вокруг объединённой ячейки
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End With
    
    Exit Sub
AddHeaderFail:
    ' Если что-то пошло не так, тихо пропускаем
End Sub