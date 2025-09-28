Attribute VB_Name = "Module1"
Option Explicit

Private Function NormalizePhone(ByVal v As Variant) As String
    Dim s As String, digits As String
    Dim i As Long, ch As String
    s = CStr(v)

    ' Âûòàùèì òîëüêî öèôðû
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i

    ' Ïðèâåä¸ì ê ðîññèéñêîìó ôîðìàòó
    ' Âàðèàíòû:
    '  - 10 öèôð (áåç êîäà ñòðàíû) -> äîáàâèì âåäóùóþ 7
    '  - 11 öèôð, ïåðâàÿ 8 -> çàìåíèì íà 7
    '  - 11 öèôð, ïåðâàÿ 7 -> îê
    ' Îñòàëüíîå — âîçâðàùàåì êàê áûëî (áåç ôîðìàòèðîâàíèÿ)
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

    Dim dictExam As Object, onlyExam As String
    Dim singleTestDate As Boolean, singleTestTime As Boolean
    Dim hasTestTime As Boolean
    Dim minTestTime As Date, maxTestTime As Date
    Dim firstTestDate As Variant, firstTestTime As Variant
    Dim examValue As String
    Dim dtVal As Variant, tmVal As Variant
    Dim dt As Date, tm As Date
    Dim removeDateColumn As Boolean, removeTimeColumn As Boolean
    Dim dateColIndex As Long, timeColIndex As Long, phoneColIndex As Long
    Dim headerText As String
    Set dictExam = CreateObject("Scripting.Dictionary")
    singleTestDate = True
    singleTestTime = True
            examValue = Trim$(CStr(tempSheet.Cells(tempRow, 1).Value))
            If Len(examValue) > 0 Then
                If Not dictExam.Exists(examValue) Then dictExam.Add examValue, 1
            End If

            dtVal = tempSheet.Cells(tempRow, 3).Value
            If IsDate(dtVal) Then
                dt = DateValue(CDate(dtVal))
                If Not hasTestDate Then
                    minTestDate = dt
                    maxTestDate = dt
                    hasTestDate = True
                    firstTestDate = dt
                Else
                    If dt < minTestDate Then minTestDate = dt
                    If dt > maxTestDate Then maxTestDate = dt
                    If singleTestDate And dt <> firstTestDate Then singleTestDate = False
                End If
            End If

            tmVal = tempSheet.Cells(tempRow, 8).Value
            If IsDate(tmVal) Then
                tm = TimeValue(CDate(tmVal))
                If Not hasTestTime Then
                    minTestTime = tm
                    maxTestTime = tm
                    hasTestTime = True
                    firstTestTime = tm
                Else
                    If tm < minTestTime Then minTestTime = tm
                    If tm > maxTestTime Then maxTestTime = tm
                    If singleTestTime And tm <> firstTestTime Then singleTestTime = False
                End If
            End If

    NormalizePhone = "+7 (" & Mid$(digits, 2, 3) & ") " & _
                     Mid$(digits, 5, 3) & "-" & Mid$(digits, 8, 2) & "-" & Mid$(digits, 10, 2)
End Function

Sub Ðàñïèñàíèå()
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

    Set app = Application
    app.ScreenUpdating = False

    ' Èñõîäíûé ëèñò — àêòèâíûé
    Set srcSheet = ActiveSheet
    lastRow = srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row

    ' Ñîçäàåì íîâóþ êíèãó äëÿ ðåçóëüòàòà
    Workbooks.Add
    Set dstSheet = ActiveSheet

    ' Ñîçäàåì âðåìåííûé ëèñò
    Set tempSheet = dstSheet.Parent.Sheets.Add(After:=dstSheet)
    On Error Resume Next
    tempSheet.Name = "TempData"
    On Error GoTo 0

    ' Çàãîëîâêè âðåìåííîãî ëèñòà (J = Ðàáîòîäàòåëü)
    tempSheet.Range("A1:J1").Value = Array( _
        "Çàÿâêà", "ÔÈÎ", "Äàòà ðîæäåíèÿ", "Òåëåôîí", "Ãðàæäàíñòâî", _
        "Ýêçàìåí", "Àóäèòîðèÿ", "Âðåìÿ", "Ãðóïïà", "Ðàáîòîäàòåëü")

    ' Ñáîð äàííûõ
    tempRow = 2
    For i = 2 To lastRow
        If Trim(srcSheet.Cells(i, 29).Value) = "Àêòèâíàÿ" Then
            fio = Trim(srcSheet.Cells(i, 1).Value & " " & srcSheet.Cells(i, 2).Value)
            If srcSheet.Cells(i, 3).Value <> "" Then fio = fio & " " & srcSheet.Cells(i, 3).Value

            tempSheet.Cells(tempRow, 1).Value = srcSheet.Cells(i, 27).Value  ' Çàÿâêà
            tempSheet.Cells(tempRow, 2).Value = fio                           ' ÔÈÎ
            tempSheet.Cells(tempRow, 3).Value = srcSheet.Cells(i, 4).Value    ' ÄÐ

            ' Òåëåôîí -> íîðìàëèçóåì ê +7 (XXX) XXX-XX-XX
            tempSheet.Cells(tempRow, 4).Value = NormalizePhone(srcSheet.Cells(i, 5).Value)

            tempSheet.Cells(tempRow, 5).Value = srcSheet.Cells(i, 6).Value    ' Ãðàæäàíñòâî
            tempSheet.Cells(tempRow, 6).Value = srcSheet.Cells(i, 8).Value    ' Ýêçàìåí
            tempSheet.Cells(tempRow, 7).Value = srcSheet.Cells(i, 10).Value   ' Àóäèòîðèÿ
            tempSheet.Cells(tempRow, 8).Value = srcSheet.Cells(i, 12).Value   ' Âðåìÿ
            tempSheet.Cells(tempRow, 9).Value = srcSheet.Cells(i, 13).Value   ' Ãðóïïà
            tempSheet.Cells(tempRow, 10).Value = srcSheet.Cells(i, 14).Value  ' Ðàáîòîäàòåëü (ñòîëáåö N)

            tempRow = tempRow + 1
        End If
    Next i

    If tempRow <= 2 Then
        MsgBox "Íåò àêòèâíûõ çàïèñåé äëÿ îáðàáîòêè.", vbInformation
        Application.DisplayAlerts = False
        On Error Resume Next: tempSheet.Delete: On Error GoTo 0
        Application.DisplayAlerts = True
        app.ScreenUpdating = True
        Exit Sub
    End If

    ' Åñòü ëè ãäå-ëèáî óêàçàííûé ðàáîòîäàòåëü?
    employersPresent = Application.WorksheetFunction.CountA(tempSheet.Range("J2:J" & tempRow - 1)) > 0

    ' Ñîðòèðîâêà: Ýêçàìåí -> Âðåìÿ -> Ãðóïïà -> ÔÈÎ
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

    ' Óíèêàëüíûå ãðóïïû ïî ñåêöèÿì ("Ýêçàìåí + Âðåìÿ")
    Set dictSectionGroups = CreateObject("Scripting.Dictionary")
    For i = 2 To tempRow - 1
        ' Ïðîâåðèì ñìåíó äàòû (ïîäçàãîëîâîê ïî äàòå)
        Dim currDate As Variant
        currDate = tempSheet.Cells(i, 7).Value ' Äàòà òåñòèðîâàíèÿ â tempSheet
        If CStr(currDate) <> CStr(previousDate) Then
            ' Âñòàâèì ñòðîêó ñ äàòîé
            
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

    ' Âî âñåõ ëè ñåêöèÿõ ðîâíî îäíà ãðóïïà?
    groupsAlwaysOne = True
    For Each key In dictSectionGroups.Keys
        If dictSectionGroups(key).Count > 1 Then
            groupsAlwaysOne = False
            Exit For
        End If
    Next key

    ' Çàãîëîâêè èòîãîâîãî ëèñòà
    If groupsAlwaysOne Then
        If employersPresent Then
            dstSheet.Range("A1:J1").Value = Array("¹", "Çàÿâêà", "ÔÈÎ", "Äàòà ðîæäåíèÿ", "Òåëåôîí", _
                                                  "Ãðàæäàíñòâî", "Ýêçàìåí", "Àóäèòîðèÿ", "Âðåìÿ", "Ðàáîòîäàòåëü")
            lastColOut = 10
        Else
            dstSheet.Range("A1:I1").Value = Array("¹", "Çàÿâêà", "ÔÈÎ", "Äàòà ðîæäåíèÿ", "Òåëåôîí", _
                                                  "Ãðàæäàíñòâî", "Ýêçàìåí", "Àóäèòîðèÿ", "Âðåìÿ")
            lastColOut = 9
        End If
    Else
        If employersPresent Then
            dstSheet.Range("A1:K1").Value = Array("¹", "Çàÿâêà", "ÔÈÎ", "Äàòà ðîæäåíèÿ", "Òåëåôîí", _
                                                  "Ãðàæäàíñòâî", "Ýêçàìåí", "Àóäèòîðèÿ", "Âðåìÿ", "Ãðóïïà", "Ðàáîòîäàòåëü")
            lastColOut = 11
        Else
            dstSheet.Range("A1:J1").Value = Array("¹", "Çàÿâêà", "ÔÈÎ", "Äàòà ðîæäåíèÿ", "Òåëåôîí", _
                                                  "Ãðàæäàíñòâî", "Ýêçàìåí", "Àóäèòîðèÿ", "Âðåìÿ", "Ãðóïïà")
            lastColOut = 10
        End If
    End If

    ' Ïåðåíîñ äàííûõ ñ çàãîëîâêàìè/ïîäçàãîëîâêàìè
    dstRow = 2
    previousGroup = ""
    previousSubGroup = ""
    numInGroup = 1

    For i = 2 To tempRow - 1
        currentGroup = Trim(CStr(tempSheet.Cells(i, 6).Value)) & " " & Format(tempSheet.Cells(i, 8).Value, "hh:mm")
        currentSubGroup = Trim(CStr(tempSheet.Cells(i, 9).Value))

        ' Çàãîëîâîê ñåêöèè
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

        ' Ïîäçàãîëîâîê ãðóïïû (åñëè ãðóïï > 1)
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

        ' Ñòðîêà ñòóäåíòà (ñ íîìåðîì)
        dstSheet.Cells(dstRow, 1).Value = numInGroup

        If groupsAlwaysOne Then
    dateColIndex = 4
    timeColIndex = 9
    phoneColIndex = 5

    removeDateColumn = hasTestDate And singleTestDate
    removeTimeColumn = hasTestTime And singleTestTime

    If removeDateColumn Then
        dstSheet.Columns(dateColIndex).Delete
        lastColOut = lastColOut - 1
        If timeColIndex > dateColIndex Then timeColIndex = timeColIndex - 1
        If phoneColIndex > dateColIndex Then phoneColIndex = phoneColIndex - 1
    End If

    If removeTimeColumn Then
        dstSheet.Columns(timeColIndex).Delete
        lastColOut = lastColOut - 1
    End If

    dstSheet.Columns("A").HorizontalAlignment = xlCenter
    dstSheet.Range(dstSheet.Cells(1, 1), dstSheet.Cells(1, lastColOut)).Interior.Color = RGB(200, 200, 200)

    If Not removeDateColumn Then
        dstSheet.Columns(dateColIndex).NumberFormat = "dd.mm.yyyy"
    End If
    If Not removeTimeColumn Then
        dstSheet.Columns(timeColIndex).NumberFormat = "hh:mm"
    End If
    dstSheet.Columns(phoneColIndex).NumberFormat = "@"
    headerText = "  "
    If dictExam.Count = 1 Then
        onlyExam = Trim$(CStr(dictExam.Keys()(0)))
        If Len(onlyExam) > 0 Then headerText = headerText & " " & UCase$(onlyExam)
    End If

        If removeDateColumn Then
            headerText = headerText & "  " & Format$(minTestDate, "dd.mm.yyyy")
            headerText = headerText & "  " & Format$(minTestDate, "dd.mm.yyyy") & " - " & Format$(maxTestDate, "dd.mm.yyyy")
    If removeTimeColumn Then
        headerText = headerText & "  " & Format$(minTestTime, "hh:mm")

    Dim titleOverride As String
    titleOverride = headerText

    ' :
    AddScheduleHeader dstSheet, lastColOut, dstRow - 1, , , titleOverride
    dstSheet.Columns("A").HorizontalAlignment = xlCenter ' íîìåðà ïî öåíòðó
    dstSheet.Columns("E").NumberFormat = "@"            ' Òåëåôîí êàê òåêñò

    ' Ãðàíèöû ïî âñåé èòîãîâîé òàáëèöå
    With dstSheet.Range(dstSheet.Cells(1, 1), dstSheet.Cells(dstRow - 1, lastColOut)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With

    
    ' Ïîäãîòîâèì titleOverride è óäàëèì ñòîëáåö Ýêçàìåí, åñëè îí îäèí
    Dim titleOverride As String
    Dim datePart As String
    If hasTestDate Then
        If minTestDate <> maxTestDate Then
            datePart = " ÍÀ " & Format(minTestDate, "dd.mm.yyyy") & " - " & Format(maxTestDate, "dd.mm.yyyy")
        Else
            datePart = " ÍÀ " & Format(minTestDate, "dd.mm.yyyy")
        End If
    Else
        datePart = ""
    End If

    Dim exColHeader As Long: exColHeader = 0
    Dim c As Long
    For c = 1 To lastColOut
        If Trim$(CStr(dstSheet.Cells(1, c).Value)) = "Ýêçàìåí" Then exColHeader = c: Exit For
    Next c

    If dictExam.Count = 1 Then
        onlyExam = dictExam.Keys()(0)
        titleOverride = "ÐÀÑÏÈÑÀÍÈÅ ÍÀ ÝÊÇÀÌÅÍ " & UCase$(onlyExam) & datePart
        If exColHeader > 0 Then
            dstSheet.Columns(exColHeader).Delete
            lastColOut = lastColOut - 1
        End If
    Else
        titleOverride = "ÐÀÑÏÈÑÀÍÈÅ" & datePart
    End If
app.ScreenUpdating = True
    ' Äîáàâëåíî: øàïêà ðàñïèñàíèÿ
    On Error Resume Next
    Call AddScheduleHeader(dstSheet, lastColOut, dstRow - 1, IIf(hasTestDate, minTestDate, Empty), IIf(hasTestDate, maxTestDate, Empty), titleOverride), IIf(hasTestDate, maxTestDate, Empty), titleOverride)
    On Error GoTo 0
    MsgBox "Ðàñïèñàíèå ãîòîâî"
End Sub

' === Äîáàâëåíî: Øàïêà "ÐÀÑÏÈÑÀÍÈÅ ÍÀ ÄÄ.ÌÌ.ÃÃÃÃ" ===
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
    
    
    ' Ïðè íàëè÷èè çàðàíåå ñîáðàííîãî çàãîëîâêà — èñïîëüçóåì åãî è íå ñêàíèðóåì òàáëèöó
    If Len(titleOverride) > 0 Then
        ' Âñòàâèì ñòðîêó ñâåðõó
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

' Íàéä¸ì êîëîíêè ïî çàãîëîâêàì
    For c = 1 To lastColOut
        If Trim$(CStr(dstSheet.Cells(findRow, c).Value)) = "Äàòà òåñòèðîâàíèÿ" Then dtCol = c
        If Trim$(CStr(dstSheet.Cells(findRow, c).Value)) = "Ýêçàìåí" Then exCol = c
    Next c
    If dtCol = 0 Then
        findRow = 2
        For c = 1 To lastColOut
            If Trim$(CStr(dstSheet.Cells(findRow, c).Value)) = "Äàòà òåñòèðîâàíèÿ" Then dtCol = c
            If Trim$(CStr(dstSheet.Cells(findRow, c).Value)) = "Ýêçàìåí" Then exCol = c
        Next c
    End If
    
    ' Åñëè äàòû ïåðåäàíû ïàðàìåòðàìè — èñïîëüçóåì èõ
    If IsDate(dtMin) And IsDate(dtMax) Then
        minDate = CDate(dtMin)
        maxDate = CDate(dtMax)
        hasDate = True
    End If

    ' Ñîáåð¸ì äèàïàçîí äàò è ìíîæåñòâî ýêçàìåíîâ
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
    
    ' Åñëè ýêçàìåí îäèí — óäàëèì ñòîëáåö "Ýêçàìåí"
    Dim titleExamPart As String: titleExamPart = ""
    If dictEx.Count = 1 And exCol > 0 Then
        Dim onlyExam As String
        onlyExam = dictEx.Keys()(0)
        titleExamPart = " ÍÀ ÝÊÇÀÌÅÍ " & UCase$(onlyExam)
        ' Óäàëÿåì ñòîëáåö è êîððåêòèðóåì øèðèíó
        dstSheet.Columns(exCol).Delete
        lastColOut = lastColOut - 1
        If dtCol > exCol Then dtCol = dtCol - 1 ' åñëè Äàòà ïðàâåå Ýêçàìåíà, å¸ èíäåêñ ñìåñòèòñÿ
    End If
    
    ' Ñôîðìèðóåì çàãîëîâîê: åñëè íåñêîëüêî äàò -> äèàïàçîí
    Dim titleText As String
    If hasDate Then
        If minDate <> maxDate Then
            titleText = "ÐÀÑÏÈÑÀÍÈÅ" & titleExamPart & " ÍÀ " & Format(minDate, "dd.mm.yyyy") & " - " & Format(maxDate, "dd.mm.yyyy")
        Else
            titleText = "ÐÀÑÏÈÑÀÍÈÅ" & titleExamPart & " ÍÀ " & Format(minDate, "dd.mm.yyyy")
        End If
    Else
        titleText = "ÐÀÑÏÈÑÀÍÈÅ" & titleExamPart
    End If
    
    ' Âñòàâèì ñòðîêó ñâåðõó
    dstSheet.Rows(headerRow).Insert Shift:=xlDown
    ' Îáúåäèíÿåì íà âñþ øèðèíó òàáëèöû
    With dstSheet.Range(dstSheet.Cells(headerRow, 1), dstSheet.Cells(headerRow, lastColOut))
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .RowHeight = 24
        ' Öâåò: ÷óòü òåìíåå, ÷åì íèæíÿÿ ñòðîêà
        Dim belowColor As Long
        belowColor = dstSheet.Range(dstSheet.Cells(headerRow + 1, 1), dstSheet.Cells(headerRow + 1, lastColOut)).Interior.Color
        If belowColor = 0 Then belowColor = RGB(200, 200, 200) ' çàïàñíîé âàðèàíò
        .Interior.Color = DarkenColor(belowColor, 0.85)
        ' Ãðàíèöû âîêðóã îáúåäèí¸ííîé ÿ÷åéêè
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End With
    
    Exit Sub
AddHeaderFail:
    ' Åñëè ÷òî-òî ïîøëî íå òàê, òèõî ïðîïóñêàåì
End Sub