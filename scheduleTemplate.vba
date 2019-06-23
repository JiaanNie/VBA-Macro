'Global Constant'
Global templateYear As Integer
'Global Const MAXROWS As Integer = 40'
Global Const MAXCOLUMNS As Integer = 201
Global Const MAXCELLWIDTH As Integer = 4




Sub CreateScheduleTemplate()
    'Set when open a new excel workbook use the sheet1 as option sheet'
    templateYear = InputBox("Enter the year of this template that you wanted to create! Example: 2019")
    Application.DisplayStatusBar = False
    Application.ScreenUpdating = False
    Call SetOptionSheet

    Dim listOfMonths As New Collection
    'Application.SendKeys "^g ^a {DEL}"'
    Set listOfMonths = CreateMonths()
    Call CreateSheets(listOfMonths)
    Call ConfigMonthSheet
    Call FormatSheetStyle
    Call AutoBlockWeekend
    Call FormatDataCol
    Call ConfigDropDown
    Call ApplyColorToOptions
    Call PopulateNames
End Sub
'Set option sheet function'
Function SetOptionSheet() As Boolean
    ActiveSheet.name = "Options"
    Dim i As Integer, j  As Integer
    Cells(1, 1).Value = "Daily Options"
    ActiveSheet.Columns("A").ColumnWidth = ActiveSheet.Columns("A").ColumnWidth * MAXCELLWIDTH
    listOfOptions = Array("Holiday AM", "Holiday PM", "Holiday All Day", _
                          "(.5)Sick/Appointment", "Working Away", "Conference/Meeting off Campus", _
                          "Sick/Appointment", "Last Day", "No Longer Working", "Stat Holiday")
    j = 2
    For i = 0 To UBound(listOfOptions)
        Cells(j, 1).Value = listOfOptions(i)
        j = j + 1
    Next i
    SetOptionSheet = True
End Function

'Create array of months name'
Function CreateMonths() As Collection
    Dim i As Integer
    Dim listOfMonths As New Collection
    Dim month

    For i = 1 To 12
        listOfMonths.Add (MonthName(13 - i))
    Next i
    Set CreateMonths = listOfMonths
End Function


'using the array of months name to create 12 worksheet whtin the work book'
Function CreateSheets(listOfMonths As Collection) As Boolean
    Dim month
    For Each month In listOfMonths
        ActiveWorkbook.Worksheets.Add
        ActiveSheet.name = month
    Next month

End Function
'get the total numbe of days for a given month and give year Note: you can set Year(Date) to what everyear ie 2035'
Function ObtainNumberOfDays(inputMonth As Long) As Integer
    ObtainNumberOfDays = day(DateSerial(templateYear, inputMonth + 1, 1) - 1)
End Function


'Set up the days for each month'
Function ConfigMonthSheet() As Boolean
    Dim currentWorkSheet As Worksheet
    Dim numberOfDays As Integer
    Dim i As Long
    Dim dayCounter As Integer
    i = 1
    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.name <> "Options" Then
            currentWorkSheet.Select
            numberOfDays = ObtainNumberOfDays(i)
            For dayCounter = 1 To numberOfDays
                Cells(dayCounter + 1, 1).Value = dayCounter
            Next dayCounter
            i = i + 1
        End If
    Next currentWorkSheet
    ConfigMonthSheet = True
End Function

'block and black out weekend'
Function AutoBlockWeekend() As Boolean
    Dim currentWorkSheet As Worksheet
    Dim currentRow As Integer
    Dim currentCol As Integer
    Dim totalColsInCurrentSheet As Integer
    Dim totalRowsInCurrentSheet As Integer
    Dim currentDay As Integer
    Dim test As String
    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.name <> "Options" Then
            totalColsInCurrentSheet = currentWorkSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Column
            totalRowsInCurrentSheet = currentWorkSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Row
            currentWorkSheet.Select
            For currentRow = 2 To totalRowsInCurrentSheet
                currentDay = Weekday(GetDateFormat(CStr(templateYear), GetMonthNumValue(ActiveSheet.name), CStr(Cells(currentRow, 1).Value)))
                If currentDay = vbSunday Or currentDay = vbSaturday Then
                    Range(Cells(currentRow, 2), Cells(currentRow, MAXCOLUMNS)).Value = "Weekend"
                End If

            Next currentRow

        End If
    Next currentWorkSheet

End Function

'get the column address by index function'
Function GetColumnAddress(columnIndex As Long) As String
    'Convert To Column Letter
    GetColumnAddress = Split(Cells(1, columnIndex).Address, "$")(1)
End Function
'get the column number base on the letter that passing in'
Function GetColumnNumber(columnLetter As String) As Long
    GetColumnNumber = Range(columnLetter & 1).Column
End Function


'get the number value for the month in a string'
Function GetMonthNumValue(inputMonth As String) As String
    Select Case inputMonth
        Case Is = "January"
            GetMonthNumValue = "01"
        Case Is = "February"
            GetMonthNumValue = "02"
        Case Is = "March"
            GetMonthNumValue = "03"
        Case Is = "April"
            GetMonthNumValue = "04"
        Case Is = "May"
            GetMonthNumValue = "05"
        Case Is = "June"
            GetMonthNumValue = "06"
        Case Is = "July"
            GetMonthNumValue = "07"
        Case Is = "August"
            GetMonthNumValue = "08"
        Case Is = "September"
            GetMonthNumValue = "09"
        Case Is = "October"
            GetMonthNumValue = "10"
        Case Is = "November"
            GetMonthNumValue = "11"
        Case Is = "December"
            GetMonthNumValue = "12"
    End Select
End Function
'this function will return yyyy-mm--dd formate as string '
Function GetDateFormat(year As String, month As String, day As String) As String
    GetDateFormat = year + "-" + month + "-" + day
End Function


Function MergeCells(rowIndex As Integer, colIndexBegine As Integer, colIndexEnd As Integer) As Boolean
    ActiveSheet.Range(Cells(rowIndex, colIndexBegine), Cells(rowIndex, colIndexEnd)).Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    Selection.Merge
End Function

'Formating Name row and day row'
Function FormatSheetStyle() As Boolean
    Dim currentWorkSheet As Worksheet
    Dim totalRowsInCurrentSheet As Integer
    Dim rowIndex As Integer
    Dim colIndex As Integer
    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.name <> "Options" Then
            currentWorkSheet.Select
            totalRowsInCurrentSheet = currentWorkSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Row
            For rowIndex = 1 To totalRowsInCurrentSheet
                If rowIndex = 1 Then
                    For colIndex = 2 To MAXCOLUMNS Step 4
                        Call MergeCells(rowIndex, colIndex, colIndex + 3)
                        Call BorderStyle(1, 1, colIndex, colIndex + 3)
                    Next colIndex
                Else
                    For colIndex = 2 To MAXCOLUMNS Step 2
                        Call MergeCells(rowIndex, colIndex, colIndex + 1)
                    Next colIndex

                End If
            Next rowIndex

        End If
    Next currentWorkSheet
End Function

'format name border'
Function BorderStyle(startRow As Integer, endRow As Integer, startEdge As Integer, endEdge As Integer) As Boolean
    ActiveSheet.Range(Cells(startRow, startEdge), Cells(endRow, endEdge)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Function

'format '
Function FormatDataCol() As Boolean
Dim currentWorkSheet As Worksheet
    Dim totalRowsInCurrentSheet As Integer
    Dim rowIndex As Integer
    Dim colIndex As Integer
    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.name <> "Options" Then
            currentWorkSheet.Select
            totalRowsInCurrentSheet = currentWorkSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Row
            For colIndex = 2 To MAXCOLUMNS Step 4
                Call BorderStyle(2, totalRowsInCurrentSheet, colIndex, colIndex + 3)
            Next colIndex
        End If
    Next currentWorkSheet
End Function

'config drop downlist for data cells'
Function ConfigDropDown()
    Dim totalRowsInCurrentSheet As Integer
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim colChar As String
    Dim customFormula As String
    Dim isA As Boolean
    isA = True
    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.name <> "Options" Then
            currentWorkSheet.Select
            totalRowsInCurrentSheet = currentWorkSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Row - 1
            For colIndex = 2 To MAXCOLUMNS Step 4
                colChar = "A"
                customFormula = "=Options!$" + colChar + "$2:$" + colChar + "$100"
                Call ConfigDropDownHelper(colIndex, 2, totalRowsInCurrentSheet, customFormula)
            Next colIndex
        End If
    Next currentWorkSheet
End Function

Function ConfigDropDownHelper(colIndex As Integer, startRow As Integer, endRow As Integer, customFormula As String) As Boolean
    ActiveSheet.Range(Cells(startRow, colIndex), Cells(endRow, colIndex)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=customFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Function



Function ApplyColorToOptions()
    Dim currentWorkSheet As Worksheet
    Dim totalOptions As Integer, i As Integer
    Dim cutomeFormula As String
    totalOptions = Sheets("Options").Range("A1").SpecialCells(xlCellTypeLastCell).Row - 1
    For i = 2 To totalOptions
        cutomeFormula = "=Options!$A$" + CStr(i)
        ApplyColorToOptionHelper (cutomeFormula)
    Next i
End Function


Function ApplyColorToOptionHelper(cutomeFormula As String)
    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.name <> "Options" Then
            currentWorkSheet.Select
            Range("B2:GS32").Select
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:=cutomeFormula
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .Color = Int(200000 * Rnd) + 100000
            End With
        End If
    Next currentWorkSheet

End Function

'SETTING FOR TIME and Name'

Function PopulateNames()
    Dim currentWorkSheet As Worksheet
    Dim listOfNames
    Dim nameSize As Integer, i As Integer, j As Integer
    listOfNames = Array("Bob", "John", "Eve")
    nameSize = UBound(listOfNames)

    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.name <> "Options" Then
            j = 2
            currentWorkSheet.Select
            For i = 0 To nameSize
                If i = 0 Then
                    Cells(1, j).Value = listOfNames(i)
                Else
                    Cells(1, j).Value = listOfNames(i)
                End If
                j = j + 4
            Next i
        End If
    Next currentWorkSheet
    Call AutoFillTime(listOfNames)
End Function





Function AutoFillTime(listOfNames)
    Dim currentWorkSheet As Worksheet
    Dim i As Integer, size As Integer
    size = UBound(listOfNames) + 1 'because it started at column 2'
    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.name <> "Options" Then
            currentWorkSheet.Select
            For i = 2 To size * 4 Step 4
                Call AutoFillTimeHelper(Cells(1, i).Value, i)
            Next i

        End If
    Next currentWorkSheet
End Function

Function AutoFillTimeHelper(name As String, columnNumber As Integer)
    Dim defaultWorkTime As String, workingDays As String
    Select Case name
            Case Is = "Bob"
                defaultWorkTime = "8:30-4:30"
                workingDays = "M,T,W,Th,F"
                Call FillingTIme(defaultWorkTime, workingDays, columnNumber + 2)

            Case Is = "John"
                defaultWorkTime = "8:30-4:30"
                workingDays = "T,Th"
                Call FillingTIme(defaultWorkTime, workingDays, columnNumber + 2)

            Case Is = "Eve"
                defaultWorkTime = "8:00-12:00"
                workingDays = "M,W,F"
                Call FillingTIme(defaultWorkTime, workingDays, columnNumber + 2)
                defaultWorkTime = "8:00-4:30"
                workingDays = "T,Th"
                Call FillingTIme(defaultWorkTime, workingDays, columnNumber + 2)
        End Select
End Function

Function FillingTIme(time As String, workingDays As String, columneToFIll As Integer)
    Dim workDay() As String
    Dim i As Integer, rowIndex As Integer, currentDay As Integer
    Dim M As Boolean, T As Boolean, W As Boolean, Th As Boolean, F As Boolean
    M = False
    T = False
    W = False
    Th = False
    F = False
        workDay = Split(workingDays, ",")
    For i = 0 To UBound(workDay)
        If workDay(i) = "M" Then
            M = True
        ElseIf workDay(i) = "T" Then
            T = True
        ElseIf workDay(i) = "W" Then
            W = True
        ElseIf workDay(i) = "Th" Then
            Th = True
        ElseIf workDay(i) = "F" Then
            F = True
        End If
    Next i


    Dim totalRowsInCurrentSheet As Integer
    totalRowsInCurrentSheet = ActiveSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Row - 1
    Debug.Print totalRowsInCurrentSheet
    For rowIndex = 2 To totalRowsInCurrentSheet
        Debug.Print rowIndex
        currentDay = Weekday(GetDateFormat(CStr(templateYear), GetMonthNumValue(ActiveSheet.name), CStr(Cells(rowIndex, 1).Value)))
        If currentDay = vbMonday And M = True Then
            Cells(rowIndex, columneToFIll).Value = time
        ElseIf currentDay = vbTuesday And T = True Then
            Cells(rowIndex, columneToFIll).Value = time
        ElseIf currentDay = vbWednesday And W = True Then
            Cells(rowIndex, columneToFIll).Value = time
        ElseIf currentDay = vbThursday And Th = True Then
            Cells(rowIndex, columneToFIll).Value = time
        ElseIf currentDay = vbFriday And F = True Then
            Cells(rowIndex, columneToFIll).Value = time
        End If

    Next rowIndex

End Function
