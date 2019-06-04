Sub MergeExcelFiles()
    Call MergeExcelBooks

    Dim outputWorkSheetName As String
    'Defining Constant'
    outputWorkSheetName = "combined output"
    'FUNCTION CALLED HERE'
    MsgBox("Combining Sheets")
    Call CreateOputFile(outputWorkSheetName)

    'FUNCTION CALLED HERER'
    Call CreateHeaderForOutputFile(outputWorkSheetName)

    'PHASE 2 COPY DATA OVER TO THE OUTPUT SHEET WITH ALL THE HEADER'
    Call CopyAndPasteData(outputWorkSheetName)

End Sub

Function DeleteEmptyCol(currentWorksheet As Worksheet) As Boolean
    Dim totalColumns As Integer
    Dim i As Integer
    Dim dataPointInColumn As Integer

    totalColumns = Range("A1").SpecialCells(xlCellTypeLastCell).Column
    For i = 1 To totalColumns
        dataPointInColumn = currentWorksheet.Application.WorksheetFunction.CountA(Range(Columns(i).Address))
        If dataPointInColumn = 1 Then
            Columns(i).EntireColumn.Delete
            i = i - 1
            totalColumns = totalColumns - 1
        End If
    Next i
    DeleteEmptyCol = True
End Function

Function MergeExcelBooks() As Boolean
        Dim fnameList, fnameCurFile As Variant
    Dim countFiles, countSheets As Integer
    Dim wksCurSheet As Worksheet
    Dim wbkCurBook, wbkSrcBook As Workbook

    fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Choose Excel files to merge", MultiSelect:=True)

    If (vbBoolean <> VarType(fnameList)) Then

        If (UBound(fnameList) > 0) Then
            countFiles = 0
            countSheets = 0

            Application.ScreenUpdating = False

            Set wbkCurBook = ActiveWorkbook

            For Each fnameCurFile In fnameList
                countFiles = countFiles + 1

                Set wbkSrcBook = Workbooks.Open(Filename:=fnameCurFile)

                For Each wksCurSheet In wbkSrcBook.Sheets
                    countSheets = countSheets + 1
                    wksCurSheet.Copy after:=wbkCurBook.Sheets(wbkCurBook.Sheets.Count)
                Next

                wbkSrcBook.Close SaveChanges:=False

            Next

            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic

            Dim currentWorksheet As Worksheet
            For Each currentWorksheet In ActiveWorkbook.Worksheets
                currentWorksheet.Select
                Call DeleteEmptyCol(currentWorksheet)
            Next currentWorksheet
            MsgBox "Processed " & countFiles & " files" & vbCrLf & "Merged " & countSheets & " worksheets", Title:="Merge Excel files"
        End If

    Else
        MsgBox "No files selected", Title:="Merge Excel files"
    End If
    MergeExcelBooks = True
End Function

Function CreateOputFile(outputWorkSheetName As String) As Boolean
    Dim containOutputWorkSheet As Boolean
    Dim currentWorkSheet As Worksheet
    containOutputWorkSheet = False
    'First find out if the output file exit in the booklet'
    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.Name = outputWorkSheetName Then
            containOutputWorkSheet = True
        End If
    Next currentWorkSheet

    'Next if output sheet does not exists create one'
    If containOutputWorkSheet = False Then
        ActiveWorkbook.Worksheets.Add
        ActiveSheet.Name = outputWorkSheetName
    End If

    CreateOutputFile = containOutputWorkSheet
End Function

Function CreateHeaderForOutputFile(outputWorkSheetName As String) As Boolean
    Dim currentWorkSheet As Worksheet
    Dim containSameHeader As Boolean
    Dim counter As Integer, currentSheetCount As Integer
    Dim finalHeaderList As New Collection
    'Defining Constant'
    currentSheetCount = 0
    containSameHeader = False
    For Each currentWorkSheet In ActiveWorkbook.Worksheets
        If currentWorkSheet.Name <> outputWorkSheetName Then
            currentWorkSheet.Select
            currentSheetCount = currentSheetCount + 1
            Dim currentSheetCol As Integer
            currentSheetCol = currentWorkSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Column

            If currentSheetCount = 1 Then
                For counter = 1 To currentSheetCol
                    finalHeaderList.Add (Cells(1, counter).Value)
                Next counter
            Else
                For counter = 1 To currentSheetCol
                    Dim i As Integer
                    For i = 1 To finalHeaderList.Count
                        If Cells(1, counter).Value = finalHeaderList(i) Then
                            containSameHeader = True
                            Exit For
                        End If
                    Next i
                    If containSameHeader = False Then
                        finalHeaderList.Add (Cells(1, counter).Value)
                    Else
                        containSameHeader = False
                    End If
                Next counter

            End If
        End If
    Next currentWorkSheet
    'Select the output sheet and wirte all the unique header into the output sheet'
    Sheets(outputWorkSheetName).Select
    For counter = 1 To finalHeaderList.Count
        Cells(1, counter).Value = finalHeaderList(counter)
    Next counter
End Function

Function CopyAndPasteData(outputWorkSheetName As String) As Boolean
Dim currentWorkSheet As Worksheet, outputSheet As Worksheet
Dim i As Integer, totalRowsInCurrentSheet As Integer, totalColumnsInCurrentSheet As Integer, pastePosition  As Integer
Dim pasteColumnIndex As Integer, totalColumnInOutputSheet As Integer
Dim headerName As String
'start pasting at the second row'
pastePosition = 2
For Each currentWorkSheet In ActiveWorkbook.Worksheets
    If StrComp(currentWorkSheet.Name, outputWorkSheetName, vbTextCompare) = 1 Then
        currentWorkSheet.Select
        totalColumnsInCurrentSheet = Range("A1").SpecialCells(xlCellTypeLastCell).Column
        totalRowsInCurrentSheet = Range("A1").SpecialCells(xlCellTypeLastCell).Row
        For i = 1 To totalColumnsInCurrentSheet
            'copy the whole column'
            currentWorkSheet.Range(Cells(2, i), Cells(totalRowsInCurrentSheet, i)).Copy
            headerName = Cells(1, i).Value
            Worksheets(outputWorkSheetName).Select
            Set outputSheet = ActiveSheet
            totalColumnInOutputSheet = Range("A1").SpecialCells(xlCellTypeLastCell).Column
            pasteColumnIndex = outputSheet.Range(Cells(1, 1), Cells(1, totalColumnInOutputSheet)).Find(headerName).Column
            Cells(pastePosition, pasteColumnIndex).Select
            ActiveSheet.Paste
            currentWorkSheet.Select
        Next i
        pastePosition = pastePosition + totalRowsInCurrentSheet - 1
    End If
Next currentWorkSheet
    CopyAndPasteData = True
End Function
