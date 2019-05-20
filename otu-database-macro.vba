Sub Format()
    'Columns(2).EntireColumn.Delete'
    Dim totalColumns As Integer
    Dim i As Integer
    Dim dataPointInColumn As Integer

    totalColumns = Range("A1").SpecialCells(xlCellTypeLastCell).Column
    For i = 1 To totalColumns
        dataPointInColumn = Application.WorksheetFunction.CountA(Range(Columns(i).Address))
        If dataPointInColumn = 1 Then
            Columns(i).EntireColumn.Delete
            i = i - 1
            totalColumns = totalColumns - 1
        End If
    Next i
End Sub
