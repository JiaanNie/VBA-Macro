Sub Format()
    'Columns(2).EntireColumn.Delete'
    Dim totalRows As Integer
    Dim totalColoumns As Integer
    Dim i As Integer
    Dim j As Integer
    Dim coloumnContainData As Boolean
    coloumnContainData = False
    totalRows = Range("A1").SpecialCells(xlCellTypeLastCell).Row
    totalColoumns = Range("A1").SpecialCells(xlCellTypeLastCell).Column
    For i = 1 To totalColoumns
        For j = 2 To totalRows
            If IsEmpty(Cells(j, i).Value) = False Then
                coloumnContainData = True
                Exit For
            End If
        Next j
        If coloumnContainData = False Then
            MsgBox (CStr(i))
            Cells(2, i).Value = "DELETTHISCOLOUM"
        End If
        coloumnContainData = False
    Next i
    For i = 1 To totalColoumns
        If Cells(2, i).Value = "DELETTHISCOLOUM" Then
            Columns(i).EntireColumn.Delete
            i = i - 1
            totalColoumns = totalColoumns - 1
        End If
    Next i
End Sub
