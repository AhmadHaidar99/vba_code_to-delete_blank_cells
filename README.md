# vba_code_to-delete_blank_cells

Sub DeleteBlankCells()
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim i As Long
    Dim j As Long
    
    'Get the last row and last column of the data set
    lastRow = 115610
    lastColumn = 13
    
    'Loop through each row in the data set
    For i = 2 To lastRow
        'Loop through each column in the data set
        For j = 1 To lastColumn
            'Check if the cell is blank
            If Cells(i, j) = "" Then
                'Delete the cell's contents and shift cells up
                Cells(i, j).Delete Shift:=xlUp
            End If
        Next j
    Next i
End Sub
