Attribute VB_Name = "Module5"
Sub skilldrill()

Worksheets("All Stocks Analysis").Activate

Dim i As Integer, j As Integer
Dim rowend As Integer

rowend = Cells(Rows.Count, "A").End(xlUp).Row

columnend = Cells(1, Columns.Count).End(xlToLeft).Column

For i = 1 To rowend
    For j = 1 To columnend
    Cells(i, j).Value = i + j
        Next j
Next i


Range("A1").Clear
End Sub
