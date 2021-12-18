Attribute VB_Name = "Module1"
Sub MacroCheck()

Dim testMessage As String

testMessage = "Hello World!"

MsgBox (testMessage)

End Sub

Sub DQAnalysis()
    
    Worksheets("DQ Analysis").Activate
    
    'Create a header row
    Cells(1, 1).Value = "DAQO (Ticker: DQ)"
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"



End Sub
