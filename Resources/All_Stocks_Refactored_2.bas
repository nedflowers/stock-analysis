Attribute VB_Name = "Module7"
Sub AllStocksAnalysisRefactored2()
    'Initialize variables for starting price and ending price
    Dim startTime As Single
    Dim endTime  As Single
    
    'Input box
    yearValue = InputBox("What year would you like to run the analysis on?")
    'Start clock
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    

    
    'Initialize array of all tickers
    Dim tickers(11) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Sheets(yearValue).Activate
    
    'Get number of rows to count
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Initialize and set tickerindex = 0
    tickerIndex = 0
    
    '1b)Create Output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
     
    '2a)Initialize tickervolume equal zero
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    '2b)Loop through all rows to find if ticker matches
    For j = 2 To RowCount
        
        '3a) Increase current ticker volume
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
    
    '3b)If this row is first row with current ticker
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
        
        'Assign ticker starting price to corresponding variable
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        
        End If
        
    '3c) If current row is last row to include ticker
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        
        'Assign ticker ending price to variable
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            tickerIndex = tickerIndex + 1
            
        'Increase ticker index if next row doesn't match current row
    Else
        
        'End if
        End If
        
    Next j
    
    '4) Loop through your arrays and output the Ticker, Total Daily Volume, and Return.
    For k = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + k, 1).Value = tickers(k)
        Cells(4 + k, 2).Value = tickerVolumes(k)
        Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
          
    Next k
    'Formatting
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold Italic"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("A3:C3").Font.Size = 14
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For l = dataRowStart To dataRowEnd
        
        If Cells(l, 3) > 0 Then
        Cells(l, 3).Interior.Color = vbGreen
        ElseIf Cells(l, 3) < 0 Then
        Cells(l, 3).Interior.Color = vbRed
        Else
        Cells(l, 3).Interior.Color = xlNone
'End if
    End If
    
        Next l
    
    'Stop clock
    endTime = Timer
    MsgBox "This code ran in" & (endTime - startTime) & "seconds for the year" & (yearValue)

End Sub

