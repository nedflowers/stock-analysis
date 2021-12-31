Sub AllStocksAnalysisRefactored()
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
    

    '1a) Create a ticker Index and set it equal to zero
    For i = 0 To 11
       tickerIndex = tickers(i)
       
    '1b)Create Output arrays
       Dim tickervolume As Long
       Dim tickerStartingPrices As Double
       Dim tickerEndingPrices As Double
     
    '2a)Initialize tickervolume equal zero
       tickervolume = 0
       
       'Activate sheet
       Sheets(yearValue).Activate
       
    '2b)Loop through all rows to find if ticker matches
       For j = 2 To RowCount

           If Cells(j, 1).Value = tickerIndex Then
    
    '3a)Increase ticker volume and add ticker volume to current ticker
               tickervolume = tickervolume + Cells(j, 8).Value
    
        'End if
           End If
    
    '3b)If this row is first row with current ticker
           If Cells(j, 1).Value = tickerIndex And Cells(j - 1, 1).Value <> tickerIndex Then
        
        'Assign ticker starting price to corresponding variable
               tickerStartingPrices = Cells(j, 6).Value


    '3c) If current row is last row to include ticker
          ElseIf Cells(j, 1).Value = tickerIndex And Cells(j + 1, 1).Value <> tickerIndex Then
        
        'Assign ticker ending price to variable
               tickerEndingPrices = Cells(j, 6).Value
        
        'Increase ticker index if next row doesn't match current row
        Else
        
        'End if
           End If
       Next j
    
    '4) Loop through your arrays and output the Ticker, Total Daily Volume, and Return.
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickerIndex
        Cells(4 + i, 2).Value = tickervolume
        Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
          
        Next i

    
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

    For k = dataRowStart To dataRowEnd
        
        If Cells(k, 3) > 0 Then
        Cells(k, 3).Interior.Color = vbGreen
        ElseIf Cells(k, 3) < 0 Then
        Cells(k, 3).Interior.Color = vbRed
        Else
        Cells(i, 3).Interior.Color = xlNone
'End if
    End If
    
        Next k
    
    'Stop clock
    endTime = Timer
    MsgBox "This code ran in" & (endTime - startTime) & "seconds for the year" & (yearValue)




End Sub
