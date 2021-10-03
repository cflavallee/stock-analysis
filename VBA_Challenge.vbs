Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

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
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create a ticker Index
    'Loop through tickers
    
    Dim tickerindex As String
        
    For j = 0 To 11
        
        tickerindex = tickers(j)
            
        
    'Create three output arrays to store values as script loops
        Dim tickerVolumes(11) As Long
        Dim tickerEndingPrices(11) As Single
        Dim tickerStartingPrices(11) As Single
    
    'Loop over all the rows in the spreadsheet.
            For i = 2 To RowCount
                                        
        'Increase volume for current ticker and store in tickerVolumes array.
                If Cells(i, 1) = tickerindex Then

                    tickerVolumes(j) = tickerVolumes(j) + Cells(i, 8).Value
         
                               
                End If
                
        'Find starting price for each ticker.
                       
                If Cells(i - 1, 1) <> tickerindex And Cells(i, 1) = tickerindex Then

                    tickerStartingPrices(j) = Cells(i, 6).Value
                 
                                
                End If
        
        'Find ending price for each tiker.
                   
                If Cells(i + 1, 1) <> tickerindex And Cells(i, 1) = tickerindex Then

                    tickerEndingPrices(j) = Cells(i, 6).Value
                    tickerindex = Cells(i + 1, 1).Value
                    
                'Exit for loop after ending price has been stored.
                Exit For
                
                End If
        
               
            'Restart i loop.
        Next i
              
    'Move on to next ticker.
    Next j
       
              
    
    'After all data is stored, loop through arrays to output the Ticker, Total Daily Volume, and Return.
    tickerindex = 0
        
    For j = 0 To 11
    
        tickerindex = tickers(j)
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + j, 1).Value = tickerindex
        Cells(4 + j, 2).Value = tickerVolumes(j)
        Cells(4 + j, 3).Value = tickerEndingPrices(j) / tickerStartingPrices(j) - 1
        
    Next j
        

    'Format output data.
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    
    'Declare variables for output range.
    dataRowStart = 4
    dataRowEnd = 15
    
    'Color code cells based on positive or negative return.
    For m = dataRowStart To dataRowEnd
        
        If Cells(m, 3) > 0 Then
            
            Cells(m, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(m, 3).Interior.Color = vbRed
            
        End If
        
    Next m
    
 
    'Output total run time in message box.
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)



End Sub
