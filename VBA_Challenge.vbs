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
    
    '1a) Create a ticker Index
        Dim tickerindex As String
        
    For j = 0 To 11
        
        tickerindex = tickers(j)
            
        
    '1b) Create three output arrays
        Dim tickerVolumes(11) As Long
        Dim tickerEndingPrices(11) As Single
        Dim tickerStartingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
     
        
           
                              
    ''2b) Loop over all the rows in the spreadsheet.
            For i = 2 To RowCount
                                        
        '3a) Increase volume for current ticker
                If Cells(i, 1) = tickerindex Then

                    tickerVolumes(j) = tickerVolumes(j) + Cells(i, 8).Value
         
                               
                End If
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
                       
                If Cells(i - 1, 1) <> tickerindex And Cells(i, 1) = tickerindex Then

                    tickerStartingPrices(j) = Cells(i, 6).Value
                 
                                
                End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
                   
                If Cells(i + 1, 1) <> tickerindex And Cells(i, 1) = tickerindex Then

                    tickerEndingPrices(j) = Cells(i, 6).Value
                    tickerindex = Cells(i + 1, 1).Value
                    
                Exit For
                
                End If
        
               
                    
            Next i

            '3d Increase the tickerIndex
              
                    
    Next j
       
              
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    tickerindex = 0
        
    For j = 0 To 11
    
        tickerindex = tickers(j)
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + j, 1).Value = tickerindex
        Cells(4 + j, 2).Value = tickerVolumes(j)
        Cells(4 + j, 3).Value = tickerEndingPrices(j) / tickerStartingPrices(j) - 1
        
    Next j
        

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For m = dataRowStart To dataRowEnd
        
        If Cells(m, 3) > 0 Then
            
            Cells(m, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(m, 3).Interior.Color = vbRed
            
        End If
        
    Next m
    
 
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)



End Sub