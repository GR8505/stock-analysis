# Challenge

**_______________________________________________________________________________________________________________________________________**

## Objectives

1) Analyze code and explain how it accomplishes a task.
2) Evaluate logic flow of for loops and conditionals.
3) Extend your pattern recognition skills to refactor existing code.

________________________________________________________________________________________________________________________________________
### Original Code ###

Sub AllStocksAnyYear()


yearValue = InputBox("What year would you like to run the analysis on?")
    
    ***Creating headings for AllStocksAnalysis spreadsheet***
    Worksheets("AllStocksAnalysis").Activate
    Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
    Cells(4, 1).Value = "Ticker"
    Cells(4, 2).Value = "Total Stock Volume"
    Cells(4, 3).Value = "Return"
        
        
        'Assigning tickers for each stock
        Dim tickers(12) As String
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
        
        
        'Creating starting price and ending price variables
        Dim startingPrice As Double
        Dim endingPrice As Double
        
        
        'Setting starting row and ending row for spreadsheet
        Worksheets(yearValue).Activate
        rowStart = 2
        rowEnd = Range("A" & Rows.Count).End(xlUp).Row
        
        
        'Creating loop to calculate the total volume traded, starting price and ending price of all 12 stocks
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
                
                
                'Creating loop to calculate total volume traded for each stock in the year
                For j = rowStart To rowEnd
                    Worksheets(yearValue).Activate
                    If Cells(j, 1) = ticker Then
                    totalVolume = totalVolume + Cells(j, 8)
                    End If
                
                Next j
                    
                    
                    'Creating loop to calculate starting price and ending price of each stock in the year
                    For k = rowStart To rowEnd
            
                        If Cells(k, 1) = ticker And Cells(k - 1, 1) <> ticker Then
                        startingPrice = Cells(k, 6).Value
                    
                    End If
                
                        If Cells(k, 1) = ticker And Cells(k + 1, 1) <> ticker Then
                        endingPrice = Cells(k, 6).Value
                    
                    End If
                
                    Next k
                    
                    
                    'Output Data into AllStocksAnalysis spreadsheet
                    Worksheets("AllStocksAnalysis").Activate
                    Cells(i + 5, 1).Value = ticker
                    Cells(i + 5, 2).Value = totalVolume
                    Cells(i + 5, 3).Value = endingPrice / startingPrice - 1
        Next i
        
                
        ' Visual Formatting
        Worksheets("AllStocksAnalysis").Activate
        Range("A4:C4").Font.Bold = True
        Range("A4:C4").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B5:B16").NumberFormat = "#,##0"
        Range("C5:C16").NumberFormat = "0.00%"
        Columns("B").AutoFit
        
        
        'Conditional Formatting
        stockStart = 5
        stockEnd = 16
        
        
                For i = stockStart To stockEnd
                    
                    If Cells(i, 3) > 0 Then
                    Cells(i, 3).Interior.Color = vbGreen
                
                    ElseIf Cells(i, 3) < 0 Then
                    Cells(i, 3).Interior.Color = vbRed
                    
                    Else
                    Cells(i, 3).Interior.Color = xlNone
                    
                    End If
                
                Next i



End Sub
