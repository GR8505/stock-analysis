# Challenge

**____________________________________________________________________________________________________________________________________**

## Objectives

1) Analyze code and explain how it accomplishes a task.
2) Evaluate logic flow of for loops and conditionals.
3) Extend your pattern recognition skills to refactor existing code.

________________________________________________________________________________________________________________________________________
### Original Code ###

Sub AllStocksAnyYear()


yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Creating headings for AllStocksAnalysis spreadsheet
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

______________________________________________________________________________________________________________________________________

### Refactored Code ###

* Listing all 12 stocks can be tedious and what if, there were 40 or 50 stocks in one year.  Searching all the names of these stocks and assigning a separate ticker will be rather time consuming.
* Used code to copy column A and paste only unique values to column M.
* From this point I created a loop to run through each unique stock in column M and assigned it as a string variable named ticker.
* If the ticker is equal to any cell in column A, it runs through the nested loops we created to calculate total volume traded and return for each stock.

_______________________________________________________________________________________________________________________________________

### This is what the refactored code looks like ###

Sub Challenge2()

yearValue = InputBox("What year would you like to run the analysis on?")
    
        'Creating headings for AllStocksAnalysis spreadsheet
        Worksheets("AllStocksAnalysis").Activate
        Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
        Cells(4, 1).Value = "Ticker"
        Cells(4, 2).Value = "Total Stock Volume"
        Cells(4, 3).Value = "Return"
    
    
        'Code to copy unique values to an empty column taken from 
        https://superuser.com/questions/884115/macro-to-copy-distinct-values-from-one-excel-sheet-to-another and               
        https://www.ozgrid.com/forum/index.php?thread/141957-copy-unique-values-from-a-list/
        Worksheets(yearValue).Activate
        Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("M1"), Unique:=True
        
        
        'Assigning variables to use in our nested loops
        
        tickerStart = 2
        
        tickerEnd = Range("M" & Rows.Count).End(xlUp).Row
        
        rowStart = 2
        
        rowEnd = Range("A" & Rows.Count).End(xlUp).Row
        
        Dim ticker As String
        
        totalVolume = 0
        
        'Creating starting price and ending price variables which is needed to calculate the return
        Dim startingPrice As Single
        Dim endingPrice As Single
        
        
        'Creating nested loop to calculate the total volume traded and return for each stock for the year
            For i = tickerStart To tickerEnd
                Worksheets(yearValue).Activate
                ticker = Cells(i, 13)
                
                 'Creating loop to calculate total volume traded for each stock in the year
                For j = rowStart To rowEnd
                    
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
                    
                    
                    'Output for total volume traded and return
                    Worksheets("AllStocksAnalysis").Activate
                    Cells(i + 3, 1).Value = ticker
                    Cells(i + 3, 2).Value = totalVolume
                    Cells(i + 3, 3).Value = endingPrice / startingPrice - 1
            
            Next i
               
               
        'Clear contents of ticker column in worksheet
        Worksheets(yearValue).Activate
        Range("M:M").Clear
                
        'Visual Formatting
        Worksheets("AllStocksAnalysis").Activate
        Range("A4:C4").Font.Bold = True
        Range("A4:C4").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B:B").NumberFormat = "#,##0"
        Range("C:C").NumberFormat = "0.00%"
        Columns("B").AutoFit
        
        
        'Conditional Formatting
        
        stockStart = 5
        stockEnd = Range("A" & Rows.Count).End(xlUp).Row
        
                
                'Creating loop to highlight increasing stocks as green and declining stocks as red
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

_______________________________________________________________________________________________________________________________________

## Conclusion ##




        




