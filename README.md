# Challenge

**_________________________________________________________________________________________________________________________________________________________**

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

Based on the 2017 results, the three stocks with the highest returns were:
1) DQ (199.45%)
2) SEDG (184.47%)
3) ENPH (129.52%)

Out of these three, SEDG had the highest level of activity in the stock market with a total volume traded of over 2 billion
(4th highest).


![](https://github.com/GR8505/stock-analysis/blob/master/2017Output.png)


However, when we look at the 2018 results, only two stocks recorded positive returns:
1) RUN (83.95%)
2) ENPH (81.92%)

![](https://github.com/GR8505/stock-analysis/blob/master/2018Output.png)


So I will advise an investor to go for either ENPH or RUN instead of DQ.
### Reasons ###
1) Although ENPH registered a decline in return from 129.52% in 2017 to 81.92% in 2018, total volume traded remains healthy.
   Total volume traded actually increased by 41.9%.
2) RUN recorded a significant increase in return from around 5.5% in 2017 to just under 84% in 2018.  Furthermore, stock activity
   improved with total volume traded moving from just under 2 billion in 2017 to around 2.2 billion in 2018.
   
   
**________________________________________________________________________________________________________________________________________________________**





        




