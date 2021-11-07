# Refactoring VBA Code in Excel

# Overview of the Project
A friend asked for help with a stock analysis for his parents. He provided a data set for two years and twelve green stocks. Although he has some Microsoft Excel skills, he wants some help using Visual Basic for Applications (VBA) to compile annual volume and return for the tweleve stocks. After reviewing the analysis with the new script, he decided he wants to analyze larger amounts of data from the stock market. Although the script created works for a small data set, it is not designed for large volumes of data and calculations. Instead of trying to write a new script, refactoring the code to accommodate larger data sets, and improve system usage at the same time was determined to be a better option.

# Results
Utilizing the code initially developed for my friend's request, I built an array, or list, of the twelve tickers in the data. This array allows me to scan the list and, if the array id equals the row of data in the sheet, displays the ticker, sums the total annual volume, and calculates the return using the end of year price divided by the start of year price as a percentage. This code looks at the list, determines if the row data in the sheet equals that list row, and performs the necessary calculations.

## Single array For Loop with Nested Loops
    For i = 0 To 11
      ticker = tickers(i)
      totalVolume = 0
    
    Sheets(yearValue).Activate
    
        For j = 2 To RowCount
        
         '5a) get total volume for current ticker
                
             If Cells(j, 1).Value = ticker Then
                 
                 totalVolume = totalVolume + Cells(j, 8).Value
                 
             End If
             
             If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
             
                 startingPrice = Cells(j, 6).Value
             
             End If
             
             If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
             
                 'set ending price
                 endingPrice = Cells(j, 6).Value
            
            End If
            
         Next j
    '6a) output results
    
    Next i
 
## By creating an index variable to count the number of rows for each ticker array, and three new arrays, we are able to reduce the For Loop/ Nested Loop codes to a simpler, more efficient script:
        For tickerIndex = LBound(tickers) To UBound(tickers)
          ticker = tickers(tickerIndex)
          tickerVolumes(tickerIndex) = 0
         
      Worksheets(yearValue).Activate
    
        For i = 2 To RowCount
    
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
          If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
         
        Next i
    Next tickerIndex

# Summary

1. Advantages and Disadvantages of Refactoring

2. Pros and Cons of Refactoring VBA Script


