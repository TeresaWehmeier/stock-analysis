# Refactoring VBA Code in Excel

# Overview of the Project
A friend asked for help with a stock analysis for his parents. He provided a data set for two years and twelve green stocks. Although he has some Microsoft Excel skills, he wants some help using Visual Basic for Applications (VBA) to compile annual volume and return for the tweleve stocks. After reviewing the analysis with the new script, he decided he wants to analyze larger amounts of data from the stock market. Although the script created works for a small data set, it is not designed for large volumes of data and calculations. Instead of trying to write a new script, refactoring the code to accommodate larger data sets, and improve system usage at the same time, was determined to be a better option.

## Results
The initial script contained a single array, or list, of the twelve tickers in the data. This array scans the list and, if the array id equals the row of data in the sheet, displays the ticker, sums the total annual volume, and calculates the return using the end of year price divided by the start of year price as a percentage. This code looks at the list, determines if the row data in the sheet equals that list row, and performs the necessary calculations.

#### Single array For Loop with Nested Loops
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
 
#### By creating an index variable to count the number of rows for each ticker array, and three new arrays, we are able to reduce the For Loop/ Nested Loop code to a cleaner, readable patterned loop that is more efficient to run:

        Dim tickerIndex As Integer
        tickerIndex = 0
        
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
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
    
#### Efficiency and Improvement
One of the objectives of refactoring is to improve readability, replace hard-coded values and magic numbers, and speed up processing. Refactoring also allows the coder to improve the patterns in the script, making it easier to navigate. To measure improvement and efficiency, a timer was created to capture the time each script takes to run. The simple script was attached to both the original and final code:
    startTime = Timer; EndTime = Timer; MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
The results from the original script were captured for both sheets 2017 and 2018, with the following results:
<img src = "https://github.com/TeresaWehmeier/stock-analysis/blob/main/Images/VBA_Script_Old_2017.png" width="60%" height="40%">

<img src = "https://github.com/TeresaWehmeier/stock-analysis/blob/main/Images/VBA_Script_Old_2018.png" width="60%" height="40%">

The new script, with additional arrays, had the following results for the same 2017 and 2018 data:
<img src = "https://github.com/TeresaWehmeier/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png" width="60%" height="40%">

<img src = "https://github.com/TeresaWehmeier/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png" width="60%" height="40%">

These results reflect a 10th of a second improvement on the final project. Although that is not a staggering number for this small data set, it will make a huge difference if we are to pull hundreds of thousands or more stock results from the stock market to analyze.

## Summary

#### Advantages and Disadvantages of Refactoring
The advantages of refactoring existing code are, first and most important, to prevent the programmer from reinventing the wheel. It is much more efficient to clean up and modernize code that has already been written. The programmer's value isn't in writing new code so much as it is in having the ability to take someone else's code and make it better. In addition, refractoring can be done in small intervals, or as time allows. If the programmer has only a short amount of time to dedicate to a refactoring project each week, it is simple enough to get in and do as much cleaning as time allows, then return later to continue. As long as the code continues to work, cleaning up comments and magic numbers can improve the code a little at a time. Finally, like all technologies over time, code changes; improvements to the coding applications improve, so it is a good time to review old code and bring it up to date. These tasks, if performed infrequently, can prevent failures in the code.

Disadvantages are the obvious: scripts that were quickly written to solve a problem may contain few comments explaining the decisions made, hard-coded data embedded in the code, and work-around scripts used to navigate broken code. All can be a nightmare to fix. Often, older applications written in VBA linger far past their shelf life, and if not reviewed regularly, can result in an application that no longer functions correctly or at all.

#### Pros and Cons of Refactoring VBA Script
Rafactoring VBA script can be useful if the application it runs still has value for the end user. There are several guides on how to refactor in any programming language, because the process is the same. However, there are additional tools that are available in Visual Studio that can help with refactoring, and given the limitations of doing this task in the VBA environment makes it the one big con to refactoring in VBA. The application's limited assistance with errors was, I found, very frustrating. However, I have not yet become proficient with the tool, so that may not be an issue as I develop my skills in VBA.



