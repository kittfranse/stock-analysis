# Stock Performance Analyzer
VBA-based stock performance analysis

### Overall Objective
The purpose of this analysis was to automate the analysis of a large database of stock information for the client.  Initially, a database of stocks and their daily performance at open, the high, low, and close were all provided in a dense list.  I wrote code to automate the summary of this data by analyzing stocks as a whole (by their name or ticker) and reducing the presentation to overall performance of the stock for the entire year of 2017 and 2018.

### Results
Two scripts were generated in the goal of analyzing this data as aforementioned. The first script went line by line for every single entry provided in the database (shown below). As a result, run times were approximately 0.8-0.9 seconds for both groups of stock data.  To mitigate this, arrays were employed to group the data so that collection of information was a run overall (also shown below. This showed a nearly 10 fold decrease in the amount of time it took for the script to run (see \Resources\VBA_Challenge_2017 and ..\VBA_Challenge_2018 for the new run times)

#### Line-by-line Analysis
`
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
           
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If

       Next j
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i`

#### Refactoring Using Arrays
`
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
        
    For i = 2 To RowCount
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
       
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

    
    Next i
 `

![VBA_Challenge_2017](https://user-images.githubusercontent.com/59892063/133951976-1cb482a0-6075-467f-a8ad-4fe36391c5fb.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/59892063/133951964-43fde59d-fec9-45c7-9731-d00ba7b0cda0.png)

### Summary

I find a disadvantage for both codes is that the tickers are hardcoded into the code. It would be much better if there was a scan for unique values in the first ticker column that was then used to generate the arrays in the refactored code.  The advantage of re-factoring code is that you already have something to work with that can execute the task at hand, re-factoring just requires a deep dive into where efficiency is lacking.  A disadvantage of re-factoring the code is that the previously designed code can prevent wider insights from being used in problem-solving that is usually more likely with fresh design.  In the oriignal vs refactored VBA scripts in particular, the speed of the code was greatly increased by refactoring. However, a disadvantage is that the buttons were already created that had to have macros reassigned that was not explicitly requested or part of the code. Had it not been something that I tested before completing the code, I may have missed that and used the original code rather than the refactored code.
