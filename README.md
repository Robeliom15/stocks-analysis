# stocks-analysis

## Overview of Project

### Purpose
The purpose of this analysis is to find out if refactoring the code made it run more efficiently. Steve wants to have a code that can run more efficiently so he can expand the dataset. This new dataset will potentially have thousands of stocks instead of a few dozen which is why this new code is needed. 

## Results

### Analysis of the Refactored Code
These are the execution time for the 2017 and 2018 using the original code:
![ stock_analysis_2017]( https://github.com/Robeliom15/stocksanalysis/blob/main/Resources/stock_analysis_2017.png?raw=true)

![ stock_analysis_2018]( https://github.com/Robeliom15/stocksanalysis/blob/main/Resources/stock_analysis_2018.png?raw=true)

These are the execution time for the new refactored code:
![VBA_challenge_2017](https://github.com/Robeliom15/stocksanalysis/blob/main/Resources/VBA_challenge_2017.png?raw=true)

![VBA_challenge_2018](https://github.com/Robeliom15/stocksanalysis/blob/main/Resources/VBA_challenge_2018.png?raw=true)

We can see that the refactored code has greatly reduced the time needed to execute the code. Additional, the 2018 year executes much faster than the 2017 year despite having the same amount of data. 


The code refactored looks like this:
```
'1a) Create a ticker Index
    Dim tickerindex As Integer
    tickerindex = 0
    '1b) Create three output arrays
    Dim tickervolumes(0 To 11) As Long
    Dim tickerstartingprices(0 To 11) As Single
    Dim tickerendingprices(0 To 11) As Single
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickervolumes(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        '3a) Increase volume for current ticker
        tickervolumes(tickerindex) = tickervolumes(tickerindex) + Cells(i, 8).Value
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
                tickerstartingprices(tickerindex) = Cells(i, 6).Value
            End If
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
                tickerendingprices(tickerindex) = Cells(i, 6).Value
            'End If
            '3d Increase the tickerIndex.
            'If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
                tickerindex = tickerindex + 1
            End If
    Next i
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1) = tickers(i)
        Cells(4 + i, 2) = tickervolumes(i)
        Cells(4 + i, 3) = tickerendingprices(i) / tickerstartingprices(i) - 1
    Next i
```

## Summary

- What are the advantages or disadvantages of refactoring code?

- How do these pros and cons apply to refactoring the original VBA script?

