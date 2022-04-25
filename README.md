# Project Background
This project is to help Steve analyze a dataset that includes the entire stock market data over the last few years. Because there are thousands of stocks, so the code I used for analyzing a dozen stocks from 2017, 2018 might not be fast enough. In order to make the code execution time shorter, I refactored the code from Module 2 to run the VBA script faster.
# Purpose of Project
The purpose of this analysis was to create macros that would be easily accessible for Steve, or any other end-user. Specifically, we were looking at 12 different stocks that Steve's parents were thinking about investing in. We wanted to see the total volume of each stock (meaning how often each stock was traded) from the beginning to the end of a year for the years we have data on, which are 2017 and 2018. Additonally, we wanted to see the percentage return for each of the 12 stocks over a year long period. By doing this analysis, Steve will be somewhat better equipped to recommend which (if any) of those stocks his parents should invest in. To do this analysis, two macros were used - the original script from the module, and a refactored version of that script. While the output of the macros were the same, the efficiency of the two versions differed, which will be discussed below.
# Results
## Results of the analysis on stock volume and performance for 2017 and 2018.
2017 was a much better year in terms of valuation for every single one of the 12 selected stocks except for the company with "RUN" as their ticker. "DQ", the oringinal stock that Steve's parents were interested in for very rational reasons, highlights the difference in performance from '17 to '18. In 2017, "DQ" price increased 199.4% but in 2018, decreased 62.6%. Overall, this seems like a very volatile sector, and for Steve's parents, who can be assumed to be close to retirement, I would suggest that they look at either index funds, or just fixed income options. In order to help visualize the change over time for the selected stocks, I've inserted the tables of all 12 stock's performances for both 2017 and 2018.

![alt text](https://github.com/tarini-mi7/stock--analysis/blob/main/resources/VBA%20Challenge%202017%20Image.png)

![alt text](https://github.com/tarini-mi7/stock--analysis/blob/main/resources/VBA%20Challenge%202018%20Image.png)


As you can see from the tables, there doesn't seem to be a relationship between the stock volume (the amount of which that specific stock was traded during a given year) and it's valuation. For example, "HASI" which in 2017 increased in value by 25.8% was traded a total of 80,949,300 times. In 2018, "HASI"'s value decreased by 20.7% while being traded 104,340,600 times (nearly 24,000,000 more trades than in 2017). Perhaps, if you just looked at that single stock, You could squint and say that there might be a very weak, negative relationship between a stock's value and the amount of times it is traded. However, "CSIQ" tells a different story. In 2017, "CSIQ" increased by 33.1% while being traded 310,592,800 times. Then, in 2018, "CSIQ" stock decreased in value by 16.3% while being traded 200,879,900 times (which is about 110,000,000 less trades than in 2017). Here for "CSIQ" (and unlike "HASI") the stock was traded about 33% less in 2018 than in 2017, but still posted a loss of value like "HASI" did. Overall, there just doesn't seem to be any correlation between a stocks volume and it's performance (which is something that Steve's parents thought back in the module).

In terms of execution time for the analysis, the refactored code was faster for both the 2017 and 2018 data sets than the non-refactored code. In order to illustrate the differences, I've uploaded a bar chart of the run times below.
![alt text](https://github.com/tarini-mi7/stock--analysis/blob/main/resources/code%20run%20time%20graph.png)


# Advantages and disadvantages of refactoring code.
An advantage of refactoring code is that doing so makes it cleaner and easier to read. It's like editing a draft in order to make it more presentable. You're not changing the behavior, or the outcome of the code, you're just making it run more efficiently. A possible disadvantage is that you may introduce bugs.

# How those advantages and disadvantages applied.
The refactored code does look cleaner and has a better format that makes it easier to read line by line. If look at the original script's conditionals used to run the analysis each stocks total volume, starting, and ending price, you can see it is less efficient than the refactored code's condtionals, which are posted below the original. One concept we learned in module number 2 was "DRY" which is an acronym for don't repeate yourself. By refactoring the code, some of the unneeded repetition from the original script could be removed.

**Original Script**
```                 
    For i = 0 To 11
        Dim ticker As String
        ticker = tickers(i)
        totalVolume = 0
        'Activate the data worksheet
        Worksheets(yearValue).Activate
        '5) Loop through rows in the data.
        For j = rowStart To RowCount
            '5a) Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                'increase totalVolume
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            '5b) Find the starting price for the current ticker.
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            '5c) Find the ending price for the current ticker.
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j
        '6) Output the data for the current ticker.
        'Activate the output worksheet
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
```
**Refactored Script**
```
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
    Next tickerIndex
    
    tickerIndex = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If

    Next i
```

Also, the refactored script does run faster than the original for both the analysis of 2017 and 2018. I'm sure this increase in processing speed is derived from having the script only contain exactly what it needs in order to run, without any unnecessary lines. The refactored code iterated over the number of rows exactly once as opposed to for each ticker symbol, which leads to a reduction in iterations by multiple factors. However, in the process of refactoring the code, I must've run into dozens of bugs that I created by incorrectly rewriting it, and fixed them to achieve the desired output. I am sure that this experience of refactoring code will definitely be useful in the future while dealing with complex code.


