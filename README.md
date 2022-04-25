# Project Background
This project is to help Steve analyze a dataset that includes the entire stock market over the last few years. Because there are thousands of stocks, so the code I used for analyzing a dozen stocks from 2017, 2018 might not be fast enough. In order to make the code executing time shorter, I refactored code from Module 2 to run the VBA script run faster# stock--analysis
# Purpose of Project
The purpose of this analysis was to create macros that would be easily accessible for Steve, or any other end-user to utilize. Specifically, we were looking at 12 different stocks that Steve's parents were thinking about investing in. We wanted to see the total volume of each stock (meaning how often each stock was traded) from the beginning to the end of the year for the years we have data on, which are 2017 and 2018. Additonally, we wanted to see the percentage return for each of the 12 stocks over a year long period. By doing this analysis, Steve will be somewhat better able to recommend which (if any) of those stocks his parents should invest in. To do this analysis, two macros were used- the original script from the module, and a refactored version of that script. While the output of the macros was the same, the efficiency of the two versions differed- something that will be discussed below.
# Results
## Results of the analysis on stock volume and performance for 2017 and 2018.
2017 was a much better year in terms of valuation for every single one of the 12 selected stocks except for the company with "RUN" as their ticker. "DQ", the oringinal stock that Steve's parents were interested in for very rational reasons, highlights the difference in performance from '17 to '18. In 2017, "DQ" price increased 199.4% but in 2018, decreaesd 62.6%. Overall, this seems like a very volatile sector, and for Steve's parents, who can be assumed to be close to retirement, I would suggest that they look at either index funds, vanguard 20XX funds, or frankly, just fixed income options. In order to help visualize the change over time for the selected stocks, I've inserted the tables of all 12 stock's performances for both 2017 and 2018.

![alt text](https://github.com/tarini-mi7/stock--analysis/blob/main/resources/VBA%20Challenge%202017%20Image.png)

![alt text](https://github.com/tarini-mi7/stock--analysis/blob/main/resources/VBA%20Challenge%202018%20Image.png)


As you can see from the tables, there doesn't seem to be a relationship between the stock volume (the amount of which that specific stock was traded during a given year) and it's valuation. For example, "HASI" which in 2017 increased in value by 25.8% was traded a total of 80,949,300 times. In 2018, "HASI"'s value decreased by 20.7% while being traded 104,340,600 times (nearly 24,000,000 more trades than in 2017). Perhaps, if you just looked at that single stock, You could squint and say that there might be a very weak, negative relationship between a stocks value and the amount of times it is traded. However, "CSIQ" tells a different story. In 2017, "CSIQ" increased by 33.1% while being traded 310,592,800 times. Then, in 2018, "CSIQ" stock decreased in value by 16.3% while being traded 200,879,900 times (which is about 110,000,000 less trades than in 2017). Here for "CSIQ" (and unlike "HASI") the stock was traded about 33% less in 2018 than in 2017, but still posted a loss of value like "HASI" did. Overall, there just doesn't seem to be any correlation between a stocks volume and it's performance (which is something that Steve's parents thought back in the module).

In terms of execution time for the analysis, the refactored code was faster for both the 2017 and 2018 data sets than the non-refactored code. The refactored code completed the analysis in .0703 seconds for 2017 as well as for 2018. The non-refactored code completed the analysis in .3593 seconds for the data from 2017, and in .3475 seconds for the analysis of the data from 2018. In order to illustrate the differences, I've uploaded a bar chart of the run times below.



# Advantages and disadvantages of refactoring code.
An advantage of refactoring code is that doing so makes it cleaner and easier to read. It's like editing a draft in order to make it more presentable. You're not changing the behavior, or the outcome of the code, you're just making it run more efficiently. A possible disadvantage is that you introduce bugs (which I certainly did) into the code by refactoring it. Those bugs will have to be addressed before you can run the code against the original to see if it indeed performs more efficiently.

# How those advantages and disadvantages applied.
The refactored code does look cleaner and has a better format that makes it easier to read line by line. If look at the original script's conditionals used to run the analysis each stocks total volume, starting, and ending price, you can see it is less efficient than the refactored code's condtionals, which are posted below the original. One concept we learned in module number 2 was "DRY" which is an acronym for don't repeate yourself. By refactoring the code, some of the unneeded repetition from the original script could be removed.

                 **Original Script**                                          
For i = 0 To 11 ticker = tickers(i)
totalvolume = 0 Worksheets(yearValue).Activate
For J = 2 To RowCount
If Cells(J,1).Value= ticker Then
TotalVolume = Totalvolume + Cells(J,8).Value
End If If Cells(J-1,1).Value <> Ticker AND Cells(J,1).Value = ticker Then
startingPrice = Cells(J,6).Value
End if
If Cells(J+1,1).Value <> ticker and Cells(J,1).Value = ticker Then
EndingPrice = Cells(J, 6).Value
End If

                **Refactored Script**
For i = 0 To 11 tickerVolumes(i) = 0 Next i For i = 2 to RowCount

  Next i
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
End If
 
If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    tickerIndex = tickerIndex + 1
End If
Also, the refactored script does run faster than the original for both the analysis of 2017 and 2018. I'm sure this increase in processing speed is derived from having the script only contain exactly what it needs in order to run, without any unnecessary lines. However, in the process of refactoring the code, I must've run into dozens of bugs that I created by incorrectly rewriting it. Given the time I had to put into writting it correctly, compared to the small difference in efficiency gained, (the refactored code was about 28 hundreths of a second faster) I don't know that it was worth it in a vacuum to refactor in this scenario. Of course, it was well worth it as a learning experience. And whenever I write much more complex scripts in the future, I'm sure that refactoring the code will be incredibly useful and in some cases necessary.


