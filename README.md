# Stocks Analysis with VBA

## Overview of Project

### Purpose
The objective of this challenge was to provide students with the programming skills in VBA to read data on a sheet 
peform calculations, utilize arryas, and output results to a new worksheet.   The exericse also taught students how
to make their code run more efficiently.

### Background

The VBA of Wallstreet project applied those programming skills to help Steve see how various stocks performed 
based on stock data from in 2017 and 2018.  There were 12 stocks with peformance history that were on the 2017 and 2018 
worksheets.  I had some trouble with coding this at first, so I created a "SmallSet" worksheet with some of the data 
so that I could code and debug easier to make sure my calculations and formatting worked.  Formatting cells with colors
on returns made it easy for Steve to see which stocks had a positive return (green color) versus a negative 
return (red color).

Prompting for a year using an input box also allowed Steve to pick which year to analyze.

Results below are based on the Refactoring code which is in Macro AllStocksAnalysisRefactored() 

## Results

### Stock Analysis for Year 2017


![Stock Results 2017](https://github.com/gaudiom4git/stocks-analysis/blob/main/resources/Year2017Results.png)

Stock results for the Year 2017 were mostly positive.  All stocks had a positive return except for ticker TERP which 
had a negative 7.2%.  Best performer was ticker DQ with almost a 200% return.  Most traded stock was SPWR with volume
of 782,187,000.   Lowest traded stock was DQ with a volume of 35,796,200.

### Stock Analysis for Year 2018

![Stock Results 2018](https://github.com/gaudiom4git/stocks-analysis/blob/main/resources/Year2018Results.png)

Stock results for the Year 2018 were mostly negative.  The only 2 that had a postive return were tickers ENPH with a 
return of 81.9% and RUN with an even higher 84.0%.  Both tickers had very high volumes. Worst performer was DQ with a 
negative 62% return.  Volume was much higher for DQ compared to 2017.  

### Runtimes with and without Refactoring

Macro	|Year	| Runtime			| Image
-------|------|---------------|---
Original Code| 2017 |0.6875s |[VBA_Challenge_2017.png](https://github.com/gaudiom4git/stocks-analysis/tree/main/resources/VBA_Challenge_2017.png).
Refactored Code | 2017 | 0.15625s |[Refactored_2017.png](https://github.com/gaudiom4git/stocks-analysis/tree/main/resources/Refactored_2017.png).
Original Code | 2018 | 0.7109375s |[VBA_Challenge_2018.png](https://github.com/gaudiom4git/stocks-analysis/tree/main/resources/VBA_Challenge_2018.png).
Refactored Code| 2018 | 0.171875s | [Refactored_2018.png](https://github.com/gaudiom4git/stocks-analysis/tree/main/resources/Refactored_2018.png).

Refactored code ran much faster for both years.

## Summary

Advantages to refactoring code is faster runtimes, usage of less memory, simpler for other coders to follow.  It
might be a little more difficult to debug refactored code as you are storing more data in memory and will have debug to step
through the program.

The original script had us reading through the whole stock data worksheet for each ticker.  So, for 12 stock tickers, you
read all the rows in the stock data 12 times, that's 12 x 3103 rows.  That's 37236 rows to process.

The refactored script had us loop through the stock data 1 time.  Within the loop, we were checking for a change in the ticker on the
stock sheet.  If there was a change, volume totals, start and end price were stored in their respective arrays.  The ticker index was 
increased on a change so that the program would move the next ticker (stock symbol) in the ticker array.   

The refactored script works for this worksheet because it is sorted by ticker and closing date.  If it was not sorted this way, the results
would be inaccurate.



