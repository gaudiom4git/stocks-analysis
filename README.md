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

![Stock Results 2017](https://github.com/gaudiom4git/stocks_analysis/blob/main/resources/Year2017Results.png)

Stock results for the Year 2017 were mostly positive.  All stocks had a positive return except for ticker TERP which 
had a negative 7.2%.  Best performer was ticker DQ with almost a 200% return.  Most traded stock was SPWR with volume
of 782,187,000.   Lowest traded stock was DQ with a volume of 35,796,200.

### Stock Analysis for Year 2018

![Stock Results 2018](https://github.com/gaudiom4git/stocks_analysis/blob/main/resources/Year2018Results.png)

Stock results for the Year 2018 were mostly negative.  The only 2 that had a postive return were tickers ENPH with a 
return of 81.9% and RUN with an even higher 84.0%.  Both tickers had very high volumes. Worst performer was DQ with a 
negative 62% return.  Volume was much higher for DQ compared to 2017.  

### Runtimes with and without Refactoring

Initial code in the AllStocksAnalysis() macro writes results to the worksheet as the script loops through the
Stock data worksheet year.  Refactored code in AllStocksAnalysisRefactored() had us store the results in arrays
which was much faster than writing to the worksheets while looping through all the data.   

The refactored code wrote to the worksheet while looping through the arrays that stored the stock result values.

## Summary

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt). 

There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

The runtime pictured here are for the refactored code for year 2018.

The runtime pictured here is for the original code for year 2018.   

Results in a % difference.

