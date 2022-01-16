# stock-analysis With VBA

Performing Stock analysis on [VBA_Challenge.xlsm](https://github.com/crdhilep/stock-analysis/blob/main/VBA_Challenge.xlsm)

### Overview of Project

## Purpose
The purpose of this project was to refactor VBA code to increase the efficency of the script to collect certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing.

## The Data

The data that is presented includes two sheets with stock information on 12 different stocks. The sheets contains  information contains a ticker value, stock issued data, opening, closing, adjusted closing, the highest and lowest price and the volume. The goal is to retrieve the ticker, the total daily volume and return on each stock.

### Results

 ## Analysis

Before refactoring, I began by copying the code that needed to create the input box, headers, ticker array and to activate the appropriate worksheet. The steps are listed in order to set the structure for the refactoring.


### Key Refactoring points:

In the original code shared has two inner iterations.
As part refactoring, I elimianated the need and made it simple.
That saved time and resources of iterating over larger records.
      
 
### Summary

Refactoring helps make our code simpler, more organized and reduce duplication.


### Result of Refactoring Stock Analysis:

Before Refactoring codes took 262 MilliSeconds to execute and After Refactoring we are able to achieve execution time below 70 MilliSeconds(~190 ms Performance Gain)

  ##2018 Before Refactoring:
 
 ![2018 Before Refactoring](/resources/VBA_Challenge_Before_Refactoring.png)
 


  ##Results After Refactoring:

  ![2017 Refactored Results](/resources/VBA_Challenge_2017.png)
 
 ![2018 Refactored Results](/resources/VBA_Challenge_2018.png)
