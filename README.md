# stock-analysis

## Overview of Project

### Purpose and Background
The client, Steve, initially asked for an analysis of a handful of green energy stock tickers for his parents to invest in. With the information provided in the initial analysis, Steve helped his parents diversify their stocks and invest in higher grossing stocks than DQ alone. Now Steve has requested a refactored workbook so that he can expand the dataset to include the entire stock market over the last few years, and have his workbook run efficiently. 

## Results

![2017_Stocks](Resources/allStocks2017replace.png)

In 2017, all stocks other than TERP had an increase in returns, with DQ showing the highest percent return of 199.4%. 

![2018_Stocks](Resources/allStocks2018replace.png)

However, in 2018 all but two stocks showed negative percent returns. Only RUN and ENPH were successful in this year.


## Summary
### Advantages and Disadvantages of Refactoring Code
Refactoring code increases the efficiency, either by taking fewer steps or improving the logic of the code without adding new functionality. This means that it will take less time to run and use up less memory on the device, as shown by the images below. 

### Original Versus Refactored in this Dataset

![time1](Resources/original2017time.png)              ![time2](Resources/refactored2017time.png)

The refactored code (pictured on right) for the year 2017 ran about 4 times faster than the original code (pictured on left).

![time3](Resources/original2018time.png)              ![time4](Resources/refactored2018time.png)

Similar results are shown for 2018. 


For a full look at the code and the Excel workbook, please [click here.](VBA_Challenge.xlsm)
