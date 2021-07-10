
# Green Stocks Analysis
## Overview of Project
Develop a workbook to present the returns and daily volume per year of stocks

### Purpose
To present a solution to expand the dataset to include the entire stock market over the last few years delivering a high-performance code to allow it to process a high volume of stock information. 


## Analysis and Challenges

### Analysis of stock performance between 2017 and 2018
By observing the green stock returns presented in the calculated information shown in the stocks analysis table of 2018 compared to 2017 it is understandable that while most of the green company stocks regressed their returns in 2018, two stocks (ENPH and RUN) kept their returns raising potentially with around 80% of increase in 2018

On the other hand, TERP stock performed poorly in both years shrinking 7% and 5% consecutively and SPWR had the sharpest fall of 44% in 2018 against a raise of only 23% back in 2017. 

![all_stocks_analysis_2017](https://raw.githubusercontent.com/ramonmsa/stock-analysis/main/Resources/support_readme/all_stocks_analysis_2017.PNG "All Stocks Analysis 2017") | ![all_stocks_analysis_2018](https://raw.githubusercontent.com/ramonmsa/stock-analysis/main/Resources/support_readme/all_stocks_analysis_2018.PNG "All Stocks Analysis 2018") 



### Analysis of the execution times of the original script and the refactored script
For this analysis it was necessary to refactor the VBA code in order to allow larger data sets to perform a faster calculation. 

The code refactor took place mainly by:
 - storing the calculation results into arrays
```VBA    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
 - And then looping through the arrays to output the _Ticker_, _Total Daily Volume_, and _Return_
```VBA
    tickerIndex = 0
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
        tickerIndex = tickerIndex + 1
        
    Next i
```

That way it was possible to optimize the performance of the code as it can be seen in the images below showing the times before and after the code was refactored.


![times_before_refactoring_2017](https://raw.githubusercontent.com/ramonmsa/stock-analysis/main/Resources/support_readme/all_stocks_time_original_script_2017_.PNG "Times before the refactoring for 2017 calcs") | ![times_after_refactoring_2017](https://raw.githubusercontent.com/ramonmsa/stock-analysis/main/Resources/support_readme/all_stocks_time_refactored_script_2017_.PNG "Times after refactoring for 2017 calcs") 

![times_before_refactoring_2018](https://raw.githubusercontent.com/ramonmsa/stock-analysis/main/Resources/support_readme/all_stocks_time_original_script_2018_.PNG "Times before the refactoring for 2018 calcs") | ![times_after_refactoring_2018](https://raw.githubusercontent.com/ramonmsa/stock-analysis/main/Resources/support_readme/all_stocks_time_refactored_script_2018_.PNG "Times after refactoring for 2018 calcs") 

### Challenges and Difficulties Encountered
The challenge in this analysis was to identify what part of the code needed to be refactored in order to increase the performance. It took a long period spent in research to come up with the solution presented before in this document. 
 
 Even though the time difference shown on the examples does not seem expressive, it is understandable that with a bigger data set the times presented will be by far apart from each other. The reason for that is as the data set grows the execution time of the code increases exponentially. 
 
## Summary

- **In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?**
    - Two main disadvantages on refactoring code is the risk of running out of time and money. 
    - On the other hand, the advantages are endless being the most important in my point of view: better performance, clean and understandable code, minimize bug occurrences. 
- **How do these pros and cons apply to refactoring the original VBA script?**
    - Other than the client’s requirement to have a scalable code to lager datasets, having a clean code where the blocks of code are well defined such as a block to read the information and calculation, and another different block of code to write the information in the worksheet. These will work as benefit to help on future maintenance and improvement.
    - However, the time consuming on research to achieve the reduction on execution time was of a considerable cost for this project.
