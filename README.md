# Stock-analysis
## Overview of the Project
Perform analysis on Stocks to help Steve
### Purpose
Refractoring the VBA code to analyze Stocks Data for Steve to make it more efficient 
#### Background
Steve has an excel file with stock data to analyze and he needs macros to run the analysis on new data in the future. By doing this analysis he wants to convince his parents to invest green enegergy stocks which has great revenue in the future. 

## Result 
### Refractoring the provided Code to run analysis for each year for all tickers, 2017 and 2018, respectivelt. In this proccess I used code which loop through all the data one time which more efficient and takes significantly less time to execute. 
1. The main difference between the code from the macros we wrote during the module and the code for this challenge is we need to use tickerIndex to access the correct index across the arrays we created for this challenge, they are tickers array and three output arrays as below:
##### Dim Tickers(12) as Single
##### Dim tickersVolume(12) as Long
##### Dim tickerStartingPrice as Single
##### Dim tickerEndingPrice as Single 
2. Create for loops to loop through the 4 arrays to output the "Ticker", "Total Daily Volume", and "Return" columns in the spread sheet as follwoing for 2017 and 2018, respectively.

![2017_Stocks](https://user-images.githubusercontent.com/65901034/173418495-57f74eb7-2b6a-47a9-bf7c-e5c212991787.png)

![2018_Stocks](https://user-images.githubusercontent.com/65901034/173418504-eb2adb21-2ba9-4e5b-99ca-ac491ce4dbad.png)

3. I measured the perforamnce of the macro after refractoring which is much faster than previous code. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/65901034/173419932-9de109c6-d970-4585-9bca-14d19d421462.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/65901034/173419922-233ffd42-b6bd-401b-bc5a-5bc6f54dabda.png)




## Summary
