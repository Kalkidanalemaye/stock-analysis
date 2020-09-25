# Project Overview

For this project, I used stock data to analyze and understabd the trends. To achieve the analysis, I used VBA which is often the program used to analyze data in the financial industry.

To start, I looked at how actively DQ was traded in 2018. This will help determine whether a stock that is traded often, will accurately reflect the value of the stock. 

1. First, sum up all of the daily volume for DQ, and it will lead to the yearly volume and a rough idea of how often it gets traded.

2. To calculate the total daily volume in 2018 use loops, which is a process of iterating through a dataset repeatedly.

Next, I looked at how DQ performed in 2018. And to measure this, calculate the yearly return for DQ. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year.

1. To calculate the yearly return, determine DQ’s first closing price and last closing price.

2. To find the first closing price of DQ’s data: (1)loop through all the rows, (2)check if the current row is the first row of DQ’s data. If so, set the starting price to the closing price in the current row.

3. Next, check for two conditions: (1)if the ticker in the current row is DQ and (2)if the ticker in the previous row is not DQ. If so, set the ending price to the current row’s closing price.

4. To check both conditions at once use logical operators.

# Challenge
Created an analysis for the stock market for Steve to compare stocks and inform his parents which one is the best option.
The analysis is refactored using 4 different loops in order to allow for multiple stock analysis. 
