# Project Overview

For this project, I used stock data to analyze and understand the trends. To achieve the analysis, I used VBA which is often the program used to analyze data in the financial industry.

To start, I looked at how actively DQ was traded in 2018. This will help determine whether a stock that is traded often, will accurately reflect the value of the stock. 

1. First, sum up all of the daily volume for DQ, and it will lead to the yearly volume and a rough idea of how often it gets traded.

2. To calculate the total daily volume in 2018 use loops, which is a process of iterating through a dataset repeatedly.

Next, I looked at how DQ performed in 2018. And to measure this, calculate the yearly return for DQ. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year.

1. To calculate the yearly return, determine DQ’s first closing price and last closing price.

2. To find the first closing price of DQ’s data: (1)loop through all the rows, (2)check if the current row is the first row of DQ’s data. If so, set the starting price to the closing price in the current row.

3. Next, check for two conditions: (1)if the ticker in the current row is DQ and (2)if the ticker in the previous row is not DQ. If so, set the ending price to the current row’s closing price.

4. To check both conditions at once use logical operators.

#### Finding

Daqo dropped over 63% in 2018, this analysis shows that DQ is not recommended stock to invest in.

# Challenge

### Objective
To create an analysis of stocks in 2017 and 2018 in order to determine the annual total volume, and return rate of each stock.

### Steps taken
First, create an array for all the stock names as ticker values and create a loop (iterated code) that would count how many rows were within each data set. 

Next, create a nested loop where the first one looped through the ticker values and the nested loop would loop through the data set rows to collect the starting and ending price.

Then, it will count and add all the volume for each individual tickers.

### Refactor

In order, to implement an efficient method of coding I refactored the initial code. To clean the code, start with removing the nested loop and creating 4 individual loops. Declare the variables and set tickerIndex to 0 before the loops.

Since we have the tickers array created, the first loop initialized an array for the ticker volume, starting and ending price, and set their values to 0.
The second looped through rows 2 to 3013. Within this loop, we had an if statement to add 1 to tickerIndex if the next row did not match the current row. The next if statements collected the starting price, ending price and ticker volume for the current ticker value, which was indicated by our tickerIndex.
Our third loop output all the data stored in each variable into a table in our new sheet labeled "All Stocks Analysis". We also formatted the cells within this loop, such as changing the table header to bold.
Our last loop added conditional formatting to our return rate column. It states that if the return rate is positive, the cell will be highlighted green and red if it is negative.
Conclusion
Breaking the nested loop apart allows us to loop through 3013 rows only once to gather our data. As a result, it cuts down the amount of time the code needs to run. Not only that, in addition to adding comments, the code is easier to read, which means it is easier to debug.

After refactoring our code, we can turn our attention to the data. In 2017, our 12 stocks were all showing green for their return rate except for one. However, when you analyze the same stocks in 2018, most of their return rates were in the negatives, with an exception of two that were showing positive. This only shows a glimpse of how these stocks were doing in 2017 and 2018, but more information and further analysis is needed to determine which stocks are worth investing in and which are not.
