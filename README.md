# Pewlett-Hackard-Analysis

# Stock Analysis

## Overview of Project

Using Excel built-in scripting language VBA to help a client predict wisest investment from a pool of select tickers.

### Purpose
Automate the process of determining the most favorable investment stock by using Visual Basic for Application programming.


## Analysis and Challenges
* Our client Steve provided a spreadsheet of various stocks including 8 colums that depicted Ticker, Date, Open, High, Low, Close, Adj Close prices along with Volume of shares traded which essentially is the number of contracts that have changed hands per day.

* The below images show the return analysis runtime for each respective year 2017 and 2018.

![2017](https://github.com/basecipher/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![2018](https://github.com/basecipher/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

### Analysis of Outcome
* Writing code with static variables resulted in average runtimes of over 1.5 seconds.
* Running code using refactored index to access data in an array took 1.13 seconds to execute "2017" and 1.05 seconds for "2018"
* Creating arrays allowed code to be cleaner, more organize, less resource intensive and quicker run times than conventional use.

### Challenges and Difficulties Encountered
* It was difficult to conceptualize and implement nested for loops, indexes and arrays at first but with practice was able to see how useful they are as stated above along easier to troubleshoot/debug.

## Results
* Steve's best stock options choices for 2018 are ENPH (81.9%) and RUN (84%) as they give off the best rate of return for the most recent year.
* From a coding standpoint refactoring the code with arrays was beneficial in three ways:  faster runtime, less resource intensive and easier to troubleshoot common errors such as "Compile Error: Expected Variable" which denotes I have used incorrectly somewhere to use a keyword for a variable name.
* I really enjoyed the process of developing meaningful code.  I particularly enjoyed creating form control buttons and see myself using them constantly in the future with my personal spreadsheets.
