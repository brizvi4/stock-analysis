# Analysis of Green Stocks

## Overview of Project

### Purpose 

Steve needs to helps his parents in making an informed decision regarding which green stocks to invest in. I am given a data of 12 different green stocks for the years of 2017 and 2018. I am interested to see the growths of the stocks in these two years. In order to save Steve the hassle of making manual calculations from the data, I am going to make use of macros in VBA. I start this project by analyzing DQ stocks. After completing the code for DQ stock, I make another macro which loops through the entire table and captures other stocks as well. This macro is stored in AllStocksAnalysis() Subroutine. In order to check how much time it takes to run the code, I inserted a timer in the code to accomplish our mission. I later found out out that there is a better way to code this sub routine and the time taken to run can also be reduced. 

I decide to refactor the same code presented in AllStocksAnalysis() Subroutine. For that, I created another subroutine called AllStocksAnalysisRefactored(). In this macro, I loop through all the data one time (as comapred to 12 times like last time) in order to collect the same information. I insert a timer in this new macro as well and notice the time it takes to run this code. 

There are two main purposes of this project:

- Figure out which green stock should Steve's parents invest in
- Compare times for both the macros and notice their difference


## Results

### Stocks Performance

By looking at the table in "All Stocks Analysis" work sheet, It can be seen clearly that with the exception of TERP stock, all other stocks performed very well in 2017. All of them had a positive growth by the end of 2017.

However, if I look at the 2018 table, it can be noticed that only ENPH and RUN stocks had a positive growth. I would highly suggest Steve to advice his parents to invest only in ENPH and run socks. Many of the stocks which were doing well in 2017 did not do too good in the following year. 

### Execution times

There is major difference in execution times for AllStocksAnalysis() and AllStocksAnalysisRefactored() subroutines. For the former sub routine, I get a value of around 2 seconds (for both years) as seen in the images below:

![image](https://user-images.githubusercontent.com/95254809/149613055-d859f5b5-fc85-404d-85e4-9e9d214829d6.png)

![image](https://user-images.githubusercontent.com/95254809/149613094-5b3752ff-3b19-41a9-aaad-582d4b2769d2.png)

However, for the latter subroutine, we get a value of around 0.4 seconds for both years. Please see the images below:

<img width="561" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/95254809/149613139-d219e79a-2ef7-4b5a-8947-ca076d55b0b9.PNG">

<img width="566" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/95254809/149613140-83532585-5f12-42d7-b463-31a46cb95a91.PNG">

It can be safely concluded that the execution time of AllStocksAnalysisRefactored() subroutine is around 5 times more faster than AllStocksAnalysis() subroutine.

## Summary

### Advantages or disadvantages of refactoring code

Refcatoring the code can greatly reduce the execution time. Secondly, it also makes the code cleaner and easier to read. Clean code is easy to update and improve. It is also very easy to scale if the need for alteration arises. For instance, if we needed to add more stocks to our data, that can easily be done. 

A major disadvantage I noticed was that if a code is very big, it can be quite difficult and confusing to refactor it. If there is an important application which has a huge amnount of code, this can be quite risky.

### Advantages and Disadvantages of the original and refactored VBA script

A slight disadvantage I noticed with the refactored VBA script was that it got really confusing to deal with tickers and tickerIndex variables. Other than that, I would say that if I had to come up with the code from start, I would have found it easier to write the code in AllStocksAnalysis() sub. This is mainly because that way, I would deal with only one ticker variable. 

Major advantage of the refactored code is that it is very fast. It is also easy to add more stocks in this code since the code only loops once through the entire table. With AllStocksAnalysis() subroutine, a great number of stocks would have taken a lot of time. 

In the end, I would conclude by saying that after weighing the pros and cons, I would definitely go with the refactored code. 
