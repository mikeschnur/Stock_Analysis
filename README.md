# Stock_Analysis

## Purpose

Steve would like to recieve an analysis of stocks in the years 2017 and 2018. After using Excel and VBA to utilize a timer, arrays, if statements and loops the analysis was performed. Although, the initial code was not ompitally formatted to receive the results as soon as possible. Thus, the code was refractored to reduce time in which the results are received.  

## 2017 vs 2018 

![2017_StockResults](https://user-images.githubusercontent.com/108902185/194435234-23c35a79-170c-40d1-8035-2188e108fd3e.png)

![2018_StockResults](https://user-images.githubusercontent.com/108902185/194435244-907364d5-9a6a-429f-9ea0-f144e0a3bd5f.png)

When comparing the the performance of 11 different stocks in the years 2017 and 2018, there are glaring differences between results by year as well as by stock. Firstly, the return for every stock besides RUN was lower in 2018 than 2017. Thus, the performance of a given stock in 2018 does not nessecarily doom it as an investment. ENPH and RUN were the only two stocks to have a positive return in both years, while TERP is the only stock to have a negative return both years. 

## Original vs. Refactored Run Times 
![ModuleTime_2017](https://user-images.githubusercontent.com/108902185/194435367-e638adb1-08a0-4cb3-a4aa-4c2a252dedc0.png)

![ModuleTime_2018](https://user-images.githubusercontent.com/108902185/194435371-c8cf8e93-0e0f-4620-9530-b1c46f97e263.png)

Both run times for the original code were around 0.6 seconds.
![VBAChallenge_2017](https://user-images.githubusercontent.com/108902185/194435450-21572791-f53d-4cb3-b6cb-b1130a0867bd.png)

![VBAChallenge_2018](https://user-images.githubusercontent.com/108902185/194435465-49b102c4-b3c5-402e-8a34-5048b6ab4cfe.png)

Both run times for the reactored code were around 0.1 seconds.

## How do the codes differ?
![Screen Shot 2022-10-06 at 7 41 04 PM](https://user-images.githubusercontent.com/108902185/194437490-b9613a96-07e3-4ecb-8484-10f63a94f946.png)

The original code featured a nested loop that had the output within the loop, which caused a delay in output time.

