# Stock Analysis

## Overview of Project

### Purpose
This project was started to help Steve analysis some 2017 and 2018 stock data he had given to me in an excel workbook. Now the purpose of this project is to refactor VBA code that I had previously written to analyze theStock Market data steve gave me to make the code run faster and possibly handle large data sets.

## Results

### Analysis of results
After refactoring the code successfully it resulted in a few things. The first thing that it resulted in was much cleaner code. As you can see in the sample below the code is very easy to follow and explains what is happening in each step.

```
  '1a) Create a ticker Index

    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickervolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickervolumes(i) = 0
        
    Next i
```

The second that this refactoring resulted in was a much faster run time for the Macros. In my original code it took approximately .734 seconds to run the Macro for the 2017 data and .727 seconds to run the Macro for the 2018 data. After the refracting it now takes approximately .117 seconds to run the Macro for the 2017 data and .109 seconds to run the Macro for the 2018 data. Below you can see the run times given by the VBA code after running the refactored Macro for each data set.

![VBA_Challenge_2017.png](https://github.com/SZapata4/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2017.png]( https://github.com/SZapata4/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary
After concluding this project, I can see that there are few cons and pros that come with refactoring code. The cons are that it is a very rigorous and time-consuming task. If you are in a time crunch and already have code that works refactoring may cause you to run in to new problems that you did not have before and could cause to miss an important deadline or turn in date. The Pros are that if you the time you can really take you coding to the next level by getting it to run at faster speeds, handle more date, and be easier to follow along with.
For this project I had enough time to refactor with being endanger of missing any deadlines so there were a lot of pros to doing this refactor. The Macro ran much faster and is now easier to read. It should also be able to handle for data in theory. The only con on this project for me was that it was very time consuming to refactor the code that I already had working. With that being said I believe in this case refactoring the code provide to be the right choice.

