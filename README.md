# stock-analysis

[VBA_Challenge.xlsm](/Resources/VBA_Challenge.xlsm)

## Overview of Project

This project is an analysis of 12 stocks using macros to pull specifics of their performace year over year for the years 2017 and 2018.


### Purpose

The purpose of this project is to show how one can write VBA scripts in Excel to format cells and pull points of data from a spreadsheet together into macros to get a more clear picture of the performance of these stocks. Those data points are stored in arrays and those arrays are kept in order and called with an index location of each stock in another array of ticker symbols. It also shows that macros can be hard coded or they can be refactored making them more dynamic. 


## Results

We can see that year over year our anaylized stocks rate of return of has dropped significantly.
![VBA_Challenge_2017.png](/Resources/2017_results.png)
![VBA_Challenge_2017.png](/Resources/2018_results.png)

In most cases, we can conclude that stocks with less volume year over year saw a decrease in rate of return or no return at all. Less volume can be attributed to less demand and less demand always attributes to less value.

It is important to note that by placing a line of code **startTime = Timer** at the beginning our code tells our program to keep track of how much time it took our code to complete, ended by another line of code, **endTime = Timer** near the bottom. By using this, we can keep track of how efficient our code is running and discover what can be changed in our code to improve that efficiency.

In this analysis, a few small pieces of code were refactored by replacing hard coded values with more dynamic ones. Before refactoring, the time it took to run our code was 1.449129 seconds for 2017 and 1.472656 for 2018. After refactoring, we can see the time needed to run was shortened by about 1.20 seconds. That is about 1/6 of the original time needed. 

![VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)


## Summary

Macros in VBA scripts can be a powerful tool used in data science to automate the calling of specific types, ranges, or values of data. They enable analysts to  sort, index, assign, omit or add too, as well as analyze and output vast or small amount of data from a dataset. They can also help us make our data best practices even better by telling us how our code is running and how it might run better for its purpose. 

### What are the advantages or disadvantages of refactoring code?

One advantage to refactoring code is that it more easily allows our code to be used in future projects, some which might be larger or more complex, without needing to change several parts of the code, or in the worst cases rewriting the entire thing. There could be a scenario such as issues of privacy or exclusion that one would not want their code to be so dynamic. 

### How do these pros and cons apply to refactoring the original VBA Script?
By refactoring, our analysis saw incredible gains in efficiency by reducing the time needed to run our code. It also allowed our code to pull data from two different years in two seperate worksheets, which was indicated by a user's input. Prior to that, our code ran for one single year on one single worksheet which was hard coded, limiting it's usefullness.
