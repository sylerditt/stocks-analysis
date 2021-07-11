# VBA of Wall Street

## Overview of Project

### Purpose and Background
The purpose of this analysis is to find the total daily volume and yearly return for each stock in the data set and be able to compare them to Steve's parents' stock in DAQO. We created a button to analyze the data so Steve could just click it in order to get the information he needed. Then, we wanted to refactor our code to see if we could get our program to run faster. 

## Results
We were able to find the difference between the stock performance in 2017 and 2018 using both the original program and the refactored program. The only ticker with a negative return in 2017 was TERP while the only ticker with positive returns in 2018 were Ewere ENPH and RUN.

However, when running our refactored program, we were able to capture the results in 0.094 seconds for 2017 and 2018 (See Figures 3 and 4) compared to 0.867 seconds when running the original program (See Figures 1 and 2). By looping through all the data in our stocks worksheet and using fewer steps, we improved the code and made the program run more quickly. 

## Summary

### Refactoring Code in General
Refactoring code in general demonstrates an understanding of VBA script and how to make your code more efficient. When first learning how to code using VBA, we learn all the steps taken to make a program run. However, not all these steps are necessary to perform the actions you are attempting to perform in order to analyze data. By looking at these steps and finding out which ones are not necessary. Looping once through the entire data set, instead of looping through small sections of data, allows code to run more quickly as it is performing less tasks. 
### Original vs. Refactored VBA Script
The advantages of the origianl VBA script involved teaching us how to code and the methods behind VBA in general. We were able to look at what a more efficient code is going through behind the scenes, so to speak, and learn useful methods for looping and analyzing specific parts of a data set. However, the advantages of the refactored VBA script were only using one loop to analyze the entire data set and running the program quicker. The disadvantages of refactoring the VBA script were using our prior knowledge to adjust the code in an efficient manner. This was difficult as the syntax and arrays were slightly different and keeping all elements of the code the same, referring to the appropriate arrays, and making sure that one loop runs appropriately proved challenging. 

After figuring out how to remain consistent throughout the refactored code, I learned useful tools to make my code more efficient and do less work that is unnecessary, while understanding the methods behind the more efficient code. 

## Figures

### Figure 1
<img src="2017.Original.png">

### Figure 2
<img src="2018.Original.png">

### Figure 3
<img src="2017.Refactored.png">

### Figure 4
<img src="2018.Refactored.png">

