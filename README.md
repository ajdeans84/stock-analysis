# Stock Analysis

## Overview of Project
   In this project, we looked at the impact of refactoring code has on runtime. By measuring performance (via time it takes to run code), we can see 

## Results
  //use images and examples of code. compare stock performance between the two years, as well as execution of original and refactored script 
  
  =formatting code left in original code so that time difference was not related to running additional formatting options (bolding and underlining, as well as conditional formatting for positive and negative returns)
  
  ![2017](VBA_Challenge_2017.png)
 
  ![2018](VBA_Challenge_2018.png)

## Summary
  //advantages or disadvantages of refactoring code
  In general, refactoring code is a way to improve it, causing the code to be more productive. It could increase the amount of time it takes the code to run, and could increase code readability. However, refactoring code does take more time and effort, as one must spend time diligently seeking out the most efficient solutions and rewriting code that otherwise solves the problem. If the code will not be run repeatedly, the extra invested time it takes to refactor the code may be more than the overall time saved by the code running faster after being refactored. That is, it could take a team an hour to refactor a block of code, but if the new code only saved 20 seconds on compile and is only ever ran 100 times, the team lost 26 minutes overall making these changes. However, for code that may be run many times by many users, the overall time saved can be a tremendous victory. 
  Refactored code also easier to continue to work with, so for larger projects, spending the extra time to clean up code in the early stages means that it is easier to maintain and update the code in the future. 
  
  //how to those pros and cons apply to refactoring the original VBA script?

  In our initial project for Steve, he wanted to see the return on a small number of stocks from 2017 and 2018. This code could be run multiple times, but as the data was historical (and not likely to change by being ran again), there was no reason for Steve to care about saving microseconds by running this data. However, as he asked for an updated program to search the entire stock market over multiple years, we can see why he would then care that time is saved, as he may run this many more times and it would likely take longer to loop through thousands of tickers instead of dozens. 
  
  [time??]
