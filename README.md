# Stock Analysis

## Overview of Project
   In this project, we looked at the impact of refactoring code has on runtime. By measuring performance (via time it takes to run code), we can see 

## Results
  //use images and examples of code. compare stock performance between the two years, as well as execution of original and refactored script 
  
  By utilizing arrays, we can eliminate the need for nested For Loops and reduce running through all rows of the worksheet once for each ticker we want to check, meaning the original code ran through over 3,000 rows 12 times in the original code, but only had to run through the rows once in the refactored code, eliminating an unnecessary 33,000 instances of checking a row for a matching ticker and possibly adding the volume. 

  2017 - 0.797 seconds before refactoring, 0.109 seconds after
  2018 - 0.789 seconds before refactoring, 0.117 seconds after

  =formatting code left in original code so that time difference was not related to running additional formatting options (bolding and underlining, as well as conditional formatting for positive and negative returns)
  
  ![2017](VBA_Challenge_2017.png)
    ![2017_Orig](VBA_Challenge_2017_Original.png)
 
  ![2018](VBA_Challenge_2018.png)
    ![2018_Orig](VBA_Challenge_2018_Original.png)
    
    
|                    | 2017          | 2018          |
| -------------      |    ---------- | ------------- |
| Before Refactoring | 0.797         | 0.789         |
| After Refactoring  | 0.109         |  0.117        |


 Original code required nested For Loops. While looping through all 12 tickers, we looped through each row, meaning that the code in the "For j = 2 to RowCount" code block was run over 36,000 times. 

 By implementing arrays, we were able to reduce the times a code block was run to just over 3,000 times (one for each row of data). Instead of looping through the tickers, we can check the ticker value against the value we would expect from our array, and update the variables of interest in their own arrays, which are output at the end instead of after each current ticker has been checked. 


  [original code required nested for loops]
  ```
  For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    For j = 2 To RowCount

  ```

  [initializing our arrays]
  ```
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
  ```
  [looping through arrays to output the values of interest]
  ```
    For i = 0 To 11
      Cells(4 + i, 1).Value = tickers(i)
      Cells(4 + i, 2).Value = tickerVolumes(i)
      Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i 
  ```

## Summary
  //advantages or disadvantages of refactoring code
  In general, refactoring code is a way to improve it, causing the code to be more productive. It could increase the amount of time it takes the code to run, and could increase code readability. However, refactoring code does take more time and effort, as one must spend time diligently seeking out the most efficient solutions and rewriting code that otherwise solves the problem. If the code will not be run repeatedly, the extra invested time it takes to refactor the code may be more than the overall time saved by the code running faster after being refactored. That is, it could take a team an hour to refactor a block of code, but if the new code only saved 20 seconds on compile and is only ever ran 100 times, the team lost 26 minutes overall making these changes. However, for code that may be run many times by many users, the overall time saved can be a tremendous victory. 
  Refactored code also easier to continue to work with, so for larger projects, spending the extra time to clean up code in the early stages means that it is easier to maintain and update the code in the future. 
  
  //how to those pros and cons apply to refactoring the original VBA script?

  In our initial project for Steve, he wanted to see the return on a small number of stocks from 2017 and 2018. This code could be run multiple times, but as the data was historical (and not likely to change by being ran again), there was no reason for Steve to care about saving microseconds by running this data. However, as he asked for an updated program to search the entire stock market over multiple years, we can see why he would then care that time is saved, as he may run this many more times and it would take longer to loop through thousands of tickers instead of dozens. 
  
  We were able to save 0.688 seconds for running through the 2017 data and 0.672 seconds for running through the 2018 data. This may not sound like much, but it was a roughly 86% time savings, as the refactored code ran in 14% of the time it took the original code to run. 
  



GRADING

=title, multiple paragraphs
=each paragraph has a heading
=subheadings break up text
=links and images formated and displayed correctly

=overview
=results, with screenshots and code
=summary with detailed statement on adv/disadv on refactoring code in general and for our example
