# stock-analysis
Overview of Project: Explain the purpose of this analysis.
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?


##Purpose

The purpose of the stock analysis was to be able to view stock results from a 2 different worksheets by using refactored code to inscrease the speed of the process and efficiency.  The original code was made to only loop through rows on one data sheet and the refactored one was to be able to do many and utilize variables in such a way that we can perform multiple loops.

The refactored code ran in about ![VBA_Challenge_2017_runtime](https://user-images.githubusercontent.com/44000986/145735052-965f37b3-fd5e-435a-8bb3-a997e7b21fa0.png)
 .08 seconds and the original code ran in .89 seconds ![VBA_Challenge_originalcode_runtime](https://user-images.githubusercontent.com/44000986/145735062-c4de7fef-02de-4216-aea8-da13d5408078.png).  So the refactored code was much faster.  This is mainly due to the extra tickerIndex created to store and hold values for each column that needed to be filled, once conditional requirments were met.  
The analysis showed that stocks ENPH and RUN were the only two consistant stocks to provide a positive return.  The volumes of the all stocks overall were higher for 2018, however they had much less return than 2017 did.  

Using refactored code is very advantageous due to the fact that a large portion of the code can be reused and repurposed, so you dont have to start from scratch.  Making modification and adding to it would yeild much quicker results.  On the downside, refactored code can be very difficult to troubleshoot.  This is especially true if the refactored code and the original code were written by two different people.  This is where comments, code structure and formatting are especially useful.  In an instance where the same person who wrote the original is writing an refactored code, it will have more positive impacts than negative.  As any part of the code that contradicts another, or some missed variable could be found much quicker.  Also understanding of the original codes logic and end goal would be useful in refactoring it.  
