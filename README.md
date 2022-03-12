# VBA_Challenge

## Overview of Project

### Purpose of the analysis

We are helping Steve to design a code in VBA to output the stock market dataset with their Total Daily Volume and Return of the specific year by asking user to input the year. We also create some buttons that allowing user to perform the analysis by one click of button. After that we will easily see the comparison of different stocks by cells formatting and visualize which stocks are having greater return.

In the project we will have two sets of code. One is original code from the module we designed before and the other one is the refactoring code we design in this challenge. We will compare both codes to determine whether refactoring code successfully made the VBA script run faster. 

In terms of refactoring code, we are not changing the code or adding new functionality, we are making the code more structure, fewer steps, more efficient and a better experience for future user. We will analyze the advantages and the disadvantages for both original code and refactoring code at the end. 

## Results

The daily volume of each stock for the year was added up and put it under "Total Daily Volume"; also, I subtracted the starting price by the ending price of each stock for the year, and then minus 1 to get the return of each stock in percentage. In addition, having conditional formatting for the result that allows user to visualize easily the difference between the stocks.


!(/assets/images/VBA_Challenge_2017.png) !(/assets/images/VBA_Challenge_2018.png)

Based on the result as the pictures shown, in 2017 there is only one stock, TERP, having negative return -7.2%. All of the other stocks are having positive return with the range from 5.5% to 199.4%. Especially "DQ" is doing phenomenal in year of 2017, it got nearly 200% for return.

In 2018, there are only two stocks, ENPH and RUN, having positive return, 81.9% and 84.0%. The rest of the stocks are in negative return and "DQ", the stock with the highest return in 2017, also dropped 62% in year 2018.

Comparing 2017 and 2018, the stock performance in 2017 is much better than 2018.

The execution times of the original script for 2017 and 2018 year took 1.046875 and 0.953125 seconds and the refactoring code for 2017 and 2018 year took 0.1523438 and 0.1679688 seconds, which were 0.8945312 and 0.7851562 seconds quicker than the original code. In general, the refactoring code is 0.75 to 0.90 faster than the original code. It looks like a small number but for a bigger project there is a huge difference.

The reason the run time is much faster from refactoring code is because we used array method to store 3 important outputs: tickerVolume, tickerStartingPrice, and tickerEndingPrice that allows VBA to read the 3013 rows only once instead of going through all the 3013 rows 12 times in the nested loop from the original code.

## Summary

### Advantages and Disadvantages of refactoring code

- Advantages

1. The better design of the code helps programming run faster and less memory need.

2. Clean code makes the code easier to understand and less complex. 

3. Refactoring code makes it better organized and structured.

- Disadvantages

1. It takes extra time for re-structuring a different code. If the project is due soon, it is time consuming and may pass deadline of the project.

2. It may not be well-structured after refactoring and even create more bugs in the code.

3. Required more coding experiences and syntax knowledges to create a quality refactoring code and make the programming more efficient.

### Pro and cons apply to refactoring code

- Pros
  
  1. The structure of code looks easier to read and simple without the nested loop, it is better user experience for future user whoever read the code, or easier for anyone to understand the code.

  2. Time efficient. Refactored VBA script gets a faster result than original. It's just about 0.6-0.7 seconds faster but if it would be a huge difference for a bigger project.

- Cons

1. If the Ticker is not in order in the dataset, the result will not be correct.

Based on the refactored code in VBA, it will increase the tickerIndex by 1 if the next ticker is not the same as the previous ticker, but what if the ticker's name is not in order as we initialize in the beginning for the array of all tickers: (tickers(0) = "AY", tickers(1) = CSIQ, and so on. Therefore, we cannot get the correct result if the dataset is not pre-sorted.

2. It took extra time for code testing. I also received some overflow error messages in the beginning. For refactoring code, I need to edit the code and add some addition codes that may cause errors and mistakes.