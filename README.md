# Stock Analysis
## Overview of Project
This research aims to look at the performance of a variety of stocks over two years, 2017 and 2018, to assist Steve in determining which stocks would be best for his parents' investment aspirations. Moreover, we will rewrite portions of the VBA script and determine whether refactoring the original code made the VBA script run more efficiently.
## Results
### Refactoring Original VBA Script.
Changes to the original script include:

**1. New tickerindex is set to zero before looping over the rows.**

I created a tickerindex variable and set it equal zero before iterating over all the rows. The tickerIndex provides numerical value to evaluate the four arrays: the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

![TickerIndex](https://github.com/msevillano89/stock-analysis/blob/main/Resources/TickerIndex.png)

**2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.**

I created arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices and defined each variable as follows:

![Arrays for tickerindex](https://github.com/msevillano89/stock-analysis/blob/main/Resources/Arrays.png)

**3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.**

I created a conditional statement to increase the tickerIndex if the next row's ticker doesn't match. The value of the tickerIndex is used in conjunction with the tickers variable to determine which index is evaluated.

tickers:

![Tickers definition](https://github.com/msevillano89/stock-analysis/blob/main/Resources/Tickers.png)

tickerindex conditional statement:

![tickerindex conditional statement](https://github.com/msevillano89/stock-analysis/blob/main/Resources/TickerIndex%20Increase.png)

**4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.**

I created a set of loops throughout the rows of a specific spreadsheet. Within the loop, I made a set of conditionals to calculate the total volume (tickerVolumes) and determine each index's Starting and Ending Prices.

![Conditional statement for Arrays](https://github.com/msevillano89/stock-analysis/blob/main/Resources/Arrays%20Increase.png)

### Stock performance between 2017 and 2018
As noted in the image below, there was a significant market crash in 2018. Most stock prices fell by 27%, with "SPWR "taking the biggest hit, with a return decrease of 293%. However, two stocks outperformed others in spite of the dire market conditions, "RUN" and "ENPH," whose returns increase by 83%. Therefore, Steve's parents should consider investing in the latter.

![Stock Performance](https://github.com/msevillano89/stock-analysis/blob/main/Resources/2017-2018_Stock%20Performance.png)

### Execution times of the original script and the refactored script.
**Original execution times**
* 2017

![Original Script 2017](https://github.com/msevillano89/stock-analysis/blob/main/Resources/VBA_Original_2017.png)

* 2018

![Original Script 2018](https://github.com/msevillano89/stock-analysis/blob/main/Resources/VBA_Original_2018.png)

**Refactored execution times**
* 2017

![Refactored Script 2017](https://github.com/msevillano89/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

* 2018

![Refactored Script 2018](https://github.com/msevillano89/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

As seen above, we can see that refactoring the VBA script reduced the run time by an average of .50 seconds.Improving the overall performance of the code. 

## Summary
### Advantages of Refactoring
One of the advantages of refactoring is that redundancies and duplications improve the code's overall effectiveness. By eliminating unnecessary lines of code, we also enhance its legibility, making it easier for other programmers to comprehend each part of the code.

The improved execution time of our VBA script can serve as an example of how refactoring can improve the run time (effectiveness) of a script. Moreover, adding additional commentary to the code will ease the ability of other programmers to reuse or troubleshoot the code in case of debugging.  
 
### Disadvantages of Refactoring
The quest of refactoring a code can lead to various problems. The refactoring process can introduce new bugs and errors, bringing frustration and discouragement for the novice coder. Also, the concept of a "cleaner code" is subjective. What can be easier for one person to follow, others could find it other to comprehend.

For example, in our refactored VBA, I found it very difficult to properly understand the purpose tickerIndex variable and how to use it in the new script properly. Not knowing how to use the variable led to an increasing number of errors and bugs every time I tried to run the script, resulting in more time I had to spend trying to fix code to get the script to run correctly.
