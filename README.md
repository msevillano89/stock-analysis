# Stock Analysis
## Overview of Project
This research aims to look at the performance of a variety of stocks over two years, 2017 and 2018, to assist Steve in determining which stocks would be best for his parents' investment aspirations. Moreover, we will rewrite portions of the VBA script and determine whether refactoring the original code made the VBA script run more efficiently.
## Results
### Refactoring Original VBA Script.
Changes to the original script include:

**1. New tickerIndex is set to zero before looping over the rows.**
I created a tickerIndex variable and set it equal zero before iterating over all the rows. The tickerIndex provides numerical value to evaluate the four arrays: the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

**2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.**
I created arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices and defined each variable as follows:
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.
I created a conditional statement to increase the tickerIndex if the next row's ticker doesn't match. The value of the tickerIndex is used in conjunction with the tickers variable to determine which index is evaluated.
'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
         
         tickerEndingPrices(tickerindex) = Cells(i, 6).Value
        
        End If
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            
            tickerindex = tickerindex + 1
The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
I created a set of loops throughout the rows of a specific spreadsheet. Within the loop, I made a set of conditionals to calculate the total volume (tickerVolumes) and determine each index's Starting and Ending Prices.

'2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerindex) Then

               tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value

        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
         
         tickerStartingPrices(tickerindex) = Cells(i, 6).Value
         
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
         
         tickerEndingPrices(tickerindex) = Cells(i, 6).Value
        
        End If

Stock performance between 2017 and 2018
As noted in the image below, there was a significant market crash in 2018. Most stock prices fell by 27%, with "SPWR "taking the biggest, with a price decrease of 293%. However, two stocks outperform the other against the dire market conditions, "RUN" and "ENPH," whose returns increase by 83%. Therefore, Steve's parents should consider investing in the latter.
[Add Image]

Execution times of the original script and the refactored script.

Original execution times
2017
[Image]
2018
[Image]

Refactored execution times
2017
[Image]
2018
[Image]

As seen above, we can see that refactoring the codes improved execution time by reducing run time by an average of .50 seconds. 

Summary

Advantages of Refactoring
One of the advantages of refactoring code is that redundancies and duplications improve the code's overall effectiveness. By eliminating unnecessary lines of code, we also enhance its legibility, making it easier for other programmers to comprehend each part of the code.

The improved execution time of our VBA script can serve as an example of how refactoring can improve the run time (effectiveness) of a script. Moreover, adding additional commentary to the code will ease the ability of other programmers to reuse or troubleshoot the code in case of debugging.  
 
Disadvantages of Refactoring
The quest of refactoring a code can lead to various problems. The refactoring process can introduce new bugs and errors, bringing frustration and discouragement for the novice coder. Also, the concept of a "cleaner code" is subjective. What can be easier for one person to follow, others could find it other to comprehend.

For example, in our refactored VBA, I found it very difficult to properly understand the purpose tickerIndex variable and how to use it in the new script properly. Not knowing how to use the variable led to an increasing number of errors and bugs every time I tried to run the script, resulting in more time I had to spend trying to fix code to get the script to run correctly. 
