# VBA Challenge

## Overview of Project
The overview of this analysis was to find out if the client's parents should invest their money into a stock besides DQ. We have to do this by expanding the current data set to include the entire stock market.

### Purpose
The purpose of this analysis is to refacory the solution code and successfully make the VBA script run faster.

## Results
    In this analysis, I created a ticker index variable and set it equal to zero before iterating the rows. I, then created three output arrays: ticker volumes, ticker starting prices, and ticker ending prices.

    Next, I created a for loop to intialize the ticker volumes. Following that, a for loop that looped over all the rows in the spreadsheet.

    Inside the for loop, I wrote a script that increases the current stock ticker volume variables and adds the ticker volume for the current stock ticker.


    Next, I began the if statements to check and make sure that current row is the first row. Then, assigned the current starting price to the variable.

    The next if then statement checks if the current row is the last row. Then, assigned the current ending prices variable.

    Next, I wrote a script that increases the ticker index if the next row's ticker doesn't match the previous row's ticker.

    Lastly, I used a for loop to loop through the arrays. These are the tickers, tickerVolumes, starting and ending prices. This was to output the Ticker, Total Daily Volumes, and Return columns in the spreadsheet.

    Finally, I ran the stock analysis. However, in the box, if I put a year (2018) it doesn't register, it just pops the same window back up.



## Summary


- Detailed statement on advantages and disadvantages of refactoring code in general.
1. The advantages of refactoring code is that it is much easier to edit and update as you go along. You can easily understand it as well because it lays it out in outlines, so easier to maintain.
2. The disadvantages is that it takes so much time to do. It also creates new errors that you have to figure out by searching on the web or trial and error.

- Detailed statement on the advantages and disadvantages or the original and refactored VBA script
1. The advantages of the original VBA script is that it is stable and it works. There are less errors. If you have a tight deadline this would be the way to go, not refactoring.
2. The disadvantages of the original VBA script is that it ran slower than the refactored one.