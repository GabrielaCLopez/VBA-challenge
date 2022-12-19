###### Stock Market Analysis

***Define variables***
The first step I took was defining all the variables I would be using in the code. I defined i and lastrow as Long. Ticker was a string. Closingprice, openingprice, start, greatestdecrease, greatestincrease, and greatest volume were Double. Stockvolume was LongLong.

***Create tables***
Next, I created by tables by assigning the cells with the names of the titles I chose.

***Define opening price***
I next defined the openingprice so the closingprice could be compared to it when the annual changes need to be calculated. 

***Set your starting position***
Start the loop at 2 because the first row is the header.

***Determine the last row***
I used the formula to determine the lastrow so that the code would loop through every value present.

***Define starting volume***
I set the stockvolume to zero so the actual volume could be calcuated within the loop.

***Start the loop***
I set i to go from 2 to the lastrow. I set an equation to calcuate the stockvolumn of each entry for that ticker. The ticker was also set to be in the first column.

Next, I set a conditional statement that would note when the ticker values are no longer the same. This would cause the summary table to get filled each time the ticker changed. The closing price was set within the loop. Using the closing and opening prices, the change was calculated. Then, this was formatted into a percentage with two decimal places.

The volume would then be added to the summary table. 

I added another conditional statement to change the colors of the cells with the percent change. If the change was positive, the cell would turn green. If the change was negative, the cell would turn red.

After this, the colume would reset to zero, and the start position would move one cell. The opening price was then set to start one cell further down as well. 

***Finding the greatest values***
To find the greatest values, I used the max and min worksheet functions to look through the percent changes and the volume columns to find the max and min values. 

These were then assigned to the greatest values table. The corresponding tickers were selected using the match function. These were then assigned to the greatest values table.

***Looping through the worksheets***
After determining that the code was successful on the first worksheet, define the worksheet as a variable and create the loop before the first variables. 

Add ws. to any value that references a position on the worksheet (i.e., cells and ranges). Put Next ws after Next i to finish the nested loop.






