# VBA_challenge
VBA Challenge Activity

#Assignment Overview

We are going to create new automated process using VBA to loop through all stocks for a year and output the certain information:

## Step 1

As we index through the rows, we look for a change in the ticker symbol in the first column to give us the range for each given ticker.

## Step 2

Using logic to grab values from the most recent change in ticker to the next change in ticker, we can get the opening and closing values, taking the difference of the two to get the change.

## Step 3

Using green and red coloration, we then designate whether the yearly change is positive or negative. We can also then produce a percentage change by dividing the change by our starting value for that given ticker.

## Step 4

While indexing through the tickers, we are adding up the volume and resetting it to zero every time we have a change in ticker, collecting the total stock volume over the year.

## Step 5

Finally, we add MaxUp, MaxDown, and MaxVol values, all starting at zero to be replaced as we index through and have a larger/lower value.

<hr/r>

- Step 1: read the file
- Create new model

PUSH
- add: git add .
- commit with message: git commit -m "added new test"
- then push: git push

PULL
- git pull

Check on things:
- git status

