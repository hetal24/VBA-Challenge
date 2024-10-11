# VBA-Challenge
# VBA Multiple year stock Data

In this activity, you will create a macro enable excel file that will check a Stocks ticker, quaterly change, percentage change, total volume ,Greatest % increase,Greatest % decrease,Greatest total volume according the given data.

## Instructions

With `Multiple_year_stock_data.xlsm` as your starting point,

* create a for loop which iterate all worksheets.

* create a other for loop which iterate all given stock data.

Once complete, your script should perform the following:

* Now check the ticker if the one ticker is not equal to other ticker.

* Check the quaterly change according to the ticker name for that use the open price of the ticker and close price of the ticker then do the difference of the ticker.
  (Ex-close price-open price)

* Check the percentage change for that use use (quaterly change/open price*100) 

* Also apply condition formatting on percentage change column where cell color change according to their negative and positive value.

* check the total volume according to the ticker name if ticker didn't match then check else part.(total volume=temp+volume)

* Check the Greatest % increase, Greatest % decrease, and Greatest total volume for that check if condition using the given stock data.

