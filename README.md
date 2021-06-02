# VBA Project: &nbsp;The VBA of Wall Street

## Background

In this project, we used VBA scripting to analyze real stock market data.

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Used this to develop scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Ran scripts on this data to generate the final report.

### Stock Market Analyst

![stock Market](Images/stockmarket.jpg)

## Steps

* Created a script that loops through all stocks for one year and outputs the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Applied conditional formatting that highlights positive change in green and negative change in red.

* The result should look as follows:

![moderate_solution](Images/moderate_solution.png)

## Additional Analysis

* The solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". &nbsp;This solution will look as follows:

![hard_solution](Images/hard_solution.png)

* Made the appropriate adjustments to the VBA script that allows it to run on every worksheet (i.e. every year) just by running the VBA script once.
