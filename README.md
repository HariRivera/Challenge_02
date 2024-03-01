# Challenge_02
Challenge #2 VBA Hari Rivera
How I developed each solution

##The StockDataSolution 
The macro is designed to go through all the sheets in the workbook challenge, analyzing and summarizing financial stock data. The goal is to process information about different stocks (tickers) and calculate the yearly change, annual percentage change, and total stock volume traded for each one. These calculations are performed for each "ticker"

Iterating Through Sheets: The loop For Each ws In ThisWorkbook.Sheets allows the code to run on each sheet in the active workbook. This ensures that the calculations are performed across all sheets.

Variable Initialization: At the start of each sheet iteration, key variables such as lastRow (the last row with data), outputRow (the output row for the results), and totalVolume (the total volume of traded stocks) are placed wiht in the code.

Results Table Headers: Headers for columns "I," "J," "K," and "L" are set for 'Ticker', 'Yearly Change', 'Percentage Change', and 'Total Stock Volume', respectively. This prepares the sheet to receive the calculations.

Finding the Opening Price: The first opening price of the year (startPrice) is identified for each ticker. This is used as a basis for calculating the yearly change.

Calculating Results per Ticker: Within the loop For i = 2 To lastRow, the script checks each row to determine if it corresponds to a new ticker or if it's a continuation of the same one. Upon completing a ticker (or at the end of the data), the yearly change, percentage change, and total volume are calculated.

Yearly Change and Percentage Change: These are calculated by subtracting the startPrice from the closing price of the last recorded day (endPrice) and dividing the difference by the startPrice, respectively.

Total Stock Volume: This is accumulated by adding up the daily volumes.

Logging Results: The results for each ticker are recorded in columns "I," "J," "K," and "L" of the current sheet, starting from the defined outputRow. This includes the ticker symbol, yearly change, percentage change, and total stock volume.

Preparation for the Next Ticker: Upon finishing with one ticker, the totalVolume is reset, and the new startPrice is set for the next ticker, ensuring that the calculations are independent for each one.

##The Value_Table macro
This code is crafted to go through each sheet of the challenge, seeking to identify de results from the previous code within the stock data: the greatest percentage increase, the greatest percentage decrease, and the greatest total volume of stock traded. These calculations are for financial analysis, offering insights into stock performance over time. Below is a detailed explanation of the macro's structure and logic:

Sheet Iteration: The macro initiates by iterating through every sheet within the workbook with the loop For Each ws In ThisWorkbook.Sheets.

Initializing Variables: For each sheet, the macro initializes variables to track the last row with data (lastRow), the greatest increase and decrease in stock value (greatestIncrease, greatestDecrease), and the greatest volume (greatestVolume). It also prepares strings to store the tickers associated with these records.

Setting Header Labels: At the top of designated columns ("P" for tickers and "Q" for values), headers are set to categorize the information clearly, facilitating easy reading and understanding of the output data.

Finding Superlatives: The loop For i = 2 To lastRow, where it examines each row's percentage change and total volume values. It compares each entry against the current records for the highest increase, the highest decrease, and the highest volume, updating these records as greater values are found. This comparison is based on the values in columns "K" for percentage changes and "L" for total volumes.

Greatest Increase and Decrease: The macro identifies the highest and lowest percentage changes in stock value. Uniquely, greatestDecrease is initially set to 1, assuming percentage change is represented as a decimal, to ensure any actual decrease is captured.

Greatest Volume: It identifies the stock with the highest trading volume, indicative of significant trading activity.

Recording Results: Upon completing the iterations, the macro records the results in a new section of the worksheet. It writes the labels "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume" along with the associated ticker symbols and values. 

Loop Continuation for All Sheets: After processing a sheet, the macro proceeds to the next sheet in the workbook, repeating the analysis, and extending across an entire workbook.

After executing this code, it delivered the next results:

In 2018
Greatest % Increase was for Ticker THB with a 141.42% increase
Greatest % Decrease was for Ticker RKS with a -90.02% decrease
Greatest Total Volume was for Ticker QKN with a total volume of 1689539560106

In 2019
Greatest % Increase was for Ticker TYU with a 190.03% increase
Greatest % Decrease was for Ticker RKS with a -91.60% decrease
Greatest Total Volume was for Ticker ZQD with a total volume of 4373008528422

In 2020
Greatest % Increase was for Ticker YDI with a 188.76% increase
Greatest % Decrease was for Ticker VNG with a -89.05% decrease
Greatest Total Volume was for Ticker QKN with a total volume of 3452956568861
