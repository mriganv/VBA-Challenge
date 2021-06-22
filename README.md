VBA Challenge
A VBA script used to analyze real stock market data in a Microsoft Excel workbook.

Background
For this project, I created a VBA (Visual Basic) script wchi will be running only once to analyze the stock market data. The data is inside a Microsoft Excel workbook and includes stock data for three years (2014, 2015, and 2016). Each year is a different tab/sheet inside the workbook. 

I created a script that will loop through all the stocks for three years of stock data that are in three different worksheets and will output 

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

*   Did a conditional formatting that will highlight positive change in green and negative change in red.

* The results are shown as below 

For the Bonus Challenge 

The script will return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" and also return the associated ticker symbol for the same in a separate table within each year sheet.

The solution look as follows:

Testing
I ran this script on both the testing Excel workbook (alphabetical_testing.xlsx) and on the final multiple year stock workbook (multiple_year_stock_data.xlsx). Script worked just fine for both and returns the output within few minutes.  


About the Script

You can find the script inside the VBAStocksdata folder of this repository. The script file is called VBA Script of stockdata



Here are screenshots of what the output looks like when I ran the scripts on my computer. These screenshots are also available in the VBAStocksdata/screenshots folder of this repository.

2014 Stock Data
Image of 2014 Stock Data


2015 Stock Data
Image of 2015 Stock Data

2016 Stock Data
Image of 2016 Stock Data
