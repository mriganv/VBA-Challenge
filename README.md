
VBA Challenge

A VBA script used to analyze real stock market data in a Microsoft Excel workbook.

Background
For this project, I created a VBA (Visual Basic) script which will be running only once to analyze the stock market data. The data is inside a Microsoft Excel workbook and includes stock data for three years (2014, 2015, and 2016). Each year is a different tab/sheet inside the workbook. 

I created a script that will loop through all the stocks for three years of stock data that are in three different worksheets and will output 

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

*   Did a conditional formatting that will highlight positive change in green and negative change in red.

* The results are shown and attached as images below

For the Bonus Challenge 

The script will return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" and also return the associated ticker symbol for the same in a separate table within each year sheet.

Testing
I ran this script on both the testing Excel workbook (alphabetical_testing.xlsx) and on the final multiple year stock workbook (multiple_year_stock_data.xlsx). Script worked just fine for both and returns the output within few minutes.  


About the Script

You can find the script submitted as vbs script file and its called VBA Script of stockdata

Here are screenshots of what the output looks like when I ran the scripts on my computer. These screenshots are also available in the Images of Multi Year Stock folder of this repository.

2014 Stock Data

![Multi Year Stock - 2014](https://user-images.githubusercontent.com/81407869/122868668-1d3d5c80-d2e0-11eb-9394-e255a09795d3.png)


2015 Stock Data

![Multi Year Stock - 2015](https://user-images.githubusercontent.com/81407869/122869772-95f0e880-d2e1-11eb-8867-09c1a1bf5d1a.png)



2016 Stock Data

![Multi Year Stock - 2016](https://user-images.githubusercontent.com/81407869/122868736-380fd100-d2e0-11eb-8b75-5f38671cbd0f.png)
