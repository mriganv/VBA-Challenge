Sub Stockmarket()

'Setting a varialbe for worksheets
Dim Ws As Worksheet

' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
For Each Ws In ActiveWorkbook.Worksheets
Ws.Activate

  ' Set an initial variable for holding the Ticker name
  Dim Ticker_Name As String
 ' Set an initial variable for holding the Yearly change
  Dim Yearly_Change As Double
  ' Set an initial variable for holding the Percent change
  Dim Percent_Change As Double
  ' Keep track of the location for each values to be entered in the summary table
  Dim Summary_Table_Row As Integer
  'Set the initial Summary table row to start at the 2nd line
  Summary_Table_Row = 2
  'Set the initial Total volume to zero
  Total_Volume = 0
  'Set a variable for the lastrow count
  Dim LastRow As Long
  'Set a variable for the open price
  Dim Open_Price As Double
  'Set a variable for the open price
  Dim Close_Price As Double
  'Set a variable for the loop
  Dim i As Long
  
  'To find the Lastrow count
  LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  'Assigning a cell to get the Open Price value
  Open_Price = Cells(2, 3).Value
  
  'Set a table header for all the columns in the summary table
   Ws.Cells(1, "I").Value = "Ticker"
   Ws.Cells(1, "J").Value = "Yearly Change"
   Ws.Cells(1, "K").Value = "Percent Change"
   Ws.Cells(1, "L").Value = "Total Stock Volume"
   
' --------------------------------------------
' LOOP THROUGH ALL THE STOCK TICKERS
' --------------------------------------------
  
   For i = 2 To LastRow
   
  ' Check if we are still within the same Ticker Name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ' Set the Ticker Name
      Ticker_Name = Cells(i, 1).Value
      
    ' Print the Ticker Names in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name
          
    ' Assinging a close price value to grab from for every ticker value
      Close_Price = Ws.Cells(i, 6).Value
      
    ' To find the  Yearly Change
      Yearly_Change = Close_Price - Open_Price
      
    ' Print the Yearly_Change to the Summary Table
      Range("J" & Summary_Table_Row).Value = Yearly_Change
         
    ' Setting conditions for the percent change values
         'Check if both open price and close price values are zero
         If (Open_Price = 0 And Close_Price = 0) Then
         'Return a value of zero to Percent change
         Percent_Change = 0
         'Check if the open price value is 0 and Close Price as a value
         ElseIf (Open_Price = 0 And Close_Price <> 0) Then
         'Return a value of 1 to Percent change
         Percent_Change = 1
         'Check if both open price and close price as a value
         Else
         'Calculate the percent change
         Percent_Change = Yearly_Change / Open_Price
         
         'Ends this series of IF/ELSE conditionals
         End If
      
    ' Print the Percent Change to the Summary Table
      Range("K" & Summary_Table_Row).Value = Percent_Change
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
    ' Add to the Total Volume
      Total_Volume = Total_Volume + Ws.Cells(i, 7).Value
        
    ' Print the Total Volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_Volume
           
    ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
    ' check the next ticker open price value
      Open_Price = Cells(i + 1, 3).Value
      
    'Reset Total Volume
     Total_Volume = 0
      
    ' If the cell immediately following a row is the same Ticker Name...
    Else
        'Add to the Total Volume
        Total_Volume = Total_Volume + Ws.Cells(i, 7).Value
        

    End If

  Next i
  
  'Set a variable to find the maximum value of the Total Volume Column in the summary table
  Dim Max As LongLong
  'Set a variable to find the Greatestincerease value
  Dim Greatestincrease As Variant
  'Set a variable to find the Greatestdecrease value
  Dim Greatestdecrease As Variant
  'Set a variable to find the last row of the summary table
  Dim SummaryLastrow As Long
  'Set a variable for the loop
  Dim j As Long
  
  ' Find the last row of the summary table
  SummaryLastrow = Ws.Cells(Rows.Count, "L").End(xlUp).Row
    
' --------------------------------------------
' LOOP THROUGH THE YEARLY CHANGE COLUMN
' --------------------------------------------
  ' Loop through all the values in the summary table to change the colors based on Negative and positive values in each cell
  For c = 2 To SummaryLastrow
      'check if the values in each row of Yearly change column is negative
      If Cells(c, "J").Value < 0 Then
         'If the cell value is less than zero or negative set the cell color to red
         Ws.Cells(c, "J").Interior.ColorIndex = 3
      'if the values are positive
      Else
         'if the cell value in greater than zero or positive set the cell color to green
         Ws.Cells(c, "J").Interior.ColorIndex = 4
      'Ends the if/Elseif conditions
      End If
  Next c
  
 ' Assinging a name to each cells to enter the values of Greatest Incease, Greatest Decrease, Greatest Total Volume, Ticker and Value
  Ws.Cells(2, "O").Value = "Greatest % Increase"
  Ws.Cells(3, "O").Value = "Greatest % Decrease"
  Ws.Cells(4, "O").Value = "Greatest Total Volume"
  Ws.Cells(1, "P").Value = "Tikcer"
  Ws.Cells(1, "Q").Value = "Value"
  
  'Find the maximum value in the Total Stock Volume column
  Max = Application.WorksheetFunction.Max(Columns("L"))
  
  'Assign a cell for the max value in the total stock volume
  Ws.Cells(4, "Q").Value = Max
  
  'Find the maximum value in the Percent change column
  Greatestincrease = Application.WorksheetFunction.Max(Columns("K"))
  
  'Assign a cell for the max value
  Ws.Cells(2, "Q").Value = Greatestincrease

  'Find the minimum value in the Percent change column
  Greatestdecrease = Application.WorksheetFunction.Min(Columns("K"))
  
  'Assign a cell for the min value
  Ws.Cells(3, "Q").Value = Greatestdecrease
  
  'Formatting the cells of max and min values to show the %
  Range("Q2:Q3").NumberFormat = "0.00%"
  
' --------------------------------------------
' LOOP THROUGH THE SUMMARY TABLE
' --------------------------------------------
For j = 2 To SummaryLastrow
   ' Check if the Greatest total Volume number matches any of the cells in the Total Stock Volume rows
   If Ws.Cells(j, "L").Value = Max Then
      
      'if it matches, return the associated Ticker Name value to the Ticker cell in the column "P"
      Ws.Cells(4, "P").Value = Ws.Cells(j, "I").Value

      'Check if the Greatestincrease number matches any of the cells in the Percent change rows
   ElseIf Ws.Cells(j, "K").Value = Greatestincrease Then

          'if it matches, return the associated Ticker Name value to the Ticker cell in the column "P"
          Ws.Cells(2, "P").Value = Ws.Cells(j, "I").Value

          'Check if the Greatestdecrease number matches any of the cells in the Percent change rows
   ElseIf Ws.Cells(j, "K").Value = Greatestdecrease Then
          'if it matches, return the associated Ticker Name value to the Ticker cell in the column "P"
          Ws.Cells(3, "P").Value = Ws.Cells(j, "I").Value
    'Ends the if/Elseif conditions
    End If

Next j

Next

End Sub




