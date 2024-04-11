# Module-2-Challenge-VBA

Helllo, 
in this repository you will find my homework for the Module 2 Challenge. This includes screenshots and a VBA script file.
This VBA script file has been made with the help of differents sources: shrawantee(Github), Xpert Learning Assistant, Tutor & and myself Amanda Lor



Instructions for the homework: 

"1)Create a script that loops through all the stocks for one year and outputs the following information:
-The ticker symbol
-Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
-The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
-The total stock volume of the stock. 

 2)Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 

 3)Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

 Note: Make sure to use conditional formatting that will highlight positive change in green and negative change in red."
 
 

 Breakdown of the homework:

1)Determine the variables

2)Label the Summary Table

3)Loop

-Volume: Total_SVolume = Total_SVolume + Cells(i, 7).Value

-Ticker Symbol: Ticker_Symbol = Cells(i, 1).Value

-Yearly Change: Yearly_Change = (Closing_Price - Opening_Price)

4)Add colors: Green for positive and Red for negative

5)Find greatest % increase, greatest % decrease and total volume:

-Maximum percent change: Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table))

-Minimum percent change: Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table))

-Total Volume: Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table))



Thank you for reading!
 
