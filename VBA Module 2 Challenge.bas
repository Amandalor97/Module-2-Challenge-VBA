Sub StockTickerSymbolAnalysis()

'This code has been made thanks to differents sources: shrawantee(Github), Xpert Learning Assistant and myself Amanda Lor

'Let's determine the variables
Dim Ticker_Symbol As String
Dim Total_SVolume As Double
        Total_SVolume = 0
Dim Ticker_Summary As Integer
        Ticker_Summary = 2
Dim Opening_Price As Double
 Opening_Price = Cells(2, 3).Value
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

For Each ws In Worksheets
ws.Activate

'Let's label the Summary Table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Let's start the loop
For i = 2 To lastRow

'Volume
Total_SVolume = Total_SVolume + Cells(i, 7).Value


'Let's look at when the next cell value differs from the one before
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Ticket name
Ticker_Symbol = Cells(i, 1).Value
Range("I" & Ticker_Summary).Value = Ticker_Symbol

Range("L" & Ticker_Summary).Value = Total_SVolume

'Yearly Change
Closing_Price = Cells(i, 6).Value
Yearly_Change = (Closing_Price - Opening_Price)
If (Opening_Price = 0) Then

                    Percent_Change = 0

                Else
                    
                    Percent_Change = Yearly_Change / Opening_Price
                
                End If
Range("K" & Ticker_Summary).Value = Percent_Change
Range("K" & Ticker_Summary).NumberFormat = "0.00%"

'Let's reset the row counter and add +1
Ticker_Summary = Ticker_Summary + 1

'Let's reset the volume
Total_SVolume = 0

'Let's reset the Opening Price and add the volume
Opening_Price = Cells(i + 1, 3)
Else
            'Let's add the volume
              tickervolume = Total_SVolume + Cells(i, 7).Value

            
            End If
        
Next i

'Let's add colors! Green for positive and Red for negative
lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
Next i

'Let's find the greatest % increase, % decrease and total volume
Dim maxIncrease As Double
Dim maxDecrease As Double
Dim maxVolume As Double
Dim maxIncreaseTicker As String
Dim maxDecreaseTicker As String
Dim maxVolumeTicker As String

'Let's add names to the cells
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

For i = 2 To lastrow_summary_table

'Maximum Percent Change
If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
   Cells(2, 16).Value = Cells(i, 9).Value
   Cells(2, 17).Value = Cells(i, 11).Value
   Cells(2, 17).NumberFormat = "0.00%"

'Minimum Percent Change
ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
       Cells(3, 16).Value = Cells(i, 9).Value
       Cells(3, 17).Value = Cells(i, 11).Value
       Cells(3, 17).NumberFormat = "0.00%"

'Maximum Volume
ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
       Cells(4, 16).Value = Cells(i, 9).Value
       Cells(4, 17).Value = Cells(i, 12).Value
       Cells(4, 17).NumberFormat = "0.00E+00"
            
    End If
        
Next i
Next ws
        
End Sub

