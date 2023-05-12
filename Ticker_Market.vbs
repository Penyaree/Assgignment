Attribute VB_Name = "Module3"
Sub Ticker_Market()
'Create the column headings
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest%Increase"
Range("O3").Value = "Gretest%Decrease"
Range("O4").Value = "Greatest Total Value"

'Define Ticker variable
Dim Tickername As String

'Set a variable to hold the total volume of ticker
Dim tickerVolume As Double
tickerVolume = 0

'Set Value to 0 for Gretest%
Range("Q2").Value = 0
Range("Q2").NumberFormat = "0.00%"
Range("Q3").Value = 0
Range("Q3").NumberFormat = "0.00%"
Range("Q4").Value = 0
Range("Q4").NumberFormat = "#,###,###,###,###"

'Set new variable for prices and percent changes
Dim open_price As Double
open_price = Cells(2, 3).Value

Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

Dim summary_ticker_row As Integer
summary_ticker_row = 2
'Define Lastrow of worksheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'Loop trough the rows by the ticker names
'Set initial and last row for worksheet

Dim I As Long
Dim j As Integer
For I = 2 To LastRow
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        'Set the ticker name
        Tickername = Cells(I, 1).Value
        tickerVolume = tickerVolume + Cells(I, 7).Value
        'Print the ticker name in the summary table
        Range("I" & summary_ticker_row).Value = Tickername
        'Print the trade volume for each ticker in the summary table
        Total_Stock_volume = tickerVolume
        Range("L" & summary_ticker_row).Value = Total_Stock_volume
        'Collect information about closing price
        close_price = Cells(I, 6).Value
        'Calculate yearly change
        yearly_change = (close_price - open_price)
        'Print the yearly change for each ticker in the summary table
        Range("J" & summary_ticker_row).Value = yearly_change
        
        'Check for the non-divisibility condition when calculating the percent change
            If (open_price = 0) Then
                percent_change = 0
            Else
            
                percent_change = yearly_change / open_price
            
            End If
    
        'Print the yearly change for each ticker in the summary table
        Range("K" & summary_ticker_row).Value = percent_change
        Range("K" & summary_ticker_row).NumberFormat = "0.00%"
        
        'Set % Increase or Decrease
        If percent_change > Range("Q2").Value Then
            Range("P2").Value = Tickername
            Range("Q2").Value = percent_change
        End If
        
        If percent_change < Range("Q3").Value Then
            Range("P3").Value = Tickername
            Range("Q3").Value = percent_change
        
        End If
        
        If tickerVolume > Range("Q4").Value Then
            Range("P4").Value = Tickername
            Range("Q4").Value = tickerVolume
        End If
        
        
        'Reset the row counter.Add one to the summary_ticker_row
        summary_ticker_row = summary_ticker_row + 1
        'Reset volume of trade to zero
        tickerVolume = 0
        'Reset the opening price
        open_price = Cells(I + 1, 3)

    Else

        'Add the volume of trade
        tickerVolume = tickerVolume + Cells(I, 7).Value
    
    End If

    Next I
'Conditional farmatting the will highlight positive change in green and negative chnge in red
'Find the last row of the summary table
lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row

'Color code yearly change
For I = 2 To lastrow_summary_table
 If Cells(I, 10).Value > 0 Then
    Cells(I, 10).Interior.ColorIndex = 10
    
    Else
    Cells(I, 10).Interior.ColorIndex = 3
 
    
    End If
    
 Next I
 
End Sub
