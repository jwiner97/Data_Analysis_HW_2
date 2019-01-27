Sub Stocks()
'Select new worksheet
Dim ws As Worksheet
  For Each ws In Worksheets
    ws.Select

        
        
'Declare variables
Dim Ticker As String
Dim Volume As Double
Dim SummaryTableRow As Integer
Dim year_open As Double
Dim year_close As Double
Dim GIValue As Double
Dim GITicker As String
Dim GDValue As Double
Dim GDTicker As String
Dim GVValue As Double
Dim GVTicker As String
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Set Variables on each sheet to zero
Volume = 0
year_open = 0
GIValue = 0
GDValue = 0
GVValue = 0
'Insert headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Chage"
Cells(1, 12).Value = "Total Stock Vol"
Cells(1, 11).Value = "Percent Change"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Set start of summary table
SummaryTableRow = 2

'Loops through rows 2 to 71226 as the worksheet with the most data was C at 71226
    For i = 2 To LastRow
    'Look for the first opening price for the first ticker and set that as opening price
      If year_open = 0 Then
          year_open = Cells(i, 3).Value
      End If
      
      'Cycle Through Data and totals up Volume, grabs the year close price on the very last day to calculate changes and percentages
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          Ticker = Cells(i, 1).Value
          year_close = Cells(i, 6).Value
          Dollar_Change = year_close - year_open
          'ISSUE: CANNOT HANDLE 0/0. Solved this by using below conditional to substitute 0 for stock with a denominator of 0
          If year_open <> 0 Then
            Percent_Change = FormatPercent(Dollar_Change / year_open)
             Else: Percent_Change = 0
            End If
          Volume = Volume + Cells(i, 7).Value
 'Conditional formatting to make positive changes green and negative changes red
            If Dollar_Change > 0 Then
            Range("J" & SummaryTableRow).Interior.ColorIndex = 4
               Else
            Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                 End If
'Input data from variables above into summary table
        Range("I" & SummaryTableRow).Value = Ticker
        Range("J" & SummaryTableRow).Value = Dollar_Change
        Range("K" & SummaryTableRow).Value = Percent_Change
        Range("L" & SummaryTableRow).Value = Volume
'Counts the SummaryTableRow to reset to next row for next iteration
          SummaryTableRow = SummaryTableRow + 1

'Reset Volume and Year_Open to Zero
          Volume = 0
          
          year_open = 0

'If tickers are the same on contiguous rows then add up the volume and keep a running total
      Else

          Volume = Volume + Cells(i, 7).Value


      End If


    Next i
'Challenge Code Section

'This iterates through the summary table portion of each tab.
'I reset the GIValue (Greatest % Increase) and similar variables to 0 on every new spreadsheet.
'If the for loop finds a value in the percentage column that is greater than the amount stored in the current GIValue,
'then it assigns this new price to the GIValue
EndSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
For j = 2 To EndSummary
    If Cells(j, 11).Value > GIValue Then
    GIValue = Cells(j, 11).Value
    GITicker = Cells(j, 9).Value
    Range("P2").Value = GITicker
    Range("Q2").Value = FormatPercent(GIValue)
    End If
  If Cells(j, 11).Value < GDValue Then
    GDValue = Cells(j, 11).Value
    GDTicker = Cells(j, 9).Value
    Range("P3").Value = GDTicker
    Range("Q3").Value = FormatPercent(GDValue)
    End If
    If Cells(j, 12).Value > GVValue Then
    GVValue = Cells(j, 12).Value
    GVTicker = Cells(j, 9).Value
    Range("P4").Value = GVTicker
    Range("Q4").Value = GVValue
    End If
Next j

ws.Columns("I:Q").AutoFit

Next ws

End Sub


