Sub Bonus_Summary()

  ' // Revision Date:  2026.Aug.12
  
    ' Initialize worksheet-level summary processing for all output sheets
    Dim ws As Worksheet
    Dim last_row, k As Long

    ' Declare variables for tracking max/min metrics across tickers
    Dim greatestTicker, leastTicker, volumeTicker As String
    Dim greatestIncrease, greatestDecrease, greatestVolume As Double

    ' Iterate through each worksheet to compute summary statistics
    For Each ws In Worksheets

        ' Define header labels for summary metrics and output fields
        ws.Cells(2, 14).Value = "Greatest % Increase:"
        ws.Cells(3, 14).Value = "Greatest % Decrease:"
        ws.Cells(4, 14).Value = "Greatest Total Volume:"
        ws.Cells(1, 15).Value = "Ticker Symbol"
        ws.Cells(1, 16).Value = "Value"

        ' Apply formatting to ensure consistent numeric and alignment presentation
        ws.Columns("J").NumberFormat = "$0.00"
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("J:L").HorizontalAlignment = xlRight

        ' Determine the last populated row in the summary output region
        last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

        ' // Upd:  2026.Aug.12 Skip sheets without summary data
        If last_row < 2 Then GoTo NextSheet

        ' Initialize baseline comparison values using the first summary row
        greatestIncrease = ws.Cells(2, 11).Value
        greatestDecrease = ws.Cells(2, 11).Value
        greatestVolume = ws.Cells(2, 12).Value        ' // Upd:  2026.Aug.12 corrected initialization

        greatestTicker = ws.Cells(2, 9).Value
        leastTicker = ws.Cells(2, 9).Value
        volumeTicker = ws.Cells(2, 9).Value

        ' Iterate through all summary rows to compute max/min metrics
        For k = 2 To last_row

            ' Evaluate greatest percent increase across tickers
            If ws.Cells(k, 11).Value > greatestIncrease Then
                greatestIncrease = ws.Cells(k, 11).Value
                greatestTicker = ws.Cells(k, 9).Value
            End If

            ' Evaluate greatest percent decrease across tickers
            If ws.Cells(k, 11).Value < greatestDecrease Then
                greatestDecrease = ws.Cells(k, 11).Value
                leastTicker = ws.Cells(k, 9).Value
            End If

            ' Evaluate highest total volume across tickers
            If ws.Cells(k, 12).Value > greatestVolume Then
                greatestVolume = ws.Cells(k, 12).Value
                volumeTicker = ws.Cells(k, 9).Value
            End If

        Next k

        ' Write computed summary metrics back to the worksheet summary section
        ws.Cells(2, 15).Value = greatestTicker
        ws.Cells(2, 16).Value = greatestIncrease

        ws.Cells(3, 15).Value = leastTicker
        ws.Cells(3, 16).Value = greatestDecrease

        ws.Cells(4, 15).Value = volumeTicker
        ws.Cells(4, 16).Value = greatestVolume

        ' Apply formatting to summary values for readability and consistency
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Columns("N:P").EntireColumn.AutoFit
        ws.Columns("P").HorizontalAlignment = xlRight

NextSheet:
    Next ws

End Sub
