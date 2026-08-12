Sub Stock_Market_Analyst()

  ' // Revision Date:  2026.Aug.12

    ' Initialize core variables and working state for processing each worksheet
    Dim ws As Worksheet
    Dim last_row, i, output_row As Long
    Dim ticker_symbol As String
    Dim year_open, year_close, yearly_change, percent_change, total_stock_volume As Double
    Dim first_open_found As Boolean        ' // Upd:  2026.Aug.12

    ' Removed legacy overflow bypass; proper logic eliminates need for error suppression
    ' // Upd:  2026.Aug.12 (On Error Resume Next removed)

    ' Iterate through all worksheets to compute yearly metrics per ticker
    For Each ws In Worksheets

        ' Initialize output headers for ticker summary results
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Track the next available output row for summary data
        output_row = 2

        ' Determine the last populated row in the dataset
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Reset state for ticker-level accumulation
        total_stock_volume = 0
        first_open_found = False            ' // Upd:  2026.Aug.12

        ' Process each row of market data for the current worksheet
        For i = 2 To last_row

            ' Accumulate total volume for the active ticker
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value    ' // Upd:  2026.Aug.12 streamlined

            ' Capture the first valid opening price for percent-change calculations
            If Not first_open_found And ws.Cells(i, 3).Value <> 0 Then        ' // Upd:  2026.Aug.12
                year_open = ws.Cells(i, 3).Value
                first_open_found = True
            End If

            ' Detect ticker boundary (transition to next ticker)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Extract closing price for the current ticker
                year_close = ws.Cells(i, 6).Value

                ' Compute yearly and percent change, handling missing open values
                If first_open_found Then
                    yearly_change = year_close - year_open
                    percent_change = Round((yearly_change / year_open) * 100, 2)
                Else
                    yearly_change = 0
                    percent_change = 0
                End If

                ' Write computed metrics to the summary output table
                ws.Range("I" & output_row).Value = ws.Cells(i, 1).Value
                ws.Range("J" & output_row).Value = Round(yearly_change, 2)
                ws.Range("K" & output_row).Value = percent_change            ' // Upd:  2026.Aug.12 numeric output only
                ws.Range("L" & output_row).Value = total_stock_volume

                ' Apply conditional formatting based on yearly performance
                Select Case yearly_change
                    Case Is > 0
                        ws.Range("J" & output_row).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & output_row).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & output_row).Interior.ColorIndex = 0
                End Select

                ' Reset ticker-level state and advance output pointer
                total_stock_volume = 0
                first_open_found = False        ' // Upd:  2026.Aug.12
                output_row = output_row + 1     ' // Upd:  2026.Aug.12 simplified increment

            End If

        Next i

    Next ws

End Sub
