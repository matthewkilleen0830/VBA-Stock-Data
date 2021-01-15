Sub Bonus_Summary():

    'Bonus summary output
    
    'Loop through output rows of data in all worksheets in workbook
    
    For Each ws In Worksheets
    
        'Write text to summary output column headers
        ws.Cells(2, 14).Value = "Greatest % Increase:"
        ws.Cells(3, 14).Value = "Greatest % Decrease:"
        ws.Cells(4, 14).Value = "Greatest Total Volume:"
        ws.Cells(1, 15).Value = "Ticker Symbol"
        ws.Cells(1, 16).Value = "Value"
    
        'Format alignment and number format for values in cells in summary output columns
        ws.Columns("J").NumberFormat = "$0.00"
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("J:L").HorizontalAlignment = xlRight
    
        'Declare variable to find the last row of data from output columns
        last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        'Declare variables and set initial values
        Dim greatest_ticker_symbol As String
        
        Dim greatest_increase As Double
        greatest_increase = ws.Cells(2, 11).Value
        
        Dim least_ticker_symbol As String
        
        Dim most_decrease As Double
        most_decrease = ws.Cells(2, 11).Value
        
        Dim greatest_ticker_volume As String
        
        Dim greatest_total_volume As Double
        greatest_ticker_volume = ws.Cells(2, 12).Value
        
        'Loop through all rows of data in output columns
        For k = 2 To last_row
        
            'Conditional statement to locate greatest % increase
            If ws.Cells(k, 11).Value > greatest_increase Then
                
                greatest_increase = ws.Cells(k, 11).Value
                greatest_ticker_symbol = ws.Cells(k, 9).Value
            
            End If
            
            'Conditional statement to locate greatest % decrease
            If ws.Cells(k, 11).Value < most_decrease Then
            
                most_decrease = ws.Cells(k, 11).Value
                least_ticker_symbol = ws.Cells(k, 9).Value
                
            End If
            
            'Conditional statement to locate greatest total volume
            If ws.Cells(k, 12).Value > greatest_total_volume Then
            
                greatest_total_volume = ws.Cells(k, 12).Value
                greatest_ticker_volume = ws.Cells(k, 9).Value
            
            End If
                
        Next k
        
        'Write values to summary output cells
        ws.Cells(2, 15).Value = greatest_ticker_symbol
        ws.Cells(2, 16).Value = greatest_increase
        
        ws.Cells(3, 15).Value = least_ticker_symbol
        ws.Cells(3, 16).Value = most_decrease
        
        ws.Cells(4, 15).Value = greatest_ticker_volume
        ws.Cells(4, 16).Value = greatest_total_volume
                
        'Format alignment and number format for values in cells in summary output columns
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Columns("N:P").EntireColumn.AutoFit
        ws.Columns("P").HorizontalAlignment = xlRight
        
        'Reset all values
        greatest_increase = 0
        most_decrease = 0
        greatest_total_volume = 0
        
    Next ws
         
End Sub