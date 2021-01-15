Sub Stock_Market_Analyst():

    'Declare variables and set initial values
    Dim ws As Worksheet
    
    Dim ticker_symbol As String
    Dim i As Long
    Dim j As Long
    
    Dim year_open As Double
    year_open = 0
    
    Dim year_close As Double
    year_close = 0
    
    Dim yearly_change As Double
    yearly_change = 0
    
    Dim percent_change As Double
    percent_change = 0
    
    Dim total_stock_volume As Long
    total_stock_volume = 0
    
    Dim output_row As Long
    
    'Workaround code for overflow error
    On Error Resume Next
    
    'Loop through all sheets in workbook
    For Each ws In Worksheets
    
        'Write text to output column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Declare variable to write to appropriate row in output columns
        output_row = 2
        j = 0
    
        'Declare variable to find the last row of data
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Loop through all rows of data in all sheets in workbook
        For i = 2 To last_row
    
            'If ticker changes, then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Store results in variables
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            'Handle zero total volume
            If total_stock_volume = 0 Then
                
                'Write values to cells
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
                
            Else
            
                'Find first non zero starting value to calculate percentage change
                If ws.Cells(output_row, 3) = 0 Then
                    
                    For find_value = output_row To i
                        
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            
                            output_row = find_value
                            
                            Exit For
                        
                        End If
                     
                     Next find_value
                
                End If
                
                'Retrieve values and calculate percentage change
                yearly_change = ws.Cells(i, 6).Value - ws.Cells(output_row, 3).Value
                percent_change = Round((yearly_change / ws.Cells(output_row, 3).Value * 100), 2)
                
                'Start of the next stock ticker
                output_row = i + 1
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = Round(yearly_change, 2)
                ws.Range("K" & 2 + j).Value = "%" & percent_change
                ws.Range("L" & 2 + j).Value = total_stock_volume
                
                'Apply conditional formatting to output column
                Select Case yearly_change
                
                    Case Is > 0
                    
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                         
                    Case Is < 0
                    
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            
                    Case Else
                    
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
            
                End Select
            
            End If
            
            'Reset variables for a new stock ticker
            total_stock_volume = 0
            yearly_change = 0
            j = j + 1
            
            Else
        
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            End If
    
        Next i
    
    Next ws

End Sub