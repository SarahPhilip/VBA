Sub StockSolution()

    'Declaring worksheet variable
    Dim ws As Worksheet
    For Each ws In Worksheets
    
        'Declaring all variables
        Dim ticker As String
        Dim total_vol As Double
        Dim r As Long
        Dim output_row As Long
        Dim last_row As Double
        Dim stock_open, stock_close As Double
        Dim y_change, p_change As Double
        Dim greatest_inc_ticker As String
        Dim greatest_inc_value As Double
        Dim greatest_dec_ticker As String
        Dim greatest_dec_value As Double
        Dim greatest_vol_ticker As String
        Dim greatest_vol_value As Double
        
        'Initializing dummy variables to 0
        greatest_inc_value = 0
        greatest_dec_value = 0
        greatest_vol_value = 0
        total_vol = 0
        
        'Initalizing the first row of the output columns
        output_row = 2
        
        'Getting the last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Automatically determines the column width
        ws.Cells.EntireColumn.AutoFit
        
        'Labeling Output Headers
        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("M1").Value = "Percent Change"
        ws.Range("N1").Value = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
    
        'Looping through the rows
        For r = 2 To last_row
            If (r = 2) Then
                'Getting the value from excel file
                ticker = ws.Cells(r, 1).Value
                stock_open = ws.Cells(r, 3).Value
            End If
            'Adding up the volume of the stock
            total_vol = total_vol + ws.Cells(r, 7).Value
            
            'Checking if the value in the next row has the same ticker symbol. If not, find closing value and calculate yearly change and percent change
            If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1) Then
                stock_close = ws.Cells(r, 6).Value
                y_change = stock_close - stock_open
                
                'If stock opening is zero, we get a zero error. So check if stock_open is zero.
                If stock_open <> 0 Then
                    p_change = y_change / stock_open
                End If
                
                'Writing the values in output column
                ws.Cells(output_row, 11).Value = ticker
                ws.Cells(output_row, 12).Value = y_change
                ws.Cells(output_row, 14).Value = total_vol
                 If stock_open <> 0 Then
                    ws.Cells(output_row, 13).Value = p_change
                    ws.Cells(output_row, 13).NumberFormat = "0.00%"
                Else
                    ws.Cells(output_row, 13).Value = "N/A"
                End If
                
                'Formatting for the positive change in green and negative change in red.
                If y_change > 0 Then
                    ws.Cells(output_row, 12).Interior.ColorIndex = 4
                ElseIf y_change < 0 Then
                    ws.Cells(output_row, 12).Interior.ColorIndex = 3
                Else
                    ws.Cells(output_row, 12).Interior.ColorIndex = 0
                End If
                           
                'Finding Greatest % increase, Greatest % Decrease and Greatest total volume
                If total_vol > greatest_vol_value Then
                    greatest_vol_value = total_vol
                    greatest_vol_ticker = ticker
                End If

                If p_change > greatest_inc_value Then
                    greatest_inc_value = p_change
                    greatest_inc_ticker = ticker
                ElseIf p_change < greatest_dec_value Then
                    greatest_dec_value = p_change
                    greatest_dec_ticker = ticker
                End If
                
                'Initializing with new values
                ticker = ws.Cells(r + 1, 1).Value
                stock_open = ws.Cells(r + 1, 3).Value
                total_vol = 0
                output_row = output_row + 1
        
            End If
        Next r
        'Writing the outsput values to the table
        ws.Cells(2, 17).Value = greatest_inc_ticker
        ws.Cells(2, 18).Value = greatest_inc_value
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = greatest_dec_ticker
        ws.Cells(3, 18).Value = greatest_dec_value
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = greatest_vol_ticker
        ws.Cells(4, 18).Value = greatest_vol_value
    Next ws
End Sub
