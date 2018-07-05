Attribute VB_Name = "Module1"
Sub StockSolution()

    'Declaring worksheet variable
    Dim ws As Worksheet
    For Each ws In Worksheets
    
        Dim Ticker As String
        Dim TotalVol As Double
        Dim i As Long
        Dim TotalRowCount As Long
        Dim lastRow As Double
        Dim StockOpening, StockClosing As Double
        Dim YearlyChange, PercentChange As Double
        Dim greatestIncTicker As String
        Dim greatestIncValue As Double
        Dim greatestDecTicker As String
        Dim greatestDecValue As Double
        Dim greatestVolTicker As String
        Dim greatestVolValue As Double
        
        greatestIncValue = 0
        greatestDecValue = 0
        greatestVolValue = 0
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Labeling headers
        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("M1").Value = "Percent Change"
        ws.Range("N1").Value = "Total Stock Volume"
        
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        Ticker = ws.Cells(2, 1).Value
        StockOpening = ws.Cells(2, 3).Value
        TotalVol = 0
        TotalRowCount = 2
        
        For i = 2 To lastRow
            TotalVol = TotalVol + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                StockClosing = ws.Cells(i, 6).Value
                YearlyChange = StockClosing - StockOpening
                
                If StockOpening <> 0 Then
                    PercentChange = YearlyChange / StockOpening
                End If
                
                ws.Cells(TotalRowCount, 11).Value = Ticker
                
                If YearlyChange > 0 Then
                    ws.Cells(TotalRowCount, 12).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    ws.Cells(TotalRowCount, 12).Interior.ColorIndex = 3
                Else
                    ws.Cells(TotalRowCount, 12).Interior.ColorIndex = 0
                End If
                
                ws.Cells(TotalRowCount, 12).Value = YearlyChange
                ws.Cells(TotalRowCount, 14).Value = TotalVol
                
                If TotalVol > greatestVolValue Then
                    greatestVolValue = TotalVol
                    greatestVolTicker = Ticker
                End If
                
                If StockOpening <> 0 Then
                    ws.Cells(TotalRowCount, 13).Value = PercentChange
                Else
                    ws.Cells(TotalRowCount, 13).Value = "N/A"
                End If
                
                If PercentChange > greatestIncValue Then
                    greatestIncValue = PercentChange
                    greatestIncTicker = Ticker
                ElseIf PercentChange < greatestDecValue Then
                    greatestDecValue = PercentChange
                    greatestDecTicker = Ticker
                End If
                
                ws.Cells(TotalRowCount, 13).NumberFormat = "0.00%"
                Ticker = ws.Cells(i + 1, 1).Value
                StockOpening = ws.Cells(i + 1, 3).Value
                TotalVol = 0
                TotalRowCount = TotalRowCount + 1
        
            End If
        Next i
        ws.Cells(2, 17).Value = greatestIncTicker
        ws.Cells(2, 18).Value = greatestIncValue
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = greatestDecTicker
        ws.Cells(3, 18).Value = greatestDecValue
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = greatestVolTicker
        ws.Cells(4, 18).Value = greatestVolValue
    Next ws
End Sub
