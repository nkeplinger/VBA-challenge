Attribute VB_Name = "Module3"
Sub testing()
    'Add column variable as integer for counting
    Dim column As Double
    
    'loop through each worksheet
    For Each ws In Worksheets
        'define columns for placing new data in each sheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly_Change"
        ws.Cells(1, 11).Value = "Percent_Change"
        ws.Cells(1, 12).Value = "Total_Stock_Volume"
        'identify the last row in col A each worksheet
        LastRowInColA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'counter variable for data rows
        Dim data_rows As Long
        data_row = 0
        
        'loop through second to last row
        For i = 2 To LastRowInColA
            'Find cells when value of next cell is different
            'define start_price for open and stock_sum for adding
            Dim start_price As Double
            start_price = ws.Cells(i, 3).Value
            'Dim stock_sum As Long
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'add name of ticker to column I/9
                data_row = data_row + 1
                ticker_name = ws.Cells(i, 1).Value
                ws.Cells(data_row + 1, 9).Value = ticker_name
                
                'identify close price
                Dim end_price As Double
                end_price = ws.Cells(i, 6).Value
                
                'find yearly change and add to column J/10
                yrchange = end_price - start_price
                ws.Cells(data_row + 1, 10).Value = yrchange
                
                'Find percent change and add to column k/11
                Percent_Change = 100 * ((end_price - start_price) / start_price)
                ws.Cells(data_row + 1, 11).Value = Percent_Change & "%"
                
            End If
            'add values for sum of values of each ticker/group of tickers
            If ws.Cells(i, 1).Value = ws.Cells(data_row + 1, 9).Value Then
                stock_sum = 0
                stock_sum = stock_sum + ws.Cells(i, 7).Value
                ws.Cells(data_row + 1, 12).Value = stock_sum
            End If

        Next i
        
    Next ws
End Sub
