Sub stock()

'Loop Through every worksheet

For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volumn"
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    start_row = 2
    year_count = 2
    total_stock_vol = 0
  
'Loop Through Ticker Column
'Check if ticker value are the same or not. If not, then print in ticker column
'Calculate the Yearly Change by each ticker
'Calculate the Percentage Change by each ticker
'Calculate the total stock volumn
    
For i = 2 To lastRow

    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim year_change As Double
    
    year_open = 0
    year_close = 0
    ticker = ws.Cells(i, 1).Value
    year_open = ws.Cells(year_count, 3).Value
    total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
    
    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
    ws.Cells(start_row, 9).Value = ticker
    year_close = ws.Cells(i, 6).Value
    year_change = year_close - year_open
    ws.Cells(start_row, 10).Value = year_change
    ws.Cells(start_row, 12).Value = total_stock_vol
    
'make sure year_open is not "0", since we cannot divide any number by 0

    If year_open <> 0 Then
        percent_change = (year_close - year_open) / year_open
        ws.Cells(start_row, 11).Value = percent_change
    Else
        ws.Cells(start_row, 11).Value = "NA"
        
    End If
    
'Star_row and year_count are placeholder to keep track of the value of ticker and year
    start_row = start_row + 1
    year_count = i + 1
    total_stock_vol = 0       'Used to reset the stock volumn count when ticker updated
    
    End If
    
    
Next i

'Conditional Formatting

last_J = ws.Cells(Rows.Count, 10).End(xlUp).Row
For r = 2 To last_J

    If ws.Cells(r, 10).Value < 0 Then
        ws.Cells(r, 10).Interior.ColorIndex = 3
    Else
     ws.Cells(r, 10).Interior.ColorIndex = 4
    End If
Next r
    
'Formatting Column "K" into two decimal place and add a "%" sign

last_K = ws.Cells(Rows.Count, 11).End(xlUp).Row
For k = 2 To last_K
    ws.Cells(k, 11).NumberFormat = "0.00%"
Next k
 
'Bonus
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volumn"

For b = 2 To last_J
    If ws.Cells(b, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & last_J)) Then
    g_incr_ticker = ws.Cells(b, 9).Value
    g_incr = ws.Cells(b, 11).Value
    ws.Range("P2").Value = g_incr_ticker
    ws.Range("Q2").Value = g_incr
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ElseIf ws.Cells(b, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & last_J)) Then
    g_decr_ticker = ws.Cells(b, 9).Value
    g_decr = ws.Cells(b, 11).Value
    ws.Range("P3").Value = g_decr_ticker
    ws.Range("Q3").Value = g_decr
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ElseIf ws.Cells(b, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & last_J)) Then
    g_tot_ticker = ws.Cells(b, 9).Value
    g_tot = ws.Cells(b, 12).Value
    ws.Range("P4").Value = g_tot_ticker
    ws.Range("Q4").Value = g_tot
    End If
    


Next b

 
Next ws


End Sub


