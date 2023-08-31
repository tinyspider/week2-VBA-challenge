Sub loop1()
For Each ws In ThisWorkbook.Worksheets:
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("j2").Value = lastrow
    ticker = "test"
    Dim total As Integer
    total = "0"
    Dim j As Integer
    j = 2
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    
    
    For i = 2 To lastrow
        If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
            price_begin = ws.Cells(i, 3)
        End If
        If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
            vol_total = vol_total + ws.Cells(i, 7)
        Else
            price_end = ws.Cells(i, 6).Value
            vol_total = vol_total + ws.Cells(i, 7)
            ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(j, 12).Value = vol_total
            ws.Cells(j, 10).Value = price_end - price_begin
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
                ws.Cells(j, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
                ws.Cells(j, 11).Interior.ColorIndex = 3
            End If
            ws.Cells(j, 11).Value = (price_end - price_begin) / price_begin

            
            vol_total = 0
            j = j + 1
        End If
    Next i
    
    last_row = j - 1
    ws.Range("j" & 2 & ":j" & last_row).NumberFormat = "0.00"
    ws.Range("k" & 2 & ":k" & last_row).NumberFormat = "0.00%"
    ws.Columns("L").ColumnWidth = 15

ws.Range("o2") = "Greatest % increase"
ws.Range("o3") = "Greatest % decrease"
ws.Range("o4") = "Greatest total volume"
ws.Range("p1") = "Ticker"
ws.Range("q1") = "Value"



index_gi = 2
index_gd = 2
index_gtv = 2
Greatest_i = ws.Cells(2, 11).Value
Greatest_d = ws.Cells(2, 11).Value
Greatest_tv = ws.Cells(2, 12).Value


For k = 3 To last_row
    If Greatest_i < ws.Cells(k, 11).Value Then
        Greatest_i = ws.Cells(k, 11).Value
        index_gi = k
    End If
    If Greatest_d > ws.Cells(k, 11).Value Then
        Greatest_d = ws.Cells(k, 11).Value
        index_gd = k
    End If
    If Greatest_tv < ws.Cells(k, 12).Value Then
        Greatest_tv = ws.Cells(k, 12)
        index_gtv = k
    End If
Next k
ws.Range("p2").Value = ws.Cells(index_gi, 9).Value
ws.Range("p3").Value = ws.Cells(index_gd, 9).Value
ws.Range("p4").Value = ws.Cells(index_gtv, 9).Value
ws.Range("q2").Value = ws.Cells(index_gi, 11).Value
ws.Range("q3").Value = ws.Cells(index_gd, 11).Value
ws.Range("q4").Value = ws.Cells(index_gtv, 12).Value
ws.Range("q2:q3").NumberFormat = "0.00%"
ws.Columns("q").ColumnWidth = 15



Next ws
End Sub
