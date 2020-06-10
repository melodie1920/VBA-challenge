Sub Greatest()

    For Each ws In Worksheets

        Dim TickerIncrease As String
        Dim ChangeIncrease As Double
        Dim TickerDecrease As String
        Dim ChangeDecrease As Double
        Dim TickerTotalVolume As String
        Dim TotalVolume As Double
        
        ChangeIncrease = 0
        ChangeDecrease = 0
        TotalVolume = 0
        LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ws.Cells(2, LastColumn + 3) = "Greatest % Increase"
        ws.Cells(3, LastColumn + 3) = "Greatest % Decrease"
        ws.Cells(4, LastColumn + 3) = "Greatest Total Volume"
        ws.Cells(1, LastColumn + 4) = "Ticker"
        ws.Cells(1, LastColumn + 5) = "Value"
        
        ws.Columns(LastColumn + 3).AutoFit
        
        For i = 2 To LastRow
            
            If ws.Cells(i, 11).Value > ChangeIncrease Then
                TickerIncrease = ws.Cells(i, 9).Value
                ChangeIncrease = ws.Cells(i, 11).Value
            End If
        
            If ws.Cells(i, 11).Value < ChangeDecrease Then
                TickerDecrease = ws.Cells(i, 9).Value
                ChangeDecrease = ws.Cells(i, 11).Value
            End If
       
            If ws.Cells(i, 12).Value > TotalVolume Then
                TickerTotalVolume = ws.Cells(i, 9).Value
                TotalVolume = ws.Cells(i, 12).Value
            End If
       
        Next i
      
    ws.Cells(2, LastColumn + 4) = TickerIncrease
    ws.Cells(2, LastColumn + 5) = ChangeIncrease
    ws.Cells(3, LastColumn + 4) = TickerDecrease
    ws.Cells(3, LastColumn + 5) = ChangeDecrease
    ws.Cells(4, LastColumn + 4) = TickerTotalVolume
    ws.Cells(4, LastColumn + 5) = TotalVolume
    ws.Cells(2, LastColumn + 5).NumberFormat = "0.00%"
    ws.Cells(3, LastColumn + 5).NumberFormat = "0.00%"
    
    Next ws

End Sub
