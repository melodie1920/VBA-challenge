Sub Stocks()

    For Each ws In Worksheets

        Dim Ticker As Integer
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim FirstDate As Double
        Dim YearlyChange As Double
        Dim StockVolume As Double

        Ticker = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        FirstDate = 99999999
        StockVolume = 0
        
        ws.Cells(1, LastColumn + 2) = "Ticker"
        ws.Cells(1, LastColumn + 3) = "Yearly Change"
        ws.Cells(1, LastColumn + 4) = "Percent Change"
        ws.Cells(1, LastColumn + 5) = "Total Stock Volume"
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(Ticker, LastColumn + 2).Value = ws.Cells(i, 1).Value
                ClosingPrice = ws.Cells(i, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                ws.Cells(Ticker, LastColumn + 3).Value = YearlyChange
                
                    If OpeningPrice = 0 Then
                       ws.Cells(Ticker, LastColumn + 4).Value = 0
                    Else
                        ws.Cells(Ticker, LastColumn + 4).Value = YearlyChange / OpeningPrice
                    End If
                    
                ws.Cells(Ticker, LastColumn + 4).NumberFormat = "0.00%"
                    
                    If YearlyChange > 0 Then
                        ws.Cells(Ticker, LastColumn + 3).Interior.ColorIndex = 4 'Green
                    Else
                        ws.Cells(Ticker, LastColumn + 3).Interior.ColorIndex = 3 'Red
                    End If
                   
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                ws.Cells(Ticker, LastColumn + 5).Value = StockVolume
                   
                Ticker = Ticker + 1
                FirstDate = 99999999
                StockVolume = 0
            Else
                
                If ws.Cells(i, 2).Value < FirstDate Then
                    OpeningPrice = ws.Cells(i, 3).Value
                    FirstDate = ws.Cells(i, 2).Value
                End If
                
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
    ws.Columns(LastColumn + 5).AutoFit
    
    Next ws
           
End Sub
