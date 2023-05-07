Sub StockData()

For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

    Dim Ticker As String
    
    Dim YearlyChange As Double
    YearlyChange = 0
        
    Dim PercentChange As Double
    PercentChange = 0
    
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
    'Lindsay through AskBCS helped me with this open value-close value part
    'I was stuck on referencing first cell and last cell where she explained I don't need to do that
    'since the loop will "redefine" those values
    Dim FirstOpen As Double
    FirstOpen = ws.Cells(2, 3).Value
    Dim LastClose As Double
    LastClose = 0
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    Dim SummaryTable As Double
    SummaryTable = 2
    
        For i = 2 To LastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker = ws.Cells(i, 1).Value
                
                LastClose = ws.Cells(i, 6).Value
                
                YearlyChange = LastClose - FirstOpen
                
                PercentChange = YearlyChange / FirstOpen
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                ws.Range("I" & SummaryTable).Value = Ticker
                ws.Range("J" & SummaryTable).Value = YearlyChange
                ws.Range("K" & SummaryTable).Value = FormatPercent(PercentChange)
                ws.Range("L" & SummaryTable).Value = TotalStockVolume
                
                SummaryTable = SummaryTable + 1
               
                FirstOpen = ws.Cells(i + 1, 3).Value
                
                YearlyChange = 0
        
                PercentChange = 0
    
                TotalStockVolume = 0
                
            Else
                    
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            End If
        
        Next i

LastRowTwo = ws.Cells(Rows.Count, 10).End(xlUp).Row

    For i = 2 To LastRowTwo
    
        If ws.Cells(i, 10).Value < 0 Then
            
            ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
            ws.Cells(i, 10).Interior.ColorIndex = 4
        
        End If
    
    Next i


ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotal As Double
    
    Dim IncreaseTicker As String
    Dim DecreaseTicker As String
    Dim TotalTicker As String

    GreatestIncrease = WorksheetFunction.Max(ws.Range("K:K"))
    IncreaseTicker = WorksheetFunction.Match(GreatestIncrease, ws.Range("K:K"), 0)
    ws.Range("Q2").Value = FormatPercent(GreatestIncrease)
    ws.Range("P2").Value = ws.Cells(IncreaseTicker, "I").Value

    
    GreatestDecrease = WorksheetFunction.Min(ws.Range("K:K"))
    DecreaseTicker = WorksheetFunction.Match(GreatestDecrease, ws.Range("K:K"), 0)
    ws.Range("Q3").Value = FormatPercent(GreatestDecrease)
    ws.Range("P3").Value = ws.Cells(DecreaseTicker, "I").Value
    
    GreatestTotal = WorksheetFunction.Max(ws.Range("L:L"))
    TotalTicker = WorksheetFunction.Match(GreatestTotal, ws.Range("L:L"), 0)
    ws.Range("Q4").Value = GreatestTotal
    ws.Range("P4").Value = ws.Cells(TotalTicker, "I").Value

ws.Columns("A:Q").AutoFit

Next ws

End Sub

