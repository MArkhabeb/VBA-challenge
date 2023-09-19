 Sub Multiple_Year_Stock_Data():
    For Each WS In Worksheets
    
            Dim Worksheet_name As String
             Dim i As Long
             Dim j As Long
             Dim StockTick As Long
             Dim LastRow As Long
             Dim LastRowI As Long
             Dim PercentChange As Double
             Dim GreatestIncrease As Double
             Dim GreatestDecrease As Double
             Dim GreatestVolume As Double
             
        Worksheet_name = WS.Name
                WS.Cells(1, 9).Value = "TICKER"
                WS.Cells(1, 10).Value = "Yearly Change"
                WS.Cells(1, 11).Value = "Percent Change"
                WS.Cells(1, 12).Value = "Total Stock Volume"
                WS.Cells(1, 16).Value = "TICKER"
                WS.Cells(1, 17).Value = "Value"
                WS.Cells(2, 15).Value = "Greatest % Increase"
                WS.Cells(3, 15).Value = "Greatest % Decrease"
                WS.Cells(4, 15).Value = "Greatest Total Volume"
            
        StockTick = 2
        
        j = 2
        
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
            
         For i = 2 To LastRow
              If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                 WS.Cells(StockTick, 9).Value = WS.Cells(i, 1).Value
                 WS.Cells(StockTick, 10).Value = WS.Cells(i, 6).Value - WS.Cells(j, 3).Value
              If WS.Cells(StockTick, 10).Value < 0 Then
                 WS.Cells(StockTick, 10).Interior.ColorIndex = 3
         Else
                 WS.Cells(StockTick, 10).Interior.ColorIndex = 4
                
         End If
                If WS.Cells(j, 3).Value <> 0 Then
                 PercentChange = ((WS.Cells(i, 6).Value - WS.Cells(j, 3).Value) / WS.Cells(j, 3).Value)
                 WS.Cells(StockTick, 11).Value = Format(PercentChange, "Percent")
         Else
                WS.Cells(StockTick, 11).Value = Format(0, "Percent")
                    
         End If
                WS.Cells(StockTick, 12).Value = WorksheetFunction.Sum(Range(WS.Cells(j, 7), WS.Cells(i, 7)))
                StockTick = StockTick + 1
                j = i + 1
        End If
        Next i
                    LastRowI = WS.Cells(Rows.Count, 9).End(xlUp).Row
                    GreatestVolume = WS.Cells(2, 12).Value
                    GreatestIncrease = WS.Cells(2, 11).Value
                    GreatestDecrease = WS.Cells(2, 11).Value
          For i = 2 To LastRowI
            
                    If WS.Cells(i, 12).Value > GreatestVolume Then
                    GreatestVolume = WS.Cells(i, 12).Value
                    WS.Cells(4, 16).Value = WS.Cells(i, 9).Value
                    
            Else
                    GreatestVolume = GreatestVolume
                
            End If
                    If WS.Cells(i, 11).Value > GreatestIncrease Then
                    GreatestIncrease = WS.Cells(i, 11).Value
                    WS.Cells(2, 16).Value = WS.Cells(i, 9).Value
                    
             Else
                    GreatestIncrease = GreatestIncrease
                
             End If
                    If WS.Cells(i, 11).Value < GreatestDecrease Then
                    GreatestDecrease = WS.Cells(i, 11).Value
                    WS.Cells(3, 16).Value = WS.Cells(i, 9).Value
                
             Else
                     GreatestDecrease = GreatestDecrease
             End If
                    WS.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
                    WS.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
                    WS.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
             Next i
        Worksheets(Worksheet_name).Columns("A:Z").AutoFit
           Next WS
           
           
End Sub


