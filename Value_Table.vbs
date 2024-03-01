Sub Value_Table()
    
    For Each ws In ThisWorkbook.Sheets
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
            Dim i As Long
            Dim greatestIncrease As Double
            Dim greatestDecrease As Double
            Dim greatestVolume As Double
            Dim tickerGreatestIncrease As String
            Dim tickerGreatestDecrease As String
            Dim tickerGreatestVolume As String
            
                ws.Cells(1, "P").Value = "Ticker"
                ws.Cells(1, "Q").Value = "Value"
    
                greatestIncrease = 0
                greatestDecrease = 1
                greatestVolume = 0

                    For i = 2 To lastRow
                        If ws.Cells(i, "K").Value > greatestIncrease Then
                            greatestIncrease = ws.Cells(i, "K").Value
                            tickerGreatestIncrease = ws.Cells(i, "I").Value
                        End If
                        
                        If ws.Cells(i, "K").Value < greatestDecrease Then
                            greatestDecrease = ws.Cells(i, "K").Value
                            tickerGreatestDecrease = ws.Cells(i, "I").Value
                        End If
                        
                        If ws.Cells(i, "L").Value > greatestVolume Then
                            greatestVolume = ws.Cells(i, "L").Value
                            tickerGreatestVolume = ws.Cells(i, "I").Value
                        End If
                    Next i
    
                With ws
                    .Cells(2, "O").Value = "Greatest % Increase"
                    .Cells(2, "P").Value = tickerGreatestIncrease
                    .Cells(2, "Q").Value = greatestIncrease
                    
                    .Cells(3, "O").Value = "Greatest % Decrease"
                    .Cells(3, "P").Value = tickerGreatestDecrease
                    .Cells(3, "Q").Value = greatestDecrease
                    
                    .Cells(4, "O").Value = "Greatest Total Volume"
                    .Cells(4, "P").Value = tickerGreatestVolume
                    .Cells(4, "Q").Value = greatestVolume
                End With
                
            Next ws
End Sub