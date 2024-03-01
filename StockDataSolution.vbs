Sub StockData()
    
    For Each ws In ThisWorkbook.Sheets
   
    
            Dim lastRow As Long
            Dim i As Long
            Dim tickerSymbol As String
            Dim startPrice As Double
            Dim endPrice As Double
            Dim totalVolume As Double
            Dim outputRow As Long
            
                    ws.Cells(1, "I").Value = "Ticker"
                    ws.Cells(1, "J").Value = "Yearly Change"
                    ws.Cells(1, "K").Value = "Percentage Change"
                    ws.Cells(1, "L").Value = "Total Stock Volume"
    
                lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
                outputRow = 2
    
                    totalVolume = 0
    
                    If IsNumeric(ws.Cells(2, "C").Value) Then
                        startPrice = ws.Cells(2, "C").Value
                    Else
                        startPrice = 0
                    End If
                    
                        
                        For i = 2 To lastRow
                            
                            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                
                                tickerSymbol = ws.Cells(i, 1).Value
                                
                                
                                If IsNumeric(ws.Cells(i, "F").Value) Then
                                    endPrice = ws.Cells(i, "F").Value
                                Else
                                    endPrice = 0
                                End If
                                
                                If IsNumeric(ws.Cells(i, "G").Value) Then
                                    totalVolume = totalVolume + ws.Cells(i, "G").Value
                            End If
            
                                        Dim yearlyChange As Double
                                        Dim percentChange As Double
                                        
                                        yearlyChange = endPrice - startPrice
                                        
                                        If startPrice <> 0 And yearlyChange <> 0 Then
                                            percentChange = yearlyChange / startPrice
                                        Else
                                            percentChange = 0
                                        End If
            
                            With ws
                                .Cells(outputRow, "I").Value = tickerSymbol
                                .Cells(outputRow, "J").Value = yearlyChange
                                .Cells(outputRow, "K").Value = percentChange
                                .Cells(outputRow, "L").Value = totalVolume
                            End With
            
                                    outputRow = outputRow + 1
                                    
                                    totalVolume = 0
                                    If i + 1 <= lastRow And IsNumeric(ws.Cells(i + 1, "C").Value) Then
                                        startPrice = ws.Cells(i + 1, "C").Value
                                    Else
                                        startPrice = 0
                                    End If
                                
                                Else
                                
                                    If IsNumeric(ws.Cells(i, "G").Value) Then
                                    totalVolume = totalVolume + ws.Cells(i, "G").Value
                                End If
                            End If
                        Next i
                    Next ws
    
End Sub
