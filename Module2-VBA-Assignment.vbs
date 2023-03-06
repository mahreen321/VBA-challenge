Attribute VB_Name = "Module1"
Sub stockmarket()

    Dim ws As Worksheet
    
    
    For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total stock volume"
    
    ws.Range("P2").Value = "Ticker"
    ws.Range("Q2").Value = "Value"
    ws.Range("O3").Value = "Greatest % Increase"
    ws.Range("O4").Value = "Greatest % decrease"
    ws.Range("O5").Value = "Greatest total volume"
    
        TickerCount = 2
        j = 2
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            For i = 2 To LastRowA
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                If ws.Cells(TickerCount, 10).Value < 0 Then
                ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                
                    Else
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                
                    End If
                
                If ws.Cells(TickerCount, 11).Value < 0 Then
                ws.Cells(TickerCount, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(TickerCount, 11).Interior.ColorIndex = 4
            End If
        
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(TickerCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                

                TickerCount = TickerCount + 1
                
                j = i + 1
                
                End If
            Next i
            
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
            For i = 2 To LastRowI
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(5, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
            
            If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            ws.Cells(3, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(5, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
                    
            
    Next ws
        
End Sub

Sub sheets()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018")

End Sub


    

