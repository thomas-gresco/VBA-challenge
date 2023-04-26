Sub Stocks():

    For Each ws In Worksheets
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim Tick As Long
        Dim LastRowA As Long
        Dim PC As Double
        Dim LastRowI As Long
        Dim GV As Double
        Dim GInc As Double
        Dim GDec As Double
        
        WorksheetName = ws.Name
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        Tick = 2
        j = 2
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRowA
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(Tick, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Tick, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
            If ws.Cells(Tick, 10).Value < 0 Then
                ws.Cells(Tick, 10).Interior.ColorIndex = 3
                
                Else
                    ws.Cells(Tick, 10).Interior.ColorIndex = 4
                
                End If
                    
            If ws.Cells(j, 3).Value <> 0 Then
                PC = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(Tick, 11).Value = Format(PC, "Percent")
                Else
                    ws.Cells(Tick, 11).Value = Format(0, "Percent")
                    
                End If
                    
            ws.Cells(Tick, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                Tick = Tick + 1
                j = i + 1
                
                End If
            
            Next i
            
        'Part 2
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        GV = ws.Cells(2, 12).Value
        GInc = ws.Cells(2, 11).Value
        GDec = ws.Cells(2, 11).Value
        
        For i = 2 To LastRowI
            If ws.Cells(i, 11).Value > GInc Then
            GInc = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
            Else
                
                GInc = GInc
                
            End If
            
            If ws.Cells(i, 11).Value < GDec Then
            GDec = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
            Else
                
                GDec = GDec
            
            End If
            
            If ws.Cells(i, 12).Value > GV Then
            GV = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
            Else
                
                GV = GV
                
            End If
                
            ws.Cells(2, 17).Value = Format(GInc, "Percent")
            ws.Cells(3, 17).Value = Format(GDec, "Percent")
            ws.Cells(4, 17).Value = Format(GV, "Scientific")
            
            Next i
            
            
    Next ws
        
End Sub
