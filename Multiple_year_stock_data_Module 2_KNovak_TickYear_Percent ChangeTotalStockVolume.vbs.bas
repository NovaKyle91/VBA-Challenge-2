Attribute VB_Name = "Module1"
Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        
        Dim i As Long
        
        Dim j As Long
        
        Dim TickerCount As Long
        
        Dim LastRowA As Long
        
        Dim LastRowI As Long
        
        Dim PercentChange As Double
        
        Dim GreatestIncrease As Double
        
        Dim GreatestDecrease As Double
        
        Dim GreatestVolume As Double
        
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
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                    
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    ws.Cells(TickerCount, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                TickerCount = TickerCount + 1
                
                j = i + 1
                
                End If
            
            Next i

    Next ws
    
End Sub
