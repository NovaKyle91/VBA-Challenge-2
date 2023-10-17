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
    
    Next ws
    
End Sub
