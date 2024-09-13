Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through rows
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 2).Value
            closePrice = ws.Cells(i, 3).Value
            volume = ws.Cells(i, 4).Value
            
            ' Quarterly change
            quarterlyChange = closePrice - openPrice
            ' percentage change
            percentChange = (closePrice - openPrice) / openPrice * 100
            
            ' max values update
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                maxIncreaseTicker = ticker
            ElseIf percentChange < maxDecrease Then
                maxDecrease = percentChange
                maxDecreaseTicker = ticker
            End If
            
            If volume > maxVolume Then
                maxVolume = volume
                maxVolumeTicker = ticker
            End If
            
            ' Output/store results
            ' conditional formatting here**ummmno
            
        Next i
    Next ws
    
    ' Output  stocks with the greatest % increase, % decrease, and total volume
    ' display results as needed
End Sub

