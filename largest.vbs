Sub largest()
    Dim PercentIncreaseTicker As String
    Dim LargestPercentIncrease As Double
    Dim LargestPercentDecrease As Double
    Dim PercentDecreaseTicker As String
    Dim LargestTotalVolume As Double
    Dim LargestTotalVolumeTicker As String
    
    
    lastRow = Range("J" & Rows.Count).End(xlUp).Row
    
    For Each ws In Worksheets
        
        LargestPercentIncrease = 0
        LargestPercentDecrease = 0
        LargestTotalVolume = 0
        
        For i = 2 To lastRow
            
            If LargestPercentIncrease < ws.Cells(i, 11).Value Then
                LargestPercentIncrease = ws.Cells(i, 11).Value
                PercentIncreaseTicker = ws.Cells(i, 9).Value
            End If
            If LargestPercentDecrease > ws.Cells(i, 11).Value Then
                LargestPercentDecrease = ws.Cells(i, 11).Value
                PercentDecreaseTicker = ws.Cells(i, 9).Value
            End If
            If LargestTotalVolume < ws.Cells(i, 12).Value Then
                LargestTotalVolume = ws.Cells(i, 12).Value
                LargestTotalVolumeTicker = ws.Cells(i, 9).Value
            End If
        Next i
        
        ws.Cells(2, 15).Value = PercentIncreaseTicker
        ws.Cells(2, 16).Value = LargestPercentIncrease
        ws.Cells(3, 15).Value = PercentDecreaseTicker
        ws.Cells(3, 16).Value = LargestPercentDecrease
        ws.Cells(4, 15).Value = LargestTotalVolumeTicker
        ws.Cells(4, 16).Value = LargestTotalVolume
    
        ws.Cells(2, 16).NumberFormat = "%0.00"
        ws.Cells(3, 16).NumberFormat = "%0.00"
    
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
    Next ws
    
    
    
End Sub
