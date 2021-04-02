Sub stonks()

    For Each ws In Worksheets
        Dim counter As Integer
        
        Dim YearOpen As Double
        Dim YearClose As Double
        Dim TotalVolume As Double
        
        Dim PercentIncreaseTicker As String
        Dim LargestPercentIncrease As Double
        Dim LargestPercentDecrease As Double
        Dim PercentDecreaseTicker As String
        Dim LargestTotalVolume As Double
        Dim LargestTotalVolumeTicker As String
        
        YearOpen = ws.Cells(2, 3).Value
        
        counter = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        WorksheetName = ws.Name
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        For i = 2 To lastRow
        
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
          
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'grab ticker symbol
                ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
                'set year close
                YearClose = ws.Cells(i, 6).Value
                'calc change
                ws.Cells(counter, 10).Value = YearClose - YearOpen
                'calc percent change and then format
                If YearOpen <> 0 Then
                
                    ws.Cells(counter, 11).Value = (YearClose - YearOpen) / YearOpen
                    
                End If
                
                ws.Cells(counter, 11).NumberFormat = "%0.00"
                
               
                
                
                'set new year open
                YearOpen = ws.Cells(i + 1, 3).Value
                ws.Cells(counter, 12).Value = TotalVolume
                
                'format the yearly change column to inclue color coding
                
                If ws.Cells(counter, 10).Value > 0 Then
             
                    ws.Cells(counter, 10).Interior.ColorIndex = 4
                
                Else
            
                    ws.Cells(counter, 10).Interior.ColorIndex = 3
                
                End If
                
                
                'reset total volume so that it wont include the values from the last ticker
                
                TotalVolume = 0
                
                counter = counter + 1
                
                
                
            End If
            
            
            
        Next i
       
        
    Next ws
    
    'calls the bonus section to produce the greatest increase, decrease, and volume
    
    Call largest
   

End Sub
