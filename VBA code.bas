Attribute VB_Name = "Module1"
Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        'assigning variables
        
        Dim WorksheetName As String
        
        Dim x As Long
        Dim y As Long
    
        Dim LastRowA As Long
        Dim LastRowI As Long
        
        Dim PerChange As Double
        
        Dim GreatIncrease As Double 'greatest increase
        Dim GreatDecrease As Double 'greatest decrease
        Dim GreatVolume As Double   'greatest volume
        
        Dim TickerCount As Long
        
        WorksheetName = ws.Name
        
        'column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        

        TickerCount = 2 'row 2
        
        y = 2 'start row is 2
        
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
            'Loop through all rows
            For x = 2 To LastRowA
            
                If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
                
                ws.Cells(TickerCount, 9).Value = ws.Cells(x, 1).Value
                
                ws.Cells(TickerCount, 10).Value = ws.Cells(x, 6).Value - ws.Cells(y, 3).Value
                
                    'Conditional formating
                    If ws.Cells(TickerCount, 10).Value < 0 Then
                
                    'background color red
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'background color green
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(y, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(x, 6).Value - ws.Cells(y, 3).Value) / ws.Cells(y, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TickerCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'write total volume in column L
                ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(y, 7), ws.Cells(x, 7)))
                
                TickerCount = TickerCount + 1 'Increase TickerCount by 1
                
                y = x + 1
                
                End If
            
            Next x
            
      
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row   'last cell of colume I that is not blank
        
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        
            'Loop for summary
            For x = 2 To LastRowI
        
                If ws.Cells(x, 12).Value > GreatVolume Then 'finding the greatest volume
                
                GreatVolume = ws.Cells(x, 12).Value
                
                ws.Cells(4, 16).Value = ws.Cells(x, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
                
                End If
                
                If ws.Cells(x, 11).Value > GreatIncrease Then  'finding the greatest increase
                GreatIncrease = ws.Cells(x, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(x, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                
                End If
                
            
                If ws.Cells(x, 11).Value < GreatDecrease Then  'finding the greatest decrease
                GreatDecrease = ws.Cells(x, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(x, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
            Next x
            
            
    Next ws
        
End Sub
