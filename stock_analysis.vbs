Attribute VB_Name = "Module1"
Sub stock_analysis()

    'loop through each woorkseet
    For Each ws In Worksheets
            
            
          'Stock Volume variable
          Dim StockVol As Double
    
          StockVol = 0
            
          'last row of column
          lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            
          'first row of data
          Row = 2
          
          'first row of new ticker
          firstRow = 2
           
           
          'headers for table 1
          ws.Cells(1, 9).Value = "Ticker"
          ws.Cells(1, 10).Value = "Yearly Change"
          ws.Cells(1, 11).Value = "Percent Change"
          ws.Cells(1, 12).Value = "Total Stock Volume"
          
          
          'headers for table 2
          ws.Cells(1, 16).Value = "Ticker"
          ws.Cells(1, 17).Value = "Value"
          ws.Cells(2, 15).Value = "Greatest % Increase"
          ws.Cells(3, 15).Value = "Greatest % Decrease"
          ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
          'cycle through the raw data
          For i = 2 To lastRow
            
            'Sum Stock Volume until ticker value changes
            StockVol = StockVol + ws.Cells(i, 7).Value
    
            'if ticker value changes, then
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                'Collect ticker symbol
                ws.Cells(Row, 9).Value = ws.Cells(i, 1).Value
    
                'calculate and write yearly change
                ws.Cells(Row, 10).Value = ws.Cells(i, 6).Value - ws.Cells(firstRow, 3).Value
    
                'Colour in yearly change
                If ws.Cells(Row, 10).Value < 0 Then
    
                    ws.Cells(Row, 10).Interior.ColorIndex = 3
    
                Else
    
                    ws.Cells(Row, 10).Interior.ColorIndex = 4
    
                End If
                
    
                
                'calculate and write percent change
                ws.Cells(Row, 11).Value = ws.Cells(Row, 10).Value / ws.Cells(firstRow, 3).Value
                
                
                'Colour in percent change
                
                If ws.Cells(Row, 11).Value < 0 Then
                    
                    ws.Cells(Row, 11).Interior.ColorIndex = 3
                    
                Else
                
                    ws.Cells(Row, 11).Interior.ColorIndex = 4
                    
                End If
                
                
                'convert decimal value to percentage format
                
                ws.Cells(Row, 11).NumberFormat = "0.00%"
                
                'write Total Stock Volume
                
                ws.Cells(Row, 12).Value = StockVol
                
    
                
                'Update the row counter for table 1
                Row = Row + 1
                
                'Update the firstRow counter for new ticker
                firstRow = i + 1
                
                'Reset stock volume for the new ticker
                StockVol = 0
                
                
                
            
            End If
            
          
          Next i
          
          
          'find last row of Table 1 to produce Table 2
          lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
          
          'initialise MaxPercent, MinPercent and MaxVolume
          MaxPercent = ws.Cells(i, 11).Value
          MinPercent = ws.Cells(i, 11).Value
          MaxVolume = ws.Cells(i, 12).Value
          
          'cycle through Table 1
          For i = 2 To lastRow
          
            'compare and update variables
            If ws.Cells(i, 11).Value > MaxPercent Then
                MaxPercent = ws.Cells(i, 11).Value
            End If
            
            If ws.Cells(i, 11).Value < MinPercent Then
                MinPercent = ws.Cells(i, 11).Value
            End If
            
            If ws.Cells(i, 12).Value > MaxVolume Then
                MaxVolume = ws.Cells(i, 12).Value
            End If
          
          
          
          Next i
          
          'store the variables in table 2
          ws.Cells(2, 17).Value = MaxPercent
          ws.Cells(3, 17).Value = MinPercent
          ws.Cells(4, 17).Value = MaxVolume
            
          'change from decimal to percent format
          ws.Cells(2, 17).NumberFormat = "0.00%"
          ws.Cells(3, 17).NumberFormat = "0.00%"
            
          'find and write the ticker symbol for MaxPercent, MinPercent and MaxVolume
          For i = 2 To lastRow
          
            If ws.Cells(i, 11).Value = MaxPercent Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
                
            If ws.Cells(i, 11).Value = MinPercent Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
             
            If ws.Cells(i, 12).Value = MaxVolume Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
                
          Next i
            
    Next ws
        
End Sub

