Sub stockTrackerHW()
    
    'iterate through each sheet
    For Each Sheet In Worksheets
    
        'adding titles to each sheet
        Sheet.Cells(1, 9) = "Ticker"
        Sheet.Cells(1, 10) = "Yearly Change"
        Sheet.Cells(1, 11) = "Percent Change"
        Sheet.Cells(1, 12) = "Total Volume"
        
        Sheet.Cells(1, 15) = "Ticker"
        Sheet.Cells(1, 16) = "Value"
        Sheet.Cells(2, 14) = "Greatest % Increase"
        Sheet.Cells(3, 14) = "Greatest % Decrease"
        Sheet.Cells(4, 14) = "Greatest Total Volume"
              
        'grab the starting open value
        Dim startingValue As Double
        startingValue = Sheet.Cells(2, 3)
            
        'Starting index for the tickers
        Dim tickerCount As Integer
        tickerStart = 2
            
        'grab first volume
        Dim total_volume As Double
        total_volume = 0
            
        'Dim currVolume As Double
        'currVolume = 0
            
                    
        'determine the last row
        lastRow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
            
        For i = 2 To lastRow
            
            'grab the current cell
            Dim currentName As String
            currentName = Sheet.Cells(i, 1)
                
            'update the volume
            Dim currVolume As Long
            currVolume = Sheet.Cells(i, 7)
            'MsgBox (currVolume)
            
            'add currVolume to totalVolume
            total_volume = total_volume + currVolume
                
            'grab the next cell
            Dim nextName As String
            nextName = Sheet.Cells(i + 1, 1).Value
                
            'if we encounter a different name
            If currentName <> nextName Then
            
                'fill in the current spot for the currentName
                Sheet.Cells(tickerStart, 9) = currentName
                
                'update the name to be searched for
                currentName = nextName
                
                'Grab the closng value at the end
                Dim closingValue As Double
                closingValue = Sheet.Cells(i, 6)
                
                'Assign the difference to annual change, as well as the Percent change
                Sheet.Cells(tickerStart, 10) = closingValue - startingValue
                
                'conditional formatting the fill
                If closingValue - startingValue > 0 Then
                    Sheet.Cells(tickerStart, 10).Interior.ColorIndex = 4
                ElseIf closingValue - startingValue < 0 Then
                    Sheet.Cells(tickerStart, 10).Interior.ColorIndex = 3
                Else
                    Sheet.Cells(tickerStart, 10).Interior.ColorIndex = 5
                End If
                
                'Divide by zero case (if it happens)
                If startingValue = 0 Then
                    Sheet.Cells(tickerStart, 11).Value = "null"
                Else
                    Sheet.Cells(tickerStart, 11) = (closingValue - startingValue) / startingValue
                End If
                
                'formatting the decimals to two places only
                Cells(tickerStart, 11).NumberFormat = "0.00%"
                    
                'update the startingValue
                startingValue = Sheet.Cells(i + 1, 3)
            
                'fill in the total_volume
                Sheet.Cells(tickerStart, 12) = total_volume
                
                'reset the volume
                total_volume = 0
                
                'Move down one row
                tickerStart = tickerStart + 1
                    
            End If

                
        Next i
    
            '------------------------------------------------BONUS SECTION(maybe) --------------------------------------------
                
			
                
            '------------------------------------------------END OF BONUS SECTION --------------------------------------------
                    
    
    Next Sheet
    
        
End Sub

