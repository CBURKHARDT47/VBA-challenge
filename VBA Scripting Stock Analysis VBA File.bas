Attribute VB_Name = "Module1"
Sub ThreeYearStockAnalysis():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        
        Dim i As Integer
        Dim j As Integer
        
        Dim TickerDay As Integer
        Dim LastRowA As Integer
        Dim LastRowI As Integer
        
        Dim PercentChange As Double
        Dim GreatestIncr As Double
        Dim GreatestDecr As Double
        Dim GreatestVol As Double
        
        'Set the WorksheetName
        WorksheetName = ws.Name
        
        'Create column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker Counter to first row
        TickerDay = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last populated cell in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

            For i = 2 To LastRowA
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I
                ws.Cells(TickerDay, 9).Value = ws.Cells(i, 1).Value
            
                'Calculate and write Yearly Change in column J
                ws.Cells(TickerDay, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                'Conditional formating for red if yearly change is less than 0 otherwise fill color will be green
                If ws.Cells(TickerDay, 10).Value < 0 Then
                
                    'Set cell background color to red if no yearly growth
                    ws.Cells(TickerDay, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green for stocks that had growth
                    ws.Cells(TickerDay, 10).Interior.ColorIndex = 4
                
                    End If
                    
                'Calculate apercent change in column K
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                'Percent formatting
                    ws.Cells(TickerDay, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerDay, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate total volume in column L
                ws.Cells(TickerDay, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickerDay by 1
                TickerDay = TickerDay + 1
                
                'Set new start row of the ticker abbreviation check
                j = i + 1
                
                End If
            
            Next i
            
        'Find last non-populated cell at bottom of column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
        'Calculations of Greatest
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
        For i = 2 To LastRowI
            
                'For greatest total volume check if next value is larger if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                'For greatest increase check if next value is larger and populate ws.Cells if so
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestIncr = GreatestIncr
                
                End If
                
                'For greatest decrease check if next value is smaller and populate ws.Cells if so
                If ws.Cells(i, 11).Value < GreatestDecr Then
                GreatestDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestDecr = GreatestDecr
                
                End If
                
            'Write results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatestIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatestVol, "Scientific")
            
            Next i
            
        'adjust column width to fit across all populated columns
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
