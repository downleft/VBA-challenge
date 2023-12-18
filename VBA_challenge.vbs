Sub stockcounterversion2()

    'Establish loop through all worksheets
    For Each ws In Worksheets
    
    '===========================================
    'For loop to work:
    'ws. needs to be added to each sheet-specific item (Ranges and Cells)
    '===========================================

    'Count the number of rows
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Establish Rolling Counters
    TickCount = 2
    OpenYear = ws.Range("C2")
    TotalCount = 0
    
    'Analyze Initial Data
    For i = 2 To LastRow
    
        'Begin Total for Stock
        TotalCount = TotalCount + Cells(i, 7)
    
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
        
            'Select Ticker Symbol
            ws.Cells(TickCount, 9) = ws.Cells(i, 1)

            'Output Yearly Change
            YearChange = ws.Cells(i, 6) - OpenYear
            ws.Cells(TickCount, 10) = YearChange
            
            'Output Percent Change
            ws.Cells(TickCount, 11) = YearChange / OpenYear
            
            'Total Stock Volume
            ws.Cells(TickCount, 12) = TotalCount
            
            'Count Resets for New Stock
            TotalCount = 0
            OpenYear = ws.Cells(i + 1, 3)
            TickCount = TickCount + 1
                
        End If
    Next i
    
    'Start Summary Analysis
    For k = 2 To (TickCount - 1)
    
        'Cell Color
        If ws.Cells(k, 10).Value < 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(k, 10).Interior.ColorIndex = 4
        End If
        
        'Find Greatest and Lowest % Change
        If ws.Cells(k, 11) > ws.Range("Q2") Then
            ws.Range("Q2") = ws.Cells(k, 11)
            ws.Range("P2") = ws.Cells(k, 9)
        ElseIf ws.Cells(k, 11) < ws.Range("Q3") Then
            ws.Range("Q3") = ws.Cells(k, 11)
            ws.Range("P3") = ws.Cells(k, 9)
        End If
        
        'Greatest Total Volume
        If ws.Cells(k, 12) > ws.Range("Q4") Then
            ws.Range("Q4") = ws.Cells(k, 12)
            ws.Range("P4") = ws.Cells(k, 9)
        End If
        
    Next k
    
    'Format Percentages
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("Q2", "Q3").NumberFormat = "0.00%"
    
    Next ws
    
End Sub