Attribute VB_Name = "Module1"
Sub TickerAnalyzer()
    'Define dimensions
    Dim ws As Worksheet
    Dim total As Double
    Dim i As Long
    Dim j As Integer
    Dim change As Double
    Dim percentChange As Double
    Dim start As Long
    Dim rowCount As Long
    
    For Each ws In Worksheets
        'Set starting values
        total = 0
        change = 0
        j = 0
        start = 2
        
        'Set header cells
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        
        'Get row number of last row with data
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To rowCount
            'If the ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Store result in variable
                total = total + ws.Cells(i, 7).Value
                
                If total = 0 Then
                'Print results
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = 0 & " %"
                ws.Range("L" & 2 + j).Value = 0
                
                Else
                    'Find first non zero row
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    
                    'Calculate change
                    change = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
                    percentChange = Round((change / ws.Cells(start, 3).Value) * 100, 2)
                    
                    'Start of next stock ticker
                    start = i + 1
                    
                    'Print results
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Round(change, 2)
                    ws.Range("K" & 2 + j).Value = percentChange & " %"
                    ws.Range("L" & 2 + j).Value = total
                
                    'Conditional formatting
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                  End If
                    'Reset ticker variables
                    j = j + 1
                    total = 0
                    change = 0
                    
            Else
                total = total + ws.Cells(i, 7).Value
            End If
        Next i
        
        'Find greatest values
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100 & " %"
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100 & " %"
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        'Return on less because header row not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        
        'Set ticker name
        ws.Range("P2").Value = ws.Range("I" & increase_number + 1).Value
        ws.Range("P3").Value = ws.Range("I" & decrease_number + 1).Value
        ws.Range("P4").Value = ws.Range("I" & volume_number + 1).Value
        
        'Set cell formats
        ws.Range("L2:L" & rowCount).NumberFormat = "0"
        ws.Range("Q4").NumberFormat = "0"
    Next ws
End Sub
