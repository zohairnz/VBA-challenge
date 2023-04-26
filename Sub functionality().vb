Sub functionality()

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

    'Insert functionality titles
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'initialize variable that will store the greatest increase and decrease percent changes, and greatest total volume
    Dim pt_increase As Variant
    Dim pt_decrease As Variant
    Dim gt_volume As Variant
    
    'Set original values to compare
    pt_increase = Cells(2, 11).Value
    pt_decrease = Cells(2, 11).Value
    gt_volume = Cells(2, 12).Value
    
    'Count the number of entries
    Dim entries As Long
    entries = Range("I1").End(xlDown).Row
    
    'Set indexing variable for ticker
    Dim index As Variant
    index = 2
    
    'Loop through entries for greatest % increase
    For i = 3 To entries
        If pt_increase < Cells(i, 11).Value Then
            pt_increase = Cells(i, 11).Value
            index = i
        End If
    Next i
    
    'List result
    Cells(2, 16).Value = Cells(index, 9).Value
    Cells(2, 17).Value = pt_increase
    
    'Loop through entries for greatest % decrease
    For i = 3 To entries
        If pt_decrease > Cells(i, 11).Value Then
            pt_decrease = Cells(i, 11).Value
            index = i
        End If
    Next i
    
    'List result
    Cells(3, 16).Value = Cells(index, 9).Value
    Cells(3, 17).Value = pt_decrease
    
    'Loop through entries for greatest % increase
    For i = 3 To entries
        If gt_volume < Cells(i, 12).Value Then
            gt_volume = Cells(i, 12).Value
            index = i
        End If
    Next i
    
    'List result
    Cells(4, 16).Value = Cells(index, 9).Value
    Cells(4, 17).Value = gt_volume
    
Next ws
    
End Sub
