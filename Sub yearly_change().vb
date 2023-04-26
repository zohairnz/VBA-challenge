Sub yearly_change()

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

    'Count the number of entries
    Dim entries As Long
    entries = Range("A1").End(xlDown).Row

    'create index for ticker entries
    Dim index As Integer
    index = 2
    
    'create variable to store intial and final price
    Dim initial_price As Variant
    Dim final_price As Variant
    initial_price = Cells(2, 3).Value
    final_price = 0
    
    'initialize variable that will store the change and percent change
    Dim yr_change As Variant
    Dim pt_change As Variant
    
    'Set title entries
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    
    'Loop through ticker names
    For i = 2 To entries + 1
        If Cells(i, 1) <> Cells(index, 9) Then
        final_price = Cells(i - 1, 6)
        yr_change = final_price - initial_price
        pt_change = yr_change / initial_price
        Cells(index, 10).Value = yr_change
        Cells(index, 11).Value = pt_change
        If Cells(index, 10).Value >= 0 Then
            Cells(index, 10).Interior.ColorIndex = 4
        Else
            Cells(index, 10).Interior.ColorIndex = 3
        End If
        If Cells(index, 11).Value >= 0 Then
            Cells(index, 11).Interior.ColorIndex = 4
        Else
            Cells(index, 11).Interior.ColorIndex = 3
        End If
        initial_price = Cells(i, 3).Value
        index = index + 1
        End If
        
    Next i
Next ws
    
End Sub
