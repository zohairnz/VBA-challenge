Sub total_stock_volume()

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

    'Count the number of entries
    Dim entries As Long
    entries = Range("A1").End(xlDown).Row

    'create index for ticker entries
    Dim index As Integer
    index = 2
    
    'create variable to store stock volume
    Dim stock_volume As Variant
    stock_volume = 0
    
    'Set first volume entry
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Loop through ticker names and extract unique names
    For i = 2 To entries + 1
        If Cells(i, 1) = Cells(index, 9) Then
        stock_volume = stock_volume + Cells(i, 7).Value
        Else
        Cells(index, 12).Value = stock_volume
        stock_volume = Cells(i, 7).Value
        index = index + 1
        End If
        
    Next i
Next ws

    
End Sub