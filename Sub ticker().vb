Sub ticker()

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

    'Count the number of entries
    Dim entries As Long
    entries = Range("A1").End(xlDown).Row

    'create index for ticker entries
    Dim index As Integer
    index = 2
    
    'Set first ticker entry
    Cells(1, 9).Value = "Ticker"
    Cells(2, 9).Value = Cells(index, 1)
    
    
    'Loop through ticker names and extract unique names
    For i = 3 To entries
        If Cells(i, 1) <> Cells(index, 9) Then
        index = index + 1
        Cells(index, 9).Value = Cells(i, 1)
        End If
        
    Next i
Next ws

    
End Sub