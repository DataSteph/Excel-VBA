```
Sub stockcode()
    Dim ws As Worksheet
    Set ws = Worksheets("2014")
    Dim ticker As String
    Dim totalvolume As Double
    Dim sumrow As Double
    Dim lastrow As Double
    sumrow = 2
    totalvolume = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker = ws.Cells(2, 1).Value
    
    For i = 2 To lastrow
        Debug.Print (i)
        If ticker = ws.Cells(i + 1, 1).Value Then
            totalvolume = totalvolume + ws.Cells(i, 7).Value
        Else
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            ws.Range("I1").Cells(1, 1).Value = "Ticker"
            ws.Range("I1").Cells(1, 2).Value = "Total Volume"
            ws.Range("I1").Cells(sumrow, 1).Value = ticker
            ws.Range("I1").Cells(sumrow, 2).Value = totalvolume
        sumrow = sumrow + 1
        totalvolume = 0
        ticker = ws.Cells(i + 1, 1).Value
    End If
    Next i
End Sub
```