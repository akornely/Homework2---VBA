Sub VBAEasy()

Dim current As Worksheet
For Each current In Worksheets

Dim ticker As String
Dim totalvolume As Double


Dim tickercount As Long
Dim lastrow As Long


tickercount = 2
volumetotal = 0
current.Cells(1, 9).Value = "ticker symbol"
current.Cells(1, 10).Value = "total volume"

lastrow = current.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
If current.Cells(i + 1, 1).Value <> current.Cells(i, 1).Value Then
ticker = current.Cells(i, 1).Value
totalvolume = totalvolume + current.Cells(i, 7).Value
current.Range("I" & tickercount).Value = ticker
current.Range("J" & tickercount).Value = totalvolume
tickercount = tickercount + 1
totalvolume = 0

Else
totalvolume = totalvolume + current.Cells(i, 7).Value

End If
Next i

Next

End Sub
