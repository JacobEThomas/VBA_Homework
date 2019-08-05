Sub StockVolume():

' Loop thru worksheets
For Each ws In Worksheets
    With ws

' Defie variables
Dim Ticker As String
Dim Volume As Double
Dim StartRow As Integer

' Find total observations
TotObv = .Cells(Rows.Count, 1).End(xlUp).Row

' Tell the program where to start
StartRow = 2

' Name the New columns that will contain Ticker and Volume
.Range("I1").Value = "Ticker"
.Range("J1").Value = "Volume"

' Set initial volume
Volume = 0

' Loop thru rows to print ticker and calculate volume in new columns
For i = 2 To TotObv:
    If .Cells(i, 1).Value = .Cells(i + 1, 1).Value Then
    Volume = Volume + .Cells(i + 1, 7).Value
    Else: Ticker = .Cells(i, 1).Value
            .Range("I" & StartRow) = Ticker
            Volume = Volume
            .Range("J" & StartRow) = Volume
            StartRow = StartRow + 1
            Volume = 0
    End If
Next i
End With
Next ws
End Sub
