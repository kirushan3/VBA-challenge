Sub uniqueidentifier()

'insert new column header
Range("J1").Value = "Ticker Symbol"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"

'Declare variables
Dim X As Integer
Dim lastrow As Long
Dim stockvol As Double
Dim firstopen As Double
Dim lastclose As Double


X = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row


firstopen = Cells(2, 3)


'Loop criteria

For Row = 2 To lastrow
stockvol = stockvol + Cells(Row, 7).Value

'If statements ticker symbol, total stock volume, yearly change, percent change

    If Cells(Row + 1, 1).Value <> Cells(Row, 1) Then
    Cells(X, 10).Value = Cells(Row, 1)

    Cells(X, 13).Value = stockvol
    
   
   lastclose = Cells(Row, 6)
   
   yearlychange = lastclose - firstopen
   
   'to avoid division 0 error
   
   If firstopen = 0 Then
   percentchange = 0
   
   Else
   percentchange = (lastclose - firstopen) / firstopen
   
   End If
   
   firstopen = Cells(Row + 1, 3)
   Cells(X, 11).Value = yearlychange
   
   'colour format
   If Cells(X, 11).Value < 0 Then
   
   Cells(X, 11).Interior.ColorIndex = 3
   
   Else
   
   Cells(X, 11).Interior.ColorIndex = 4
   
   End If
   
   
   Cells(X, 12).Value = percentchange
   
    X = X + 1
    stockvol = 0
    
    End If
        
Next Row

'percent formatting for percent change column
Dim lastrow2 As Long
lastrow2 = Cells(2, 12).End(xlDown).Row
Range("L2:L" & lastrow2).NumberFormat = "0.00%"

End Sub
