Attribute VB_Name = "Module6"
Sub uniqueidentifier()
'Loop thru all sheets and declare active sheet variable
Dim sheet As Worksheet
Dim starting_sheet As Worksheet
Set starting_sheet = ActiveSheet
For Each sheet In ThisWorkbook.Worksheets
    sheet.Activate
'code
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
'challenge header
Cells(2, 15).Value = "Greatest Percent Increase"
Cells(3, 15).Value = "Lowest Percent Increase"
Cells(4, 15).Value = "Greatest Total Volume"
Range("P1").Value = "Ticker S"
Range("Q1").Value = "Value"
Cells(4, 17).ColumnWidth = 20
'percent formatting for percent change column
Dim lastrow2 As Long
lastrow2 = Cells(2, 12).End(xlDown).Row
Range("L2:L" & lastrow2).NumberFormat = "0.00%"
'challenge
For maxi = 2 To lastrow2
    If Cells(maxi, 12) = Application.WorksheetFunction.Max(Range("L2:L" & lastrow2)) Then
    Cells(2, 17).Value = Cells(maxi, 12).Value
    Cells(2, 16).Value = Cells(maxi, 10).Value
    Cells(2, 17).NumberFormat = "0.00%"
    End If
Next maxi
For Min = 2 To lastrow2
    If Cells(Min, 12) = Application.WorksheetFunction.Min(Range("L2:L" & lastrow2)) Then
    Cells(3, 17).Value = Cells(Min, 12).Value
    Cells(3, 16).Value = Cells(Min, 10).Value
    Cells(3, 17).NumberFormat = "0.00%"
    End If
Next Min
Dim lastrow3 As Long
lastrow3 = Cells(2, 13).End(xlDown).Row
For Valu = 2 To lastrow3
    If Cells(Valu, 13) = Application.WorksheetFunction.Max(Range("M2:M" & lastrow3)) Then
    Cells(4, 17).Value = Cells(Valu, 13).Value
    Cells(4, 16).Value = Cells(Valu, 10).Value
    End If
Next Valu
sheet.Cells(1, 1) = 1 'This sets cell A1 to each sheet to 1
Next 'This will close the loop from the top
starting_sheet.Activate 'Activate the worksheet that was originally active  'This will bring you back to your original active sheet.
End Sub
