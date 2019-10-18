Sub stocks()

For Each ws In Worksheets

'Declare Variables
    Dim lastrow As Long
    Dim summaryrow as Long
    Dim ticker As String
    Dim yearopen As Double
    Dim yearclose As Double
    Dim volume As Double
    Dim change As Double
    Dim pchange As Double

'Determine Last Row and Add Necessary Columns
    volume = 0
    summaryrow = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

'Sort Data
    ws.Range("A1:G" & lastrow).Sort , key1:=ws.Range("A1"), order1:=xlAscending, key2:=ws.Range("B2"), order2:=xlAscending, Header:=xlYes

'Determine Yearly Gain/Loss, % Gain/Loss, and Volume of Stocks
    For i = 2 To lastrow
        If ws.Cells(i, 3).Value = 0 and ws.Cells(i, 6).Value = 0 And ws.Cells(i, 7).Value = 0 Then
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        yearopen = ws.Cells(i, 3).Value
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        yearclose = ws.Cells(i, 6).Value
        volume = volume + ws.Cells(i, 7).Value
        change = yearclose-yearopen
        pchange = change/yearopen
        ws.Range("I" & summaryrow).Value = ticker
        ws.Range("J" & summaryrow).Value = change
        ws.Range("k" & summaryrow).Value = pchange
        ws.Range("L" & summaryrow).Value = volume
        summaryrow = summaryrow + 1
        volume = 0
        Else
        volume = volume + ws.Cells(i, 7).Value
        End If
    Next i

'Declare Additional Variables
    Dim lastsummaryrow as Long
    Dim greatsymbol as Range
    Dim leastsymbol as Range
    Dim greatvolsym as Range
    Dim greatchange as Double
    Dim leastchange as Double
    Dim greatvolume as Double

'Determine Last Row of Summarized Data, Add Necessary Columns and Find Min and Max % Change and Volume
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"   
    lastsummaryrow = ws.Cells(Rows.Count, 9).End(xlup).Row
    greatchange = WorksheetFunction.Max(ws.Range("K:K"))
    Set greatsymbol = ws.Range("K:K").Find(greatchange, Lookat:=xlWhole)
    leastchange = WorksheetFunction.Min(ws.Range("K:K"))
    Set leastsymbol = ws.Range("K:K").Find(leastchange, Lookat:=xlWhole)
    greatvolume = WorksheetFunction.Max(ws.Range("L:L"))
    Set greatvolsym = ws.Range("L:L").Find(greatvolume, Lookat:=xlWhole)
    ws.Range("O2") = greatsymbol.Offset(, -2)
    ws.Range ("P2") = greatchange
    ws.Range("O3") = leastsymbol.Offset(, -2)
    ws.Range ("P3") = leastchange
    ws.Range("O4") = greatvolsym.Offset(, -3)
    ws.Range ("P4") = greatvolume

'Fix Column Formating
    ws.Range("I:P").EntireColumn.AutoFit
    ws.Range("K:K").NumberFormat="0.00%"
    ws.Range("P2:P3").NumberFormat="0.00%"

'Format Background Color of Stock Summary Data Based on % Gain/Loss
    For j = 2 To lastsummaryrow
        If ws.Cells(j, 10).Value > 0 Then
        ws.Range(ws.Cells(j, 9), ws.Cells(j, 12)).Interior.ColorIndex = 4
        ElseIf ws.Cells(j, 10).Value < 0 Then
        ws.Range(ws.Cells(j, 9), ws.Cells(j, 12)).Interior.ColorIndex = 3
        End If
    Next j

Next ws

End Sub