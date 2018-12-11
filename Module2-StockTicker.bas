Attribute VB_Name = "Module1"
Sub LoopThroughWorkbook()
Dim ws As Worksheet
For Each ws In Worksheets
ws.Select
Call StockTicker
ws.Columns("I:Q").AutoFit
Application.Goto Range("A1"), True
ActiveWindow.VisibleRange(1, 1).Select
Next
End Sub

Sub StockTicker()

'set variables
Dim Ticker As Variant
Dim Ticker_Total As Variant
Dim LastRow As Variant
Dim StockOpen As Variant
Dim StockClose As Variant
Dim Difference As Variant
Dim Summary_Row As Variant
Dim Percent_Change As Variant
Dim GreatestIncrease As Variant
Dim GreatestDecrease As Variant
Dim GreatestVolume As Variant
StockOpen = Cells(2, 3).Value

'set initial values and headers
LastRow = Cells(Rows.Count, 2).End(xlUp).Row
Summary_Row = 2
Range("I1").Value = "Ticker"
Range("M1").Value = "Total Change"
Range("N1").Value = "% Change"
Range("J1").Value = "Volume"
Range("K1").Value = "Stock Open"
Range("L1").Value = "Stock Close"
Range("K2").Value = Cells(2, 3)
Range("I1:N1").Font.Bold = True
Range("I1:N1").HorizontalAlignment = xlCenter
Range("I1:N1").Interior.ColorIndex = 15
Range("I1:N1").Borders.ColorIndex = 56

Range("P1").Value = "Stat"
Range("P2").Value = "Greatest % Increase"
Range("P3").Value = "Greatest % Decrease"
Range("P4").Value = "Greatest Total Volume"
Range("Q1").Value = "Ticker"
Range("R1").Value = "Value"
Range("P1:R1").Font.Bold = True
Range("P1:R1").HorizontalAlignment = xlCenter
Range("P1:R1").Interior.ColorIndex = 15
Range("P1:R1").Borders.ColorIndex = 56


'set formatting
Range("J:M").Style = "Currency"
Range("N:N").NumberFormat = "0.00%"


'Start Loop to go through rows and check for a different stock ticker. If different get ticker symbol and amounts
For i = 2 To LastRow

'If the stock open value is not 0 then do the following
If StockOpen <> 0 Then
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value
StockClose = Cells(i, 6).Value
YearlyChange = StockClose - StockOpen
Ticker_Total = Ticker_Total + Cells(i, 7).Value

'Calculate ticker total and print to summary table
Range("I" & Summary_Row).Value = Ticker
Range("J" & Summary_Row).Value = Ticker_Total
Range("K" & Summary_Row).Value = StockOpen
Range("L" & Summary_Row).Value = StockClose
Range("M" & Summary_Row).Value = YearlyChange
Range("N" & Summary_Row).Value = (YearlyChange / StockOpen)



'Reset ticker total to 0 for next ticker set and get stock open
Summary_Row = Summary_Row + 1
Ticker_Total = 0
StockOpen = Cells(i + 1, 3).Value

Else
Ticker_Total = Ticker_Total + Cells(i, 7).Value



End If
End If

Next i

'Add new variables and values for hard homework
GreatestIncrease = Application.WorksheetFunction.Max(Columns("N"))
Range("R2").Value = GreatestIncrease
Range("R2").NumberFormat = "0.00%"
GreatestDecrease = Application.WorksheetFunction.Min(Columns("N"))
Range("R3").Value = GreatestDecrease
Range("R3").NumberFormat = "0.00%"
GreatestVolume = Application.WorksheetFunction.Max(Columns("J"))
Range("R4").Value = GreatestVolume
Range("R4").Style = "Currency"

'Find matching ticker value to max, min, and greatest value determined above
For i = 2 To LastRow
If Cells(i, 14).Value = GreatestIncrease Then
    Range("Q2").Value = Cells(i, 9).Value
    
ElseIf Cells(i, 14).Value = GreatestDecrease Then
    Range("Q3").Value = Cells(i, 9).Value
    
ElseIf Cells(i, 10).Value = GreatestVolume Then
    Range("Q4").Value = Cells(i, 9).Value

End If
Next i

'Set conditional formatting
Dim rg As Range
Dim cond1 As FormatCondition, cond2 As FormatCondition
Set rg = Range("M2:N2", Range("M2:N2").End(xlDown))
 
'clear any existing conditional formatting
rg.FormatConditions.Delete
 
'define the rule for each conditional format
Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "0")
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "0")

 
'define the format applied for each conditional format
With cond1
.Interior.Color = vbGreen
.Font.Color = vbBlack
End With
 
With cond2
.Interior.Color = vbRed
.Font.Color = vbWhite
End With


End Sub





