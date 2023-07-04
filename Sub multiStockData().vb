Sub multiStockData()
' Variable declaration
  Dim openPrice As Double
  Dim totalSv As Double
  Dim closePrice As Double
  Dim i As Long
  Dim lastRow As Long
  Dim curSheet As Worksheet
  Dim currentName As String
  Dim nextName As String
  Dim yearlyChange As Double
  Dim percentChange As Double
  Dim outputRow As Integer
  Dim ticker As String
  Dim value As Long
  ' Loop through all the worksheets
  For Each curSheet In Worksheets
  outputRow = 2
    ' Assigning the cell values
    curSheet.Cells(1, 9).value = "Ticker"
    curSheet.Cells(1, 10).value = "Yearly Change"
    curSheet.Cells(1, 11).value = "Percent Change"
    curSheet.Cells(1, 12).value = "Total Stock Volume"
    curSheet.Cells(2, 15).value = "Greatest % Increase"
    curSheet.Cells(3, 15).value = "Greatest % decrease"
    curSheet.Cells(4, 15).value = "Greatest Total Volume"
    curSheet.Cells(1, 16).value = "Ticker"
    curSheet.Cells(1, 17).value = "value"
    totalSv = 0
    openPrice = curSheet.Cells(2, 3).value
    lastRow = curSheet.Cells(curSheet.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
      currentName = curSheet.Cells(i, 1).value
      nextName = curSheet.Cells(i + 1, 1).value
      ' Check if it's a new ticker
      If nextName = currentName Then
        totalSv = totalSv + curSheet.Cells(i, 7).value
      Else
        closePrice = curSheet.Cells(i, 6).value
        yearlyChange = closePrice - openPrice
        percentChange = yearlyChange / openPrice
        curSheet.Cells(outputRow, 9).value = currentName
        curSheet.Cells(outputRow, 10).value = yearlyChange
        curSheet.Cells(outputRow, 11).value = percentChange
        curSheet.Cells(outputRow, 12).value = totalSv
' Apply conditional formatting
        If yearlyChange > 0 Then
          curSheet.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
        ElseIf yearlyChange < 0 Then
          curSheet.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
        End If
        outputRow = outputRow + 1
        totalSv = 0
        openPrice = curSheet.Cells(i + 1, 3).value
        End If
        Next i
        
          'bonus
    'greatest percentage increase
   curSheet.Range("Q2") = WorksheetFunction.Max(curSheet.Range("f2:K" & lastRow))
   'greatest percentage decrease
   curSheet.Range("Q3") = WorksheetFunction.Min(curSheet.Range("F2:K" & lastRow))
   'greatest total volume
   curSheet.Range("q4") = WorksheetFunction.Max(curSheet.Range("L2:l" & lastRow))
 increasenumber = WorksheetFunction.Match(WorksheetFunction.Max(curSheet.Range("K2:K" & lastRow)), curSheet.Range("K2:K" & lastRow), 0)
 decreasenumber = WorksheetFunction.Match(WorksheetFunction.Min(curSheet.Range("K2:k" & lastRow)), curSheet.Range("k2:k" & lastRow), 0)
 increasetSotalvolume = WorksheetFunction.Match(WorksheetFunction.Max(curSheet.Range("L2:L" & lastRow)), curSheet.Range("L2:L" & lastRow), 0)
 curSheet.Range("p2") = curSheet.Cells(increasenumber + 1, 9)
 curSheet.Range("p3") = curSheet.Cells(decreasenumber + 1, 9)
 curSheet.Range("p4") = curSheet.Cells(increasetotalvolume + 1, 9)
'  curSheet.Range("p2") = curSheet.WorksheetFunction.Match(WorksheetFunction.Max(curSheet.Range("K2:K" & lastRow)), curSheet.Range("K2:K" & lastRow), 0) + 1, 9).value
'   curSheet.Range("p3") = curSheet.WorksheetFunction.Match(WorksheetFunction.Min(curSheet.Range("K2:K" & lastRow)), curSheet.Range("K2:K" & lastRow), 0) + 1, 9).value
'    curSheet.Range("p4") = curSheet.WorksheetFunction.Match(WorksheetFunction.Max(curSheet.Range("L2:L" & lastRow)), curSheet.Range("L2:L" & lastRow), 0) + 1, 9).value
      
  Next curSheet
End Sub

 

