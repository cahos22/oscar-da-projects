
'Q: how is this different?
'How is this more efficient?
Sub CalculateTotalStockVolume()

  Dim ticker As String
  Dim prevTicker As String
  Dim totalVolume As Double

  ' row index in summary table
  Dim sumRow As Integer

  Dim lastRow As Long
  
  For Each ws In Worksheets
        totalVolume = 0
        sumRow = 2
        lastRow = ws.Range("A1048576").End(xlUp).Row
        prevTicker = ws.Cells(2, 1).Value
        ticker = ""
        ws.Range("I1").Cells(1, 1).Value = "Ticker"
        ws.Range("I1").Cells(1, 2).Value = "Total Stock Volume"
        
        For i = 2 To lastRow + 1
            ticker = ws.Cells(i, 1).Value
            Debug.Print (i)
            If ticker = prevTicker Then
              'Add to the Brand Total
                 totalVolume = totalVolume + ws.Cells(i, 7).Value
            Else
                ' Print the Credit Card Brand in the Summary Table
                  ws.Range("I1").Cells(sumRow, 1).Value = prevTicker
                  ' Print the Brand Amount to the Summary Table
                  ws.Range("I1").Cells(sumRow, 2).Value = totalVolume
                  ' Add one to the summary table row
                  sumRow = sumRow + 1
                  ' Start the new Brand Total
                  totalVolume = ws.Cells(i, Columns.Count).Value
                  prevTicker = ticker
            End If
        Next i
  Next ws

End Sub

Sub ClearCellsSample()

  Dim lastRow As Long
  Dim sumRow As Long
  
  For Each ws In Worksheets
        sumRow = 2
        lastRow = ws.Range("A1048576").End(xlUp).Row
        ws.Range("I1").Cells(1, 1).Value = ""
        ws.Range("J1").Cells(1, 1).Value = ""
        ws.Range("K1").Cells(1, 1).Value = ""
        ws.Range("L1").Cells(1, 1).Value = ""
        ws.Range("J1").EntireColumn.Interior.ColorIndex = 0
        
        For i = 2 To lastRow + 1
            ticker = ws.Cells(i, 1).Value
            Debug.Print (i)
            If ticker = prevTicker Then
              'Add to the Brand Total
                 totalVolume = totalVolume + ws.Cells(i, 7).Value
            Else
   
                  ws.Range("I1").Cells(sumRow, 1).Value = ""
                  ws.Range("J1").Cells(sumRow, 1).Value = ""
                  ws.Range("K1").Cells(sumRow, 1).Value = ""
                  ws.Range("L1").Cells(sumRow, 1).Value = ""
                  sumRow = sumRow + 1

            End If
        Next i
  Next ws

End Sub



