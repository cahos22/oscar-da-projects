Sub CalculateStockModerateSample()

  Dim ticker As String
  Dim prevTicker As String
  Dim totalVolume As Double
  Dim sumOpen As Double
  Dim sumClose As Double
  Dim yearlyChange As Double
  Dim percentChange As Double
  ' row index in summary table
  Dim sumRow As Integer

  Dim lastRow As Long
  
  Dim POSITIVE_COLOR As Integer
  Dim NEGATIVE_COLOR As Integer
  
  POSITIVE_COLOR = 4 'green
  NEGATIVE_COLOR = 3 'red
  
  Dim cellColor As Integer
   
  For Each ws In Worksheets
        totalVolume = 0
        sumOpen = 0
        sumClose = 0
        yearlyChange = 0
        percentChange = 0
        cellColor = 0
        sumRow = 2
        lastRow = ws.Range("A1048576").End(xlUp).Row
        prevTicker = ws.Cells(2, 1).Value
        ticker = ""
        ws.Range("I1").Cells(1, 1).Value = "Ticker"
        ws.Range("I1").Cells(1, 2).Value = "Yearly Change"
        ws.Range("I1").Cells(1, 3).Value = "Percent Change"
        ws.Range("I1").Cells(1, 4).Value = "Total Stock Volume"
       
        For i = 2 To lastRow + 1
            ticker = ws.Cells(i, 1).Value
            Debug.Print (i)
            If ticker = prevTicker Then
              'Add to the Brand Total
                 totalVolume = totalVolume + ws.Cells(i, 7).Value
                 sumOpen = sumOpen + ws.Cells(i, 3).Value
                 sumClose = sumClose + ws.Cells(i, 6).Value
            Else
                ' Print the Credit Card Brand in the Summary Table
                  ws.Range("I1").Cells(sumRow, 1).Value = prevTicker
                  
                  yearlyChange = sumClose - sumOpen
                  If yearlyChange > 0 Then
                     cellColor = POSITIVE_COLOR
                  Else
                     cellColor = NEGATIVE_COLOR
                  End If
                  
                  ws.Range("I1").Cells(sumRow, 2).Value = yearlyChange
                  ws.Range("I1").Cells(sumRow, 2).Interior.ColorIndex = cellColor
                  
                  
                  percentChange = (sumClose / sumOpen) - 1
                  ws.Range("I1").Cells(sumRow, 3).Value = percentChange
                 ' ws.Range("I1").Cells(sumRow, 3).Style = "Percent"
                  
                  ' Print the Brand Amount to the Summary Table
                  
                  ws.Range("I1").Cells(sumRow, 4).Value = totalVolume
                  ' Add one to the summary table row
                  sumRow = sumRow + 1
                  ' Start the new Brand Total
                  sumOpen = ws.Cells(i, 3)
                  sumClose = ws.Cells(i, 6)
                  totalVolume = ws.Cells(i, 7).Value
                  
                  prevTicker = ticker
            End If
        Next i
        ws.Range("K2:K" & i).Style = "Percent"
        ws.Range("K2:K" & i).NumberFormat = "0.0000%"
        
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



