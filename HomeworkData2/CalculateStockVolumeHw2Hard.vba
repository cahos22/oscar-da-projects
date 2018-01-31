

Sub CalculateStockHardSample()

  Dim ticker As String
  Dim prevTicker As String
  Dim totalVolume As Double
  Dim sumOpen As Double
  Dim sumClose As Double
  Dim yearlyChange As Double
  Dim percentChange As Double
  
  Dim greatestPercentageIncrease As Double
  Dim greatestPercentageDecrease As Double
  Dim greatestTotalVolume As Double
  ' row index in summary table
  Dim sumRow As Integer

  Dim lastRow As Long
  
  Dim POSITIVE_COLOR As Integer
  Dim NEGATIVE_COLOR As Integer
  
  POSITIVE_COLOR = 4 'green
  NEGATIVE_COLOR = 3 'red
  
  Dim cellColor As Integer
   
  Dim greatestTable(1 To 4, 1 To 3) As String
  
  Dim percentageIncrIndex As Integer
  Dim percentageDecrIndex As Integer
  Dim grTotalVolumeIndex As Integer

  percentageIncrIndex = 2
  percentageDecrIndex = 3
  grTotalVolumeIndex = 4
  
  greatestTable(1, 1) = ""
  greatestTable(1, 1) = "Ticker"
  greatestTable(1, 1) = "Value"
  greatestTable(percentageIncrIndex, 1) = "Greatest % Increase"
  greatestTable(percentageDecrIndex, 1) = "Greatest % Decrease"
  greatestTable(grTotalVolumeIndex, 1) = "Greatest Total Volume"
    
   
  For Each ws In Worksheets
        totalVolume = 0
        sumOpen = 0
        sumClose = 0
        yearlyChange = 0
        percentChange = 0
        cellColor = 0
        sumRow = 2
        greatestPercentageIncrease = 0
        greatestPercentageDecrease = 0
        greatestTotalVolume = 0
        
     
        greatestTable(percentageIncrIndex, 2) = ""
        greatestTable(percentageIncrIndex, 3) = "0.00"
        greatestTable(percentageDecrIndex, 2) = ""
        greatestTable(percentageDecrIndex, 3) = "0.00"
        greatestTable(grTotalVolumeIndex, 2) = ""
        greatestTable(grTotalVolumeIndex, 3) = "0.00"
        
        
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
                  
                  ' Calculate Percentage change
                  percentChange = (sumClose / sumOpen) - 1
                  ws.Range("I1").Cells(sumRow, 3).Value = percentChange
              
                  ' Print Total Volume
                  ws.Range("I1").Cells(sumRow, 4).Value = totalVolume
                  
             
                  
                  ' Calculate greatest percentage increase
                  If (percentChange > 0# And percentChange > CDbl(greatestTable(percentageIncrIndex, 3))) Then
                    greatestTable(percentageIncrIndex, 2) = prevTicker
                    greatestTable(percentageIncrIndex, 3) = Str(percentChange)
                  End If
                  
                    ' Calculate greatest percentage decrease
                  If (percentChange < 0# And percentChange < CDbl(greatestTable(percentageDecrIndex, 3))) Then
                    greatestTable(percentageDecrIndex, 2) = prevTicker
                    greatestTable(percentageDecrIndex, 3) = Str(percentChange)
                  End If
                  

                  
                  ' Calculate greatest Total Volume
                  If totalVolume > CDbl(greatestTable(grTotalVolumeIndex, 3)) Then
                  greatestTable(grTotalVolumeIndex, 2) = prevTicker
                    greatestTable(grTotalVolumeIndex, 3) = totalVolume
                  End If
                  
                  ' Add one to the summary table row
                  sumRow = sumRow + 1
                  ' Start the new Brand Total
                  sumOpen = ws.Cells(i, 3)
                  sumClose = ws.Cells(i, 6)
                  totalVolume = ws.Cells(i, 7).Value
                  
                  prevTicker = ticker
            End If
        Next i
        ws.Range("K2:K" & sumRow).Style = "Percent"
        ws.Range("K2:K" & sumRow).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").Style = "Percent"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
            
        For row_i = 1 To 4
            num = ws.Cells(row_i, 3)
     
            For col_j = 1 To 3
                ws.Range("O1").Cells(row_i, col_j) = greatestTable(row_i, col_j)
                ws.Range("O1").Cells(row_i, col_j) = greatestTable(row_i, col_j)
                ws.Range("O1").Cells(row_i, col_j) = greatestTable(row_i, col_j)
            Next col_j
        Next row_i
        
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



