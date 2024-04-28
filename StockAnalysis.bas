Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim maxPercentageIncrease As Double
    Dim maxPercentageDecrease As Double
    Dim maxVolume As Double
    Dim maxPercentageIncreaseTicker As String
    Dim maxPercentageDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim outputRow As Long
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row of data in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set initial output row
        outputRow = 2
        
        ' Output headers in the current worksheet
        ws.Range("H1").Value = "Ticker Symbol"
        ws.Range("I1").Value = "Quarterly Change"
        ws.Range("J1").Value = "Percentage Change"
        ws.Range("K1").Value = "Total Stock Volume"
        
        ' Reset variables for each worksheet
        maxPercentageIncrease = 0
        maxPercentageDecrease = 0
        maxVolume = 0
        
        ' Loop through each row of data in the current worksheet
        For i = 2 To lastRow
            ' Check if the next row has a different ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Store data for the current ticker
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                closingPrice = ws.Cells(i, 6).Value
                volume = ws.Cells(i, 7).Value
                
                ' Calculate quarterly change and percentage change
                quarterlyChange = closingPrice - openingPrice
                percentageChange = ((closingPrice - openingPrice) / openingPrice)
                
                ' Output data in the current worksheet
                ws.Cells(outputRow, 8).Value = ticker
                ws.Cells(outputRow, 9).Value = quarterlyChange
                ws.Cells(outputRow, 10).Value = percentageChange
                ws.Cells(outputRow, 11).Value = volume
                
                ' Find the stock with the greatest percentage increase
                If percentageChange > maxPercentageIncrease Then
                    maxPercentageIncrease = percentageChange
                    maxPercentageIncreaseTicker = ticker
                End If
                
                ' Find the stock with the greatest percentage decrease
                If percentageChange < maxPercentageDecrease Then
                    maxPercentageDecrease = percentageChange
                    maxPercentageDecreaseTicker = ticker
                End If
                
                ' Find the stock with the greatest total volume
                If volume > maxVolume Then
                    maxVolume = volume
                    maxVolumeTicker = ticker
                End If
                
                ' Move to the next output row
                outputRow = outputRow + 1
            End If
        Next i
        
        ' Output the stocks with the greatest percentage increase, percentage decrease, and total volume in the current worksheet
        ws.Range("L1").Value = "Greatest % Increase"
        ws.Range("M1").Value = "Greatest % Decrease"
        ws.Range("N1").Value = "Greatest Total Volume"
        
        ws.Range("L2").Value = maxPercentageIncreaseTicker
        ws.Range("M2").Value = maxPercentageDecreaseTicker
        ws.Range("N2").Value = maxVolumeTicker
    Next ws
End Sub

