Attribute VB_Name = "Module1"
Sub AnalyzeStocks()
    Dim worksheetObj As Worksheet
    Dim stockTicker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim rowIdx As Long
    Dim lastRow As Long
    Dim summaryRow As Integer
    Dim openPrice As Double
    Dim closePrice As Double
    Dim priceRow As Long
    
    For Each worksheetObj In ThisWorkbook.Worksheets
        lastRow = worksheetObj.Cells(worksheetObj.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        priceRow = 2

        worksheetObj.Cells(1, 9).Value = "Stock Ticker"
        worksheetObj.Cells(1, 10).Value = "Yearly Change"
        worksheetObj.Cells(1, 11).Value = "Percent Change"
        worksheetObj.Cells(1, 12).Value = "Total Volume"

        For rowIdx = 2 To lastRow
            If worksheetObj.Cells(rowIdx + 1, 1).Value <> worksheetObj.Cells(rowIdx, 1).Value Then
                stockTicker = worksheetObj.Cells(rowIdx, 1).Value
                totalVolume = totalVolume + worksheetObj.Cells(rowIdx, 7).Value

                worksheetObj.Cells(summaryRow, 9).Value = stockTicker
                worksheetObj.Cells(summaryRow, 12).Value = totalVolume

                openPrice = worksheetObj.Cells(priceRow, 3).Value
                closePrice = worksheetObj.Cells(rowIdx, 6).Value
                yearlyChange = closePrice - openPrice

                If openPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openPrice
                End If

                worksheetObj.Cells(summaryRow, 10).Value = yearlyChange
                worksheetObj.Cells(summaryRow, 11).Value = percentChange
                worksheetObj.Cells(summaryRow, 11).NumberFormat = "0.00%"

                If yearlyChange > 0 Then
                    worksheetObj.Cells(summaryRow, 10).Interior.ColorIndex = 4
                Else
                    worksheetObj.Cells(summaryRow, 10).Interior.ColorIndex = 3
                End If

                summaryRow = summaryRow + 1
                priceRow = rowIdx + 1
                totalVolume = 0
            Else
                totalVolume = totalVolume + worksheetObj.Cells(rowIdx, 7).Value
            End If
        Next rowIdx
        
        Dim maxPercentIncrease As Double
        Dim maxPercentChangeTicker As String
        Dim minPercentDecrease As Double
        Dim minPercentChangeTicker As String
        Dim maxTotalVolume As Double
        Dim maxTotalVolumeTicker As String
        
        maxPercentIncrease = Application.WorksheetFunction.Max(worksheetObj.Range(worksheetObj.Cells(2, 11), worksheetObj.Cells(summaryRow - 1, 11)).Value)
        minPercentDecrease = Application.WorksheetFunction.Min(worksheetObj.Range(worksheetObj.Cells(2, 11), worksheetObj.Cells(summaryRow - 1, 11)).Value)
        maxTotalVolume = Application.WorksheetFunction.Max(worksheetObj.Range(worksheetObj.Cells(2, 12), worksheetObj.Cells(summaryRow - 1, 12)).Value)
        
        worksheetObj.Cells(2, 17).NumberFormat = "0.00%"
        worksheetObj.Cells(3, 17).NumberFormat = "0.00%"

        For rowIdx = 2 To summaryRow - 1
            If worksheetObj.Cells(rowIdx, 11).Value = maxPercentIncrease Then
                maxPercentChangeTicker = worksheetObj.Cells(rowIdx, 9).Value
                Exit For
            End If
        Next rowIdx
        
        For rowIdx = 2 To summaryRow - 1
            If worksheetObj.Cells(rowIdx, 11).Value = minPercentDecrease Then
                minPercentChangeTicker = worksheetObj.Cells(rowIdx, 9).Value
                Exit For
            End If
        Next rowIdx
        
        For rowIdx = 2 To summaryRow - 1
            If worksheetObj.Cells(rowIdx, 12).Value = maxTotalVolume Then
                maxTotalVolumeTicker = worksheetObj.Cells(rowIdx, 9).Value
                Exit For
            End If
        Next rowIdx
        
        worksheetObj.Cells(1, 16).Value = "Stock"
        worksheetObj.Cells(1, 17).Value = "Value"
        worksheetObj.Cells(2, 16).Value = maxPercentChangeTicker
        worksheetObj.Cells(2, 17).Value = maxPercentIncrease
        worksheetObj.Cells(3, 16).Value = minPercentChangeTicker
        worksheetObj.Cells(3, 17).Value = minPercentDecrease
        worksheetObj.Cells(2, 15).Value = "Max % Increase"
        worksheetObj.Cells(3, 15).Value = "Min % Decrease"
        worksheetObj.Cells(4, 15).Value = "Max Total Volume"
        worksheetObj.Cells(4, 16).Value = maxTotalVolumeTicker
        worksheetObj.Cells(4, 17).Value = maxTotalVolume
    Next worksheetObj
End Sub

