Attribute VB_Name = "Module1"
Sub YearlyChangeWorkbook()
    Dim ws As Worksheet
    Dim ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim i As Long
    Dim lastRow As Long
    Dim SummaryTableRow As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PriceRow As Long
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        SummaryTableRow = 2
        PriceRow = 2

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

                ws.Cells(SummaryTableRow, 9).Value = ticker
                ws.Cells(SummaryTableRow, 12).Value = TotalStockVolume

                OpenPrice = ws.Cells(PriceRow, 3).Value
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice

                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If

                ws.Cells(SummaryTableRow, 10).Value = YearlyChange
                ws.Cells(SummaryTableRow, 11).Value = PercentChange
                ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"

                If YearlyChange > 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                End If

                SummaryTableRow = SummaryTableRow + 1
                PriceRow = i + 1
                TotalStockVolume = 0
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        Dim maxPercentIncrease As Double
        Dim maxPercentChangeTicker As String
        Dim minPercentDecrease As Double
        Dim minPercentChangeTicker As String
        Dim maxTotalVolume As Double
        Dim maxTotalVolumeTicker As String
        
        maxPercentIncrease = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(SummaryTableRow - 1, 11)).Value)
        minPercentDecrease = Application.WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(SummaryTableRow - 1, 11)).Value)
        maxTotalVolume = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(SummaryTableRow - 1, 12)).Value)
        
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"

        For i = 2 To SummaryTableRow - 1
            If ws.Cells(i, 11).Value = maxPercentIncrease Then
                maxPercentChangeTicker = ws.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        For i = 2 To SummaryTableRow - 1
            If ws.Cells(i, 11).Value = minPercentDecrease Then
                minPercentChangeTicker = ws.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        For i = 2 To SummaryTableRow - 1
            If ws.Cells(i, 12).Value = maxTotalVolume Then
                maxTotalVolumeTicker = ws.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = maxPercentChangeTicker
        ws.Cells(2, 17).Value = maxPercentIncrease
        ws.Cells(3, 16).Value = minPercentChangeTicker
        ws.Cells(3, 17).Value = minPercentDecrease
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        ws.Cells(4, 16).Value = maxTotalVolumeTicker
        ws.Cells(4, 17).Value = maxTotalVolume
    Next ws
End Sub

