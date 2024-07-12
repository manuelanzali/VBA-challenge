
Sub VBAchallenge()
    'Create Variables
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    '---------------------------
    'Loop through all sheets
    '---------------------------
    For Each ws In ThisWorkbook.Sheets
        '---------------------------
        'Extract and Insert the ws name
        '---------------------------
        'Create Variables to hold Ticker Symbol as a string, Quarterly Change as a Double, Percentage Change as a Double, Total Stock Volume as a Long, Stock Volume as Long, Opening Price as Double, and Closing Price as Double
        Dim lastRow As Long
        Dim Tickersymbol As String
        Dim QuarterlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        Dim StockVolume As Double
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        'Determine the last row
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        'Add Columns for the "Ticker", "Quarterly Change", "Percent Change", and "Total Stock Volume" once
        If ws.Cells(1, 9).Value <> "Ticker" Then
            ws.Cells(1, 9).Resize(1, 4).Value = Array("Ticker", "Quarterly Change", "Percent Change", "Total Stock Volume")
        End If
        '------------------------
        'Calculate Values for "Quarterly Change" and "Percent Change"
        '---
        ' Arrays to hold data
        Dim openingPrices() As Double
        Dim closingPrices() As Double
        Dim stockVolumes() As Double
        ' Initialize arrays
        ReDim openingPrices(2 To lastRow)
        ReDim closingPrices(2 To lastRow)
        ReDim stockVolumes(2 To lastRow)
        ' Load data into arrays
        For i = 2 To lastRow
            If IsNumeric(ws.Cells(i, 3).Value) Then
                openingPrices(i) = ws.Cells(i, 3).Value
            Else
                openingPrices(i) = 0
            End If
            If IsNumeric(ws.Cells(i, 6).Value) Then
                closingPrices(i) = ws.Cells(i, 6).Value
            Else
                closingPrices(i) = 0
            End If
            If IsNumeric(ws.Cells(i, 7).Value) Then
                stockVolumes(i) = ws.Cells(i, 7).Value
            Else
                stockVolumes(i) = 0
            End If
        Next i
        ' Calculate changes and write to sheet
        Dim percentChanges() As Double
        Dim quarterlyChanges() As Double
        ReDim percentChanges(2 To lastRow)
        ReDim quarterlyChanges(2 To lastRow)
        For i = 2 To lastRow
            QuarterlyChange = closingPrices(i) - openingPrices(i)
            If openingPrices(i) <> 0 Then
                PercentChange = (QuarterlyChange / openingPrices(i)) * 100
            Else
                PercentChange = 0
            End If
            quarterlyChanges(i) = QuarterlyChange
            percentChanges(i) = PercentChange
        Next i
        ' Write changes back to sheet in bulk
        ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10)).Value = Application.Transpose(quarterlyChanges)
        ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)).Value = Application.Transpose(percentChanges)
        ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)).NumberFormat = "0.00%"
        'Calculate Values for "Total Stock Volume" per ticker symbol
        '---
        
        Dim currentTicker As String
        Dim lastTicker As String
        TotalStockVolume = 0
        
        For i = 2 To lastRow
            currentTicker = ws.Cells(i, 1).Value
            StockVolume = ws.Cells(i, 7).Value
            
            If i = 2 Then
                lastTicker = currentTicker
                TotalStockVolume = StockVolume
            ElseIf currentTicker = lastTicker Then
                TotalStockVolume = TotalStockVolume + StockVolume
            Else
                ' Write the total for the previous ticker
                ws.Cells(i - 1, 12).Value = TotalStockVolume
                
                ' Reset for the new ticker
                lastTicker = currentTicker
                TotalStockVolume = StockVolume
            End If
            
            ' Write the ticker symbol
            ws.Cells(i, 9).Value = currentTicker
            
            ' Write the total for the last ticker in the list
            If i = lastRow Then
                ws.Cells(i, 12).Value = TotalStockVolume
            End If
        Next i
        '-----------------------------
        'Calculate Values for Greatest%Increase, Greatest%Decrease and GreatestTotalVolume and copy into new columns
        '--
        'Define variables
        Dim GreatestpercentIncrease As Double
        Dim GreatestpercentDecrease As Double
        Dim GreatestTotalVolume As Double
        Dim GreatestpercentIncreaseTicker As String
        Dim GreatestpercentDecreaseTicker As String
        Dim GreatestTotalVolumeTicker As String
        ' Initialize variables
        GreatestpercentIncrease = -1
        GreatestpercentDecrease = 1
        GreatestTotalVolume = 0
        ' Find greatest values
        For i = 2 To lastRow
            If IsNumeric(ws.Cells(i, 11).Value) And IsNumeric(ws.Cells(i, 12).Value) Then
                If ws.Cells(i, 11).Value > GreatestpercentIncrease Then
                    GreatestpercentIncrease = ws.Cells(i, 11).Value
                    GreatestpercentIncreaseTicker = ws.Cells(i, 9).Value
                End If
                If ws.Cells(i, 11).Value < GreatestpercentDecrease Then
                    GreatestpercentDecrease = ws.Cells(i, 11).Value
                    GreatestpercentDecreaseTicker = ws.Cells(i, 9).Value
                End If
                If ws.Cells(i, 12).Value > GreatestTotalVolume Then
                    GreatestTotalVolume = ws.Cells(i, 12).Value
                    GreatestTotalVolumeTicker = ws.Cells(i, 9).Value
                End If
            End If
        Next i
        ' Write results to the sheet
        If ws.Cells(1, 16).Value <> "Ticker" Then
            ws.Cells(1, 16).Resize(1, 2).Value = Array("Ticker", "Value")
        End If
        ws.Cells(2, 15).Resize(3, 1).Value = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
        ws.Cells(2, 16).Value = GreatestpercentIncreaseTicker
        ws.Cells(3, 16).Value = GreatestpercentDecreaseTicker
        ws.Cells(4, 16).Value = GreatestTotalVolumeTicker
        ws.Cells(2, 17).Value = GreatestpercentIncrease
        ws.Cells(3, 17).Value = GreatestpercentDecrease
        ws.Cells(4, 17).Value = GreatestTotalVolume
        'Apply Conditional Formatting to QuarterlyChange
        For i = 2 To lastRow
            If IsNumeric(ws.Cells(i, 10).Value) Then
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                End If
            End If
        Next i
    Next ws
End Sub
