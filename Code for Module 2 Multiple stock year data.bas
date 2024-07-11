
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
        
        
        'Add a Column for the "Ticker", "Quarterly Change","Percent Change",and "Total Stock Volume"
        ws.Range("I1:L1").EntireColumn.Insert
        
        
        'Add the following words to Row 1, columns 9 to 12: "Ticker", "Quarterly Change","Percent Change",and "Total Stock Volume"
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        

        '--------------------------
        'Retrieve Values
        '--------------------------
        
        'Retrieve values in <ticker> column to new ticker column
        '----
        
        'Copy values from the source column into the target column in the same worksheet
        ws.Range("A2:A" & lastRow).Copy Destination:=ws.Range("I2:I" & lastRow)
    
        '------------------------
        'Calculate Values for "Quarterly Change" and "Percent Change"
        '---
        
        For i = 2 To lastRow
        
        'Retrieve opening price and closing price values from source columns
        OpeningPrice = ws.Cells(i, 3).Value
        ClosingPrice = ws.Cells(i, 6).Value
        
        'Calculate the quarterly change
        QuarterlyChange = ClosingPrice - OpeningPrice
        PercentChange = (QuarterlyChange / OpeningPrice) * 100
        
        'Copy calculated value into new column for quarterly change, then percent change, respectively
        ws.Cells(i, 10).Value = QuarterlyChange
        ws.Cells(i, 11).Value = PercentChange
        
        'Format PercentChange as a percentage
        ws.Cells(i, 11).NumberFormat = "0.00%"
        
        Next i
        
        
        '------------------------
        'Calculate Values for "Total Stock Volume" and copy into new column
        '---
    
        
        For i = 2 To lastRow
            TotalStockVolume = 0
            For j = 2 To i
            
        
            StockVolume = ws.Cells(j, 7).Value
            TotalStockVolume = TotalStockVolume + StockVolume
        
        Next j
    
            'Copy value into new column in each row
            ws.Cells(i, 12).Value = TotalStockVolume
        
        Next i
        
        '-----------------------------
        'Calculate Values for Greatest%Increase, Greatest%Decrease and GreatestTotalVolume and copy into new columns
        '--
        
        'Define variables
        Dim GreatestpercentIncrease As Double
        Dim GreatestpercentDecrease As Double
        Dim GreatestTotalVolume As Double
        
        'Add a Column for each variable
        ws.Range("O1:Q1").EntireColumn.Insert
        
        'Add the following words to Row 1, columns 16 and 17: "Ticker", "Value"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Add the following words to Row 2 to 4, column 15: Greatest%Increase, Greatest%Decrease and GreatestTotalVolume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Reset and Calculate values
        GreatestpercentIncrease = 0
        GreatestpercentDecrease = 0
        GreatestTotalVolume = 0
        
        For i = 2 To lastRow
        
        If ws.Cells(i, 11).Value > GreatestpercentIncrease Then
        GreatestpercentIncrease = ws.Cells(i, 11)
        ElseIf ws.Cells(i, 11).Value < GreatestpercentDecrease Then
        GreatestpercentDecrease = ws.Cells(i, 11)
        
        End If
            Next i
        
        For i = 2 To lastRow
        
        If ws.Cells(i, 12).Value > GreatestTotalVolume Then
        GreatestTotalVolume = ws.Cells(i, 12).Value
        
        End If
            Next i
            
        
        'Copy Values into the new columns
        ws.Cells(2, 17).Value = GreatestpercentIncrease
        ws.Cells(3, 17).Value = GreatestpercentDecrease
        ws.Cells(4, 17).Value = GreatestTotalVolume
        
        'Copy <ticker> associated with each value above into the new columns
        For i = 2 To lastRow
        
        If ws.Cells(i, 11).Value = GreatestpercentIncrease Then
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11).Value = GreatestpercentDecrease Then
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 12).Value = GreatestTotalVolume Then
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        
        End If
            Next i
            
        'Apply Conditional Formatting to QuarterlyChange
        For i = 2 To lastRow
        
        If ws.Cells(i, 10).Value > 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
        
        End If
            Next i
            
    Next ws
    
End Sub






