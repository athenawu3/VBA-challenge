Attribute VB_Name = "OutputStockChanges"
Sub OutputStockChanges()

    Dim ws As Worksheet

For Each ws In Worksheets

    ' Declare variables
    Dim i As Long
    Dim lastRow As Long
    Dim ticker As String
    Dim outputRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim tickerCount As Long
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalStockVol As Variant
    Dim stockSumRange As Range
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseRow As Long
    Dim greatestDecreaseRow As Long
    Dim greatestVolumeRow As Long
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Set up variables
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    outputRow = 2
    tickerCount = 0
    
    ' Count how many of the first ticker there is
    For Each cell In ws.Range("A2:A" & lastRow)
        If cell.Value = "AAF" Then
            tickerCount = tickerCount + 1
        End If
    Next cell
    
    ' Set range and give headers to new columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Columns("I:Q").AutoFit
    
    ' Loop through each row and output the following information
    For i = 3 To lastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
            ' calculations
            ticker = ws.Cells(i - 1, 1).Value
            openPrice = ws.Cells(i - tickerCount, 3).Value
            closePrice = ws.Cells(i - 1, 6).Value
            quarterlyChange = closePrice - openPrice
            percentageChange = quarterlyChange / openPrice
            Set stockSumRange = ws.Range("G" & (i - tickerCount) & ":G" & (i - 1))
            totalStockVol = Application.WorksheetFunction.Sum(stockSumRange)
            
            ' outputs
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).NumberFormat = "0.00%"
            ws.Cells(outputRow, 11).Value = percentageChange
            ws.Cells(outputRow, 12).NumberFormat = "0"
            ws.Cells(outputRow, 12).Value = totalStockVol
            
            ' conditional formatting
            If quarterlyChange < 0 Then
                ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf quarterlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Else
            End If
            If percentageChange < 0 Then
                ws.Cells(outputRow, 11).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf percentageChange > 0 Then
                    ws.Cells(outputRow, 11).Interior.Color = RGB(0, 255, 0) ' Green
                Else
            End If
        
            outputRow = outputRow + 1
            
        ElseIf i = lastRow Then
        
            ' calculations
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i + 1 - tickerCount, 3).Value
            closePrice = ws.Cells(i, 6).Value
            quarterlyChange = closePrice - openPrice
            percentageChange = quarterlyChange / openPrice
            Set stockSumRange = ws.Range("G" & (i + 1 - tickerCount) & ":G" & (i))
            totalStockVol = Application.WorksheetFunction.Sum(stockSumRange)
            
            ' outputs
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).NumberFormat = "0.00%"
            ws.Cells(outputRow, 11).Value = percentageChange
            ws.Cells(outputRow, 12).NumberFormat = "0"
            ws.Cells(outputRow, 12).Value = totalStockVol
            
            ' conditional formatting
            If quarterlyChange < 0 Then
                ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf quarterlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Else
            End If
            If percentageChange < 0 Then
                ws.Cells(outputRow, 11).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf percentageChange > 0 Then
                    ws.Cells(outputRow, 11).Interior.Color = RGB(0, 255, 0) ' Green
                Else
            End If
     
            outputRow = outputRow + 1
            
        Else
        End If
        
    Next i
    
    ' greatest % increase
    greatestIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & outputRow - 1))
    greatestIncreaseRow = Application.WorksheetFunction.Match(greatestIncrease, ws.Range("K2:K" & outputRow), 0) + 1
    greatestIncreaseTicker = ws.Cells(greatestIncreaseRow, 9).Value
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q2") = greatestIncrease
    ws.Range("P2") = greatestIncreaseTicker
    
    ' greatest % decrease
    greatestDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & outputRow - 1))
    greatestDecreaseRow = Application.WorksheetFunction.Match(greatestDecrease, ws.Range("K2:K" & outputRow), 0) + 1
    greatestDecreaseTicker = ws.Cells(greatestDecreaseRow, 9).Value
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q3") = greatestDecrease
    ws.Range("P3") = greatestDecreaseTicker
    
    ' greatest stock volume
    greatestVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & outputRow - 1))
    greatestVolumeRow = Application.WorksheetFunction.Match(greatestVolume, ws.Range("L2:L" & outputRow), 0) + 1
    greatestVolumeTicker = ws.Cells(greatestVolumeRow, 9).Value
    ws.Range("Q4").NumberFormat = "0"
    ws.Range("Q4") = greatestVolume
    ws.Range("P4") = greatestVolumeTicker
    
Next ws

End Sub
