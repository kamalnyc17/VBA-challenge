Sub MultiSheetStockTickers()
    ' setting up error trapping
    On Error Resume Next
    
    ' initialization section
    Dim strTicker, strTickerPercentIncrease, strTickerPercentDecrease, strTickerTotalVolume  As String
    Dim intOpeningPrice, intClosingPrice, intYearlyChange, intPercentChange, intTotalStockVolume As Double
    Dim intGreatestPercentIncrease, intGreatestPercentDecrease, intGreatestTotalVolume As Double
    Dim intLastRow, intResultsRow As Integer
    
    ' process each sheet one at a time
    For Each ws In Worksheets
        
        ' assign initial value for each sheet
        intLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        intResultsRow = 2
        intSummaryRow = 2
        intGreatestPercentIncrease = 0
        intGreatestPercentDecrease = 0
        intGreatestTotalVolume = 0
        
        ' main loop starts here
        For i = 2 To intLastRow
            
            ' create headers at the beginning
            If (i = 2) Then
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
                ws.Cells(1, 15).Value = "Ticker"
                ws.Cells(1, 16).Value = "Value"
                ws.Range("I1:L1").Columns.AutoFit
                ws.Range("J2:J" & intLastRow).NumberFormat = "0.00"
                ws.Range("K2:K" & intLastRow).NumberFormat = "0.00%"
                ws.Range("L2:L" & intLastRow).NumberFormat = "00000"
                ws.Range("I1:I" & intLastRow).HorizontalAlignment = xlCenter
            End If
            
            ' initialize variables for current stock
            strTicker = ws.Cells(i, 1).Value
            intOpeningPrice = ws.Cells(i, 3).Value
            intYearlyChange = 0
            intPercentChange = 0
            intTotalStockVolume = 0
            
            ' calculate totals of the current stock
            While (ws.Cells(i, 1).Value = strTicker)
                intClosingPrice = ws.Cells(i, 6).Value
                intTotalStockVolume = intTotalStockVolume + ws.Cells(i, 7).Value
                
                ' incrementing the row number inside the sub-loop
                i = i + 1
            Wend
            
            ' calculate current stock's data
            intYearlyChange = (intClosingPrice - intOpeningPrice)
            If (intOpeningPrice <> 0) Then
                intPercentChange = (intClosingPrice - intOpeningPrice) / intOpeningPrice
            Else
                intPercentChange = 100
            End If
            
            ' update table with the current stock information
            ws.Range("I" & intResultsRow) = strTicker
            ws.Range("J" & intResultsRow) = intYearlyChange
            If (intYearlyChange > 0) Then
                ws.Range("J" & intResultsRow).Interior.ColorIndex = 4
            Else
                ws.Range("J" & intResultsRow).Interior.ColorIndex = 3
            End If
            ws.Range("K" & intResultsRow) = intPercentChange
            ws.Range("L" & intResultsRow) = intTotalStockVolume
            
            ' increase result row's value by one for the next stock symbol
            intResultsRow = intResultsRow + 1
            
            ' calculating summary section (challenge)
            If (intPercentChange > intGreatestPercentIncrease) Then
                strTickerPercentIncrease = strTicker
                intGreatestPercentIncrease = intPercentChange
            End If
            If (intPercentChange < intGreatestPercentDecrease) Then
                strTickerPercentDecrease = strTicker
                intGreatestPercentDecrease = intPercentChange
            End If
            If (intTotalStockVolume > intGreatestTotalVolume) Then
                strTickerTotalVolume = strTicker
                intGreatestTotalVolume = intTotalStockVolume
            End If
            
        Next i
        
        ' challenges - summary section
        ws.Cells(2, 14).Value = "Greatest % increase"
        ws.Cells(2, 15).Value = strTickerPercentIncrease
        ws.Cells(2, 16).Value = intGreatestPercentIncrease
        ws.Cells(3, 14).Value = "Greatest % decrease"
        ws.Cells(3, 15).Value = strTickerPercentDecrease
        ws.Cells(3, 16).Value = intGreatestPercentDecrease
        ws.Cells(4, 14).Value = "Greatest total volume"
        ws.Cells(4, 15).Value = strTickerTotalVolume
        ws.Cells(4, 16).Value = intGreatestTotalVolume
        ws.Range("N2:N4").Columns.AutoFit
        ws.Range("P2:P3").NumberFormat = "0.00%"
        ws.Range("P4").NumberFormat = "00000"
        ws.Range("O1:O4").HorizontalAlignment = xlCenter
        ws.Range("P1").HorizontalAlignment = xlRight
        
    Next ws
    
    ' notifying the user about job completion
    MsgBox "Congratulations! Process Completed Successfully", vbOKOnly, "Job Status"
    
End Sub

