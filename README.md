# VBA-challenge
Sub GREAT SUCCESS()
    Dim year As Integer
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim openPrice As Double, closePrice As Double
    Dim currentTicker As String, previousTicker As String
    Dim outputRow As Long ' Track the row for output
    
    ' Loop through each year
    For year = 2018 To 2020
        ' Set the worksheet for the current year
        Set ws = Worksheets(CStr(year))
        
        ' Reset variables for each sheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        previousTicker = ws.Cells(2, "A").Value
        openPrice = ws.Cells(2, "C").Value
        outputRow = 2
        
        ' Initialize other variables for each sheet
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim tickerGreatestIncrease As String
        Dim tickerGreatestDecrease As String
        Dim tickerGreatestVolume As String
        
        ' Initialize with first values
        greatestIncrease = ws.Cells(2, "J").Value
        greatestDecrease = ws.Cells(2, "J").Value
        greatestVolume = ws.Cells(2, "K").Value
        tickerGreatestIncrease = ws.Cells(2, "A").Value
        tickerGreatestDecrease = ws.Cells(2, "A").Value
        tickerGreatestVolume = ws.Cells(2, "A").Value
        
        For i = 2 To lastRow
            currentTicker = ws.Cells(i, "A").Value
            
            ' Check if still on the same ticker or it's the last row
            If currentTicker <> previousTicker Or i = lastRow Then
                ' If it's a new ticker or last row, record the closePrice of the previous ticker
                ' and calculate the yearly change for the previous ticker
                If i = lastRow And currentTicker = previousTicker Then
                    ' Special case for the last row if it's the same ticker
                    closePrice = ws.Cells(i, "F").Value
                Else
                    closePrice = ws.Cells(i - 1, "F").Value
                End If
                
                ' Calculate yearly change for the previous ticker and write to column I
                ws.Cells(outputRow, "I").Value = closePrice - openPrice
                
                ' Calculate yearly percent change for the previous ticker and write to column J
                If openPrice <> 0 Then
                    ws.Cells(outputRow, "J").Value = ((closePrice - openPrice) / openPrice) * 100
                Else
                    ws.Cells(outputRow, "J").Value = 0 ' Avoid division by zero
                End If
                
                ' Sum trading volume for the previous ticker and write to column K
                ws.Cells(outputRow, "K").Value = Application.WorksheetFunction.SumIf(ws.Range("A:A"), previousTicker, ws.Range("G:G"))
                
                ' Update greatestIncrease, greatestDecrease, greatestVolume, and corresponding tickers
                If greatestIncrease < ws.Cells(outputRow, "J").Value Then
                    greatestIncrease = ws.Cells(outputRow, "J").Value
                    tickerGreatestIncrease = previousTicker
                End If
                
                If greatestDecrease > ws.Cells(outputRow, "J").Value Then
                    greatestDecrease = ws.Cells(outputRow, "J").Value
                    tickerGreatestDecrease = previousTicker
                End If
                
                If greatestVolume < ws.Cells(outputRow, "K").Value Then
                    greatestVolume = ws.Cells(outputRow, "K").Value
                    tickerGreatestVolume = previousTicker
                End If
                
                ' Move to the next row for output
                outputRow = outputRow + 1
                
                ' Reset openPrice for the new ticker
                openPrice = ws.Cells(i, "C").Value
            End If
            
            ' Update previousTicker for the next iteration
            previousTicker = currentTicker
        Next i
        
        ' Populate the results in Cells N2, N3, N4, O2, O3, O4 for the current sheet
        ws.Cells(2, "N").Value = tickerGreatestIncrease
        ws.Cells(3, "N").Value = tickerGreatestDecrease
        ws.Cells(4, "N").Value = tickerGreatestVolume
        ws.Cells(2, "O").Value = greatestIncrease
        ws.Cells(3, "O").Value = greatestDecrease
        ws.Cells(4, "O").Value = greatestVolume
    Next year
End Sub
