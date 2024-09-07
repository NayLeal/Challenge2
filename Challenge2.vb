Sub StockAnalysis()

    ' Variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim quarterlyOpen As Double
    Dim quarterlyClose As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim summaryRow As Integer
    
    ' Initialize summary row
    summaryRow = 2

    ' Loop through all worksheets (quarters)
    For Each ws In ThisWorkbook.Worksheets

        ' Set initial variables for each worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        totalVolume = 0
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        ' Loop through each row in the worksheet
        For i = 2 To lastRow

            ' Check if we are still on the same ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' New ticker found

                ' Record the opening price at the start of the ticker
                quarterlyOpen = ws.Cells(i, 3).Value
                
                ' Add the volume for the current ticker
                totalVolume = ws.Cells(i, 7).Value
                
            Else
                ' Same ticker, add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If

            ' If the next row has a different ticker, record the closing price
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastRow Then
                ' Closing price at the end of the ticker
                quarterlyClose = ws.Cells(i, 6).Value
                
                ' Calculate the quarterly change
                quarterlyChange = quarterlyClose - quarterlyOpen
                
                ' Calculate the percentage change
                If quarterlyOpen <> 0 Then
                    percentageChange = (quarterlyChange / quarterlyOpen) * 100
                Else
                    percentageChange = 0
                End If
                
                ' Output the ticker, quarterly change, percentage change, and total volume
                ws.Cells(summaryRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = percentageChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Determine greatest increase, decrease, and volume
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    greatestIncreaseTicker = ws.Cells(i, 1).Value
                End If

                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    greatestDecreaseTicker = ws.Cells(i, 1).Value
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ws.Cells(i, 1).Value
                End If
                
                ' Move to the next row in summary
                summaryRow = summaryRow + 1
                
            End If

        Next i

    Next ws

    ' Output the greatest values
    With ThisWorkbook.Sheets(1)
        .Cells(1, 14).Value = "Greatest % Increase"
        .Cells(2, 14).Value = greatestIncreaseTicker
        .Cells(2, 15).Value = greatestIncrease

        .Cells(4, 14).Value = "Greatest % Decrease"
        .Cells(5, 14).Value = greatestDecreaseTicker
        .Cells(5, 15).Value = greatestDecrease

        .Cells(7, 14).Value = "Greatest Total Volume"
        .Cells(8, 14).Value = greatestVolumeTicker
        .Cells(8, 15).Value = greatestVolume
    End With

    MsgBox "Stock Analysis Complete"

End Sub
