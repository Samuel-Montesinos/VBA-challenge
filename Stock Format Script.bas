Attribute VB_Name = "Module1"
Sub CalculateStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim outputRow As Long
    Dim ticker As String
    Dim dateString As String
    Dim dateValue As Date
    Dim openPrice As Double
    Dim highPrice As Double
    Dim lowPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim yearStartPrice As Double
    Dim yearEndPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    For Each ws In ThisWorkbook.Worksheets
        ws.Range("I:L").ClearContents
        ws.Range("O:P").ClearContents
        ws.Range("Q:Q").NumberFormat = "General"

        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ticker = ws.Range("A2").Value
        yearStartPrice = ws.Range("C2").Value

        outputRow = 2

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        For i = 2 To lastRow
            ticker = ws.Range("A" & i).Value
            dateString = ws.Range("B" & i).Value
            dateValue = CDate(Left(dateString, 4) & "/" & Mid(dateString, 5, 2) & "/" & Right(dateString, 2))
            openPrice = ws.Range("C" & i).Value
            highPrice = ws.Range("D" & i).Value
            lowPrice = ws.Range("E" & i).Value
            closePrice = ws.Range("F" & i).Value
            volume = ws.Range("G" & i).Value

            If ticker <> ws.Range("A" & (i - 1)).Value Then
                yearStartPrice = openPrice
                totalVolume = 0
            End If

            If Month(dateValue) = 12 And Day(dateValue) = 31 Then
                yearEndPrice = closePrice
                yearlyChange = yearEndPrice - yearStartPrice
                percentageChange = yearlyChange / yearStartPrice
                totalVolume = totalVolume + volume

                ws.Range("I" & outputRow).Value = ticker
                ws.Range("J" & outputRow).Value = yearlyChange
                ws.Range("K" & outputRow).Value = percentageChange
                ws.Range("L" & outputRow).Value = totalVolume

                ws.Range("K" & outputRow).NumberFormat = "0.00%"

                If yearlyChange < 0 Then
                    ws.Range("J" & outputRow).Interior.Color = RGB(255, 0, 0)
                ElseIf yearlyChange > 0 Then
                    ws.Range("J" & outputRow).Interior.Color = RGB(0, 255, 0)
                End If

                If percentageChange < 0 Then
                    ws.Range("K" & outputRow).Interior.Color = RGB(255, 0, 0)
                ElseIf percentageChange > 0 Then
                    ws.Range("K" & outputRow).Interior.Color = RGB(0, 255, 0)
                End If

                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    greatestIncreaseTicker = ticker
                End If

                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    greatestDecreaseTicker = ticker
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If

                outputRow = outputRow + 1
            Else
                totalVolume = totalVolume + volume
            End If
        Next i

        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("P2").Value = greatestIncreaseTicker
        ws.Range("P3").Value = greatestDecreaseTicker
        ws.Range("P4").Value = greatestVolumeTicker
        ws.Range("Q1").Value = "Percent Change"
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("Q4").Value = greatestVolume
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    Next ws
End Sub


