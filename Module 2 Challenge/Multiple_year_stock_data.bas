Attribute VB_Name = "Module1"
Sub stock_analysis()
    Dim i As Long
    Dim lastRow As Long
    Dim endRow As Long
    Dim startRow As Long
    Dim tickerRow As Long
    Dim quarters As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double


    For Each quarters In Worksheets
        With quarters
            ' Headers
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Quarterly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
            .Cells(1, 17).Value = "Ticker"
            .Cells(1, 18).Value = "Value"
            .Cells(2, 16).Value = "Greatest % Increase"
            .Cells(3, 16).Value = "Greatest % Decrease"
            .Cells(4, 16).Value = "Greatest Total Volume"
            

            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

            greatestIncrease = -1 ' Initialize to very low number
            greatestDecrease = 1 ' Initialize to very high number
            greatestVolume = 0
            tickerRow = 2

            startRow = 2 ' Initialize startRow before the loop
            openPrice = .Cells(startRow, 3).Value ' Initialize openPrice
            

                       

            ' Data for ticker, quarterly change, percent change, and total stock value
            For i = 2 To lastRow
                If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Or i = lastRow Then
                    endRow = i
                    closePrice = .Cells(endRow, 6).Value
                    ticker = .Cells(i, 1).Value

                    .Cells(tickerRow, 9).Value = ticker
                    totalVolume = WorksheetFunction.Sum(.Range(.Cells(startRow, 7), .Cells(endRow, 7)))
                    .Cells(tickerRow, 12).Value = totalVolume
                    quarterlyChange = closePrice - openPrice

                    .Cells(tickerRow, 10).Value = quarterlyChange
                    'Avoid division by zero error
                    If openPrice <> 0 Then
                        percentChange = quarterlyChange / openPrice
                        
                    Else
                        percentChange = 0 'or any other appropriate default value
                    End If
                    .Cells(tickerRow, 11).Value = percentChange

                    ' Update greatest values
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        greatestIncreaseTicker = ticker
                    End If
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        greatestDecreaseTicker = ticker
                    End If
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = ticker
                    End If

                    tickerRow = tickerRow + 1
                    startRow = i + 1
                    openPrice = .Cells(startRow, 3).Value
                End If
            Next i
            
            ' Output greatest values
            .Cells(2, 17).Value = greatestIncreaseTicker
            .Cells(2, 18).Value = greatestIncrease
            .Cells(3, 17).Value = greatestDecreaseTicker
            .Cells(3, 18).Value = greatestDecrease
            .Cells(4, 17).Value = greatestVolumeTicker
            .Cells(4, 18).Value = greatestVolume
        End With ' End With block for quarters
    Next quarters
End Sub

Sub Color()

    Dim i As Long
    Dim lastRow As Long
    Dim quarters As Worksheet

    For Each quarters In Worksheets
        With quarters
            lastRow = .Cells(.Rows.Count, 10).End(xlUp).Row

            
            For i = 2 To lastRow
                .Cells(i, 11).NumberFormat = "0.00%"
                If .Cells(i, 10) > 0 Then
                    .Cells(i, 10).Interior.ColorIndex = 4  ' Green
                ElseIf .Cells(i, 10) < 0 Then
                    .Cells(i, 10).Interior.ColorIndex = 3  ' Red
                End If
            Next i

            
             .Range("R2:R3").NumberFormat = "0.00%"
        End With
    Next quarters

End Sub
