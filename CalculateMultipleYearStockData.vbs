' CalculateMultipleYearStockData
' This script script loops through all the stocks for one year and outputs information in the same sheet
Sub CalculateMultipleYearStockData()

    ' Declare variables
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim StartRow As Long
    Dim EndRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim ResultRow As Long
    Dim GreatestPercentIncrease As Double 
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double

    ' Loop through each worksheet in the workbook
    For Each ws In Worksheets
       
        ' Set Max and Min as 0
        GreatestPercentIncrease = 0
        GreatestPercentDecrease = 0
        GreatestTotalVolume = 0

        ' Find the last row of each worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Name of column headers which will store the results
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' set start of the row and start of the result row
        StartRow = 2
        ResultRow = 2
       
        ' Loop through all rows of the worksheet
        For i = 2 To LastRow
            ' Check if we are still within the same stock
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                EndRow = i
               
                ' Find the value of the ticker, opening price, and closing price
                Ticker = ws.Cells(StartRow, 1).Value
                OpeningPrice = ws.Cells(StartRow, 3).Value
                ClosingPrice = ws.Cells(EndRow, 6).Value
               
                ' Calculate the yearly change and percentage change
                YearlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentChange = (YearlyChange / OpeningPrice) * 100
                    PercentChange = Round(PercentChange, 2)
                Else
                    PercentChange = 0
                End If
               
                ' Compute the total stock volume for the year
                TotalStockVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(StartRow, 7), ws.Cells(EndRow, 7)))
               
                ' Populate the results in the new columns
                ws.Cells(ResultRow, 9).Value = Ticker
                ws.Cells(ResultRow, 10).Value = YearlyChange
                'Colour the cell green if value is greater than 0 else color cell red
                If ws.Cells(ResultRow, 10).Value >= 0 Then
                    ws.Cells(ResultRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(ResultRow, 10).Interior.ColorIndex = 3
                End If

                ws.Cells(ResultRow, 12).Value = TotalStockVolume

               
                ' Check greastest precentage increase and decrease
                If PercentChange > GreatestPercentIncrease Then
                    GreatestPercentIncrease = PercentChange
                    TickerNameMaxIncrease = Ticker
                End If
               
                If PercentChange < GreatestPercentDecrease Then
                    GreatestPercentDecrease = PercentChange
                    TickerNameMinDecrease = Ticker
                End If
                ws.Cells(ResultRow, 11).Value = PercentChange & "%"

                If TotalStockVolume > GreatestTotalVolume Then
                    GreatestTotalVolume = TotalStockVolume
                    TickerMaxVolume = Ticker
                End If



                ' Adjust counters for the next stock
                StartRow = i + 1
                ResultRow = ResultRow + 1
            End If

        Next i

        ' Display the greatest increase and decrease values. Also the ticker name
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = TickerNameMaxIncrease
        ws.Cells(3, 16).Value = TickerNameMinDecrease
        ws.Cells(4, 16).Value = TickerMaxVolume
        ws.Cells(2, 17).Value = GreatestPercentIncrease & "%"
        ws.Cells(3, 17).Value = GreatestPercentDecrease & "%"
        ws.Cells(4, 17).Value = GreatestTotalVolume



    Next ws

    ' message box displaying complete message
    MsgBox ("Script Complete")

End Sub
