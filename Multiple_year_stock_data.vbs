' Stock Market Analyst - The VBA of Wall Street
Sub TickerAnalysis():

    ' Loop through the worksheets
    For Each ws In Worksheets

        ' Setup the column headers and labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Define variables
        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalStockVolume As Double
        TotalStockVolume = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
        PreviousAmount = 2
        Dim PercentChange As Double
        Dim LastRowValue As Long
        

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            ' Add to the total stock volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            ' Has the ticker symbol changed?
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set Ticker Name
                TickerName = ws.Cells(i, 1).Value
                ' Print the ticker name to "ticker" in the summary table
                ws.Range("I" & SummaryTableRow).Value = TickerName
                ' Print the ticker total amount to "total stock volume" in summary
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                ' Reset the counter "total stock volume"
                TotalStockVolume = 0

                ' Set Yearly Open, Yearly Close and Yearly Change Name
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

                ' Determine the percent change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                ' Apply formatting adding percentrage & decimal places
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                ' Apply conditional formatting using Green for a positive change and Red for negative changes
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                ' Increase the summary table row by one
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            Next i


        ws.Columns("I:L").AutoFit

    Next ws

End Sub

