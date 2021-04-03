' VBA Homework:VBA of Wall Street
Sub test():

    ' Loop Through Worksheets
    For Each ws In Worksheets

        ' Column Names
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    

        ' Declare Variables
        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalTicker As Double
        TotalTicker = 0
        Dim FinalRow As Long
        FinalRow = 2
        Dim YearOpen As Double
        Dim YearClose As Double
        Dim YearChange As Double
        Dim PreviousAmount As Long
        PreviousAmount = 2
        Dim PercentChange As Double
    
        ' Determine Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            ' Add To Ticker Total
            TotalTicker = TotalTicker + ws.Cells(i, 7).Value
            ' Check If Same Ticker Name
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                ' Set Ticker Name
                TickerName = ws.Cells(i, 1).Value
                ' Print The Ticker Name
                ws.Range("I" & FinalRow).Value = TickerName
                ' Print The Ticker Total
                ws.Range("L" & FinalRow).Value = TotalTicker
                ' Reset Ticker Total
                TotalTicker = 0

                ' Set Open, Close and Year Change
                YearOpen = ws.Range("C" & PreviousAmount)
                YearClose = ws.Range("F" & i)
                YearChange = YearClose - YearOpen
                ws.Range("J" & FinalRow).Value = YearChange

                ' Percent Change Formula
                If YearOpen = 0 Then
                    PercentChange = 0
                Else
                    YearOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearChange / YearOpen
                End If
                ' Format Percentage
                ws.Range("K" & FinalRow).NumberFormat = "0.00%"
                ws.Range("K" & FinalRow).Value = PercentChange
                ' Format Green or Red 
                If ws.Range("J" & FinalRow).Value >= 0 Then
                    ws.Range("J" & FinalRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                ' Add One To Final Row
                FinalRow = FinalRow + 1
                PreviousAmount = i + 1
                End If
            Next i

            
        ' Format Table Columns To Auto Fit
        ws.Columns("I:L").AutoFit

    Next ws

End Sub
