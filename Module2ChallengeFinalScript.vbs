Attribute VB_Name = "Module1"
Sub StockDataAnalysis():

    ' define variables
    Dim total As Double
    Dim row As Long
    Dim rowCount As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim summaryTableRow As Long
    Dim stockStartRow As Long
    Dim startValue As Long
    Dim lastTicker As String
    Dim lastExtraRow As Long
    Dim greatestIncreaseRow As Double
    Dim greatestDecreaseRow As Double
    Dim greatestTotVolRow As Double
    
    ' loop through all the stock data
    For Each ws In Worksheets
    
    ' print the summary section headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' print the aggregate section headers
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ' initialize the values
    summaryTableRow = 0
    total = 0
    quarterlyChange = 0
    stockStartRow = 2
    startValue = 2
    
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
    
    ' find last ticker to exit loop
    lastTicker = ws.Cells(rowCount, 1).Value
    
    For row = 2 To rowCount
    
    ' check for changes in ticker
    If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
    
        total = total + ws.Cells(row, 7).Value
        
        ' check if total volume is 0
        If total = 0 Then
           ' print the results in in columns I:J
            ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
            ws.Range("J" & 2 + summaryTableRow).Value = 0
            ws.Range("K" & 2 + summaryTableRow).Value = 0 & "&"
            ws.Range("L" & 2 + summaryTableRow).Value = 0
           
        Else
            ' find the first non-zero start value
            If ws.Cells(startValue, 3).Value = 0 Then
                
                For findValue = startValue To row
                    ' check if next value does not equal 0
                    If ws.Cells(findValue, 3).Value <> 0 Then
                        startValue = findValue
                        Exit For
                    End If
                Next findValue
            End If
        
            ' calculate the quarterly change
            quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(startValue, 3).Value
            
            ' calculate the percent change
            
            percentChange = quarterlyChange / ws.Cells(startValue, 3).Value
    
            ' print results in column I:L
            ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
            ws.Range("J" & 2 + summaryTableRow).Value = quarterlyChange
            ws.Range("K" & 2 + summaryTableRow).Value = percentChange
            ws.Range("L" & 2 + summaryTableRow).Value = total
            
            
            ' conditional formatting for quarterly change
            If quarterlyChange > 0 Then
                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4 ' green
            ElseIf quarterlyChange < 0 Then
                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3 ' red
            Else
                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0 ' white
            End If
                
            ' conditional formatting for percent change
            If percentChange > 0 Then
                ws.Range("K" & 2 + summaryTableRow).Interior.ColorIndex = 4 ' green
            ElseIf quarterlyChange < 0 Then
                ws.Range("K" & 2 + summaryTableRow).Interior.ColorIndex = 3 ' red
            Else
                ws.Range("K" & 2 + summaryTableRow).Interior.ColorIndex = 0 ' white
            End If
                
            ' reset the value
            total = 0
            averageChange = 0
            quarterlyChange = 0
            startValue = row + 1
            summaryTableRow = summaryTableRow + 1
        
        End If
        
    Else
        total = total + ws.Cells(row, 7).Value
    End If
    
Next row

        ' clean up and update the summary table row
    
        summaryTableRow = ws.Cells(Rows.Count, "I").End(xlUp).row
    
        lastExtraRow = ws.Cells(Rows.Count, "J").End(xlUp).row
    
        For e = summaryTableRow To lastExtraRow
    
            For Column = 9 To 12
                ws.Cells(e, Column).Value = ""
                ws.Cells(e, Column).Interior.ColorIndex = 0
            Next Column
        Next e
    
    ' find max and min
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2))
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2))
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2))
     
    ' match tickers
    greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
    greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
    greatestTotVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2)), ws.Range("L2:L" & summaryTableRow + 2), 0)
    
    ' display the tickers
    ws.Range("P2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
    ws.Range("P3").Value = ws.Cells(greatestDecreaseRow + 1, 9).Value
    ws.Range("P4").Value = ws.Cells(greatestTotVolRow + 1, 9).Value
    
    ' format the summary table
    For s = 0 To summaryTableRow
        ws.Range("J" & 2 + s).NumberFormat = "0.00%"
        ws.Range("K" & 2 + s).NumberFormat = "0.00%"
        ws.Range("L" & 2 + s).NumberFormat = "#,###"
    Next s
    
    ' format the summary aggregates
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "#,###"
    
     ' autofit columns
    ws.Columns("A:Q").AutoFit

Next ws

End Sub
