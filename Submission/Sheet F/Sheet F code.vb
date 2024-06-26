Option Explicit
Sub stocks():

    Dim i As Long ' row number
    Dim cell_vol As LongLong ' contents of column G (cell volume)
    Dim vol_total As LongLong ' what is going to go in column L
    Dim ticker As String ' what is going to go in column I
    Dim ws As Worksheet
    
    Dim k As Long ' leaderboard row
    
    Dim ticker_close As Double
    Dim ticker_open As Double
    Dim price_change As Double
    Dim percent_change As Double
    
    ' asked the xpert
    Dim lastRow As Long
    For Each ws In ThisWorkbook.Worksheets
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row

    vol_total = 0
    k = 2
    
    ' Write Leaderboard Columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    
    
    
    
    
    ' assign open for first ticker
    ticker_open = ws.Cells(2, 3).Value

    For i = 2 To lastRow:
        cell_vol = ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value

        ' LOOP rows 2 to 254
        ' check if next row ticker is DIFFERENT
        ' if the same, then we only need to add to the vol_total
        ' if DIFFERENT, then we need add last row, write out to the leaderboard
        ' reset the vol_total to 0

        If (ws.Cells(i + 1, 1).Value <> ticker) Then
            ' OMG we have a different ticker, panic
            vol_total = vol_total + cell_vol
            
            ' get the closing price of the ticker
            ticker_close = ws.Cells(i, 6).Value
            price_change = ticker_close - ticker_open
            
            ' Check if open price is 0 first - to protect from Kevin the Hacker
            If (ticker_open > 0) Then
                percent_change = price_change / ticker_open
            Else
                percent_change = 0
            End If

            ws.Cells(k, 9).Value = ticker
            ws.Cells(k, 10).Value = price_change
            ws.Cells(k, 11).Value = percent_change
            ws.Cells(k, 12).Value = vol_total
            
            ' formatting
            If (price_change > 0) Then
                ws.Cells(k, 10).Interior.ColorIndex = 4 ' Green
            ElseIf (price_change < 0) Then
                ws.Cells(k, 10).Interior.ColorIndex = 3 ' Red
            Else
                ws.Cells(k, 10).Interior.ColorIndex = 2 ' White
            End If

            ' reset
            vol_total = 0
            k = k + 1
            ticker_open = Cells(i + 1, 3).Value ' look ahead to get next ticker open
        Else
            ' we just add to the total
            vol_total = vol_total + cell_vol
        End If
    Next i
    
    ' Style my leaderboard
    ws.Columns("K:K").NumberFormat = "0.00%"
    ws.Columns("I:L").AutoFit
    
    
    ws.Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastRow)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastRow)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastRow)) * 100
    
    Dim increase_number As Integer
    Dim decrease_number As Integer
    Dim volume_number As Integer
    
    
    
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastRow)), Range("K2:K" & lastRow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastRow)), Range("K2:K" & lastRow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastRow)), Range("L2:L" & lastRow), 0)
        ws.Range("P2") = Cells(increase_number + 1, 9)
        ws.Range("P3") = Cells(decrease_number + 1, 9)
        ws.Range("P4") = Cells(volume_number + 1, 9)
        Next ws
End Sub







