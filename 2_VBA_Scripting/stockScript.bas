Attribute VB_Name = "Module1"
Sub Stocks()


For Each ws In Worksheets
'To work on just the current worksheet:
'Dim ws As Worksheet
'Set ws = ActiveSheet

ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Total Stock Volume"
'Long data type is able to store longer values and Double is able to store even more
Dim numRows As Long
numRows = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

Dim ticker As String

Dim volume As Double
Dim cellValue As Double

ticker = ws.Cells(2, 1).Value
For i = 2 To numRows
    
    If ws.Cells(i, 1).Value = ticker Then
        cellValue = ws.Cells(i, 7).Value
        volume = volume + cellValue
        ticker = ws.Cells(i, 1).Value
    
    Else
        'Print out current values for ticker and volume
        Dim sourceCol As Integer, rowCount As Integer, currentRow As Integer
        Dim currentRowValue As String

        'print ticker
        sourceCol = 10
        rowCount = ws.Cells(ws.Rows.Count, sourceCol).End(xlUp).Row
        rowCount = rowCount + 1
       
        'For the current row to the last row in the source Column
        For currentRow = 2 To rowCount
        'captures the value in the currentRow
        currentRowValue = ws.Cells(currentRow, sourceCol).Value
            'if that value is empty or "" then put in the ticker
            If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                ws.Cells(currentRow, sourceCol).Value = ticker
            End If
        Next currentRow
            
        'print volume
        sourceCol = 11
        rowCount = ws.Cells(ws.Rows.Count, sourceCol).End(xlUp).Row
        rowCount = rowCount + 1
        
        For currentRow = 2 To rowCount
        currentRowValue = ws.Cells(currentRow, sourceCol).Value
            If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                ws.Cells(currentRow, sourceCol).Value = volume
            End If
        Next currentRow
        
        'Reset volume and ticker
        ticker = ws.Cells(i, 1).Value
        volume = 0 + ws.Cells(i, 7).Value
    End If
        'ticker = ws.Cells(i, 1).Value
Next i
    
'print last ticker and volume
    sourceCol = 10
        rowCount = ws.Cells(ws.Rows.Count, sourceCol).End(xlUp).Row
        rowCount = rowCount + 1
        
        For currentRow = 2 To rowCount
       
        currentRowValue = ws.Cells(currentRow, sourceCol).Value
            If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                ws.Cells(currentRow, sourceCol).Value = ticker
            End If
        Next currentRow
    sourceCol = 11
        rowCount = ws.Cells(ws.Rows.Count, sourceCol).End(xlUp).Row
        rowCount = rowCount + 1
        
        For currentRow = 2 To rowCount
        currentRowValue = ws.Cells(currentRow, sourceCol).Value
            If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                ws.Cells(currentRow, sourceCol).Value = volume
            End If
        Next currentRow
        
    ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percent Change"

Dim ticker2 As String

'Find the last row
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

ticker2 = ws.Cells(1, 1).Value
Dim percentIncrease As Double
    
For i = 2 To lastRow
    'if the date of Jan 1 then capture the value of the open price on that date
    'lastFour = Right(ws.Cells(i, 2).Value, 4)
    If ws.Cells(i, 1).Value <> ticker2 And ws.Cells(i, 3).Value <> 0 Then
        firstOpen = ws.Cells(i, 3).Value
        ticker2 = ws.Cells(i, 1).Value
        
    'if the date is Jan 31st then capture the value of the end price on that date
    ElseIf ws.Cells(i + 1, 1).Value <> ticker2 And ws.Cells(i, 6).Value <> 0 Then
        lastClose = ws.Cells(i, 6).Value
        ticker2 = ws.Cells(i, 1).Value
        'calculate the yearly change
        yearlyChange = lastClose - firstOpen
        percentIncrease = yearlyChange / firstOpen
    
        'find the corresponding ticker in the pivot table and print the yearly change
        lastTicker = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
        For j = 2 To lastTicker
            If ws.Cells(j, 10).Value = ticker2 Then
                ws.Cells(j, 12).Value = yearlyChange
                ws.Cells(j, 13).Value = percentIncrease
                'formatting
                If yearlyChange >= 0 Then
                    ws.Cells(j, 12).Interior.ColorIndex = 43
                Else
                    ws.Cells(j, 12).Interior.ColorIndex = 3
                End If
                ws.Cells(j, 13).NumberFormat = "0.00%"
            End If
        Next j
    Else
    End If
Next i
    
Next ws




Dim tickerPercentIncrease As String
Dim tickerPercentDecrease As String
Dim tickerVolume As String

Dim greatestPercentIncrease As Double
Dim greatestPercentDecrease As Double
Dim greatestVolume As Double



'Find greatest percent increase
For Each ws In Worksheets

greatestPercentIncrease = 0
greatestPercentDecrease = 0
greatestVolume = 0

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Dim lastTicker2 As String
lastTicker2 = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row

    For i = 2 To lastTicker2
        If ws.Cells(i, 13).Value >= greatestPercentIncrease Then
            tickerPercentIncrease = ws.Cells(i, 10).Value
            greatestPercentIncrease = ws.Cells(i, 13).Value
        End If
    Next i


    For i = 2 To lastTicker2
        If ws.Cells(i, 13).Value <= greatestPercentDecrease Then
            tickerPercentDecrease = ws.Cells(i, 10).Value
            greatestPercentDecrease = ws.Cells(i, 13).Value
        End If
    Next i

    For i = 2 To lastTicker2
        If ws.Cells(i, 11).Value >= greatestVolume Then
            tickerVolume = ws.Cells(i, 10).Value
            greatestVolume = ws.Cells(i, 11).Value
        End If
    Next i
    
'Does this go inside or outsed the Next i

ws.Range("P4").Value = tickerVolume
ws.Range("Q4").Value = greatestVolume
ws.Range("P3").Value = tickerPercentDecrease
ws.Range("Q3").Value = greatestPercentDecrease
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("P2").Value = tickerPercentIncrease
ws.Range("Q2").Value = greatestPercentIncrease
ws.Range("Q2").NumberFormat = "0.00%"

Next ws

End Sub
