Sub StockSummerMedium()
    'Module to sum yearly stock data
    
    'Define variables to hold the number of rows
    Dim numRows As Long
    numRows = ActiveSheet.UsedRange.Rows.Count
    
    'Define counter variables and variables to hold values during loop
    'loop counter
    Dim i As Long
    'ticker
    Dim ticker As String
    ticker = Range("A2").Value
    'total stock volume
    Dim totalVolume As LongLong
    totalVolume = 0
    'open price and closing price
    Dim openPrice As Double
    Dim closingPrice As Double
    openPrice = Cells(2, 3).Value
    closingPrice = Cells(2, 6).Value
    'change and percent change
    Dim change As Double
    Dim percentChange As Double
    change = 0#
    percentChange = 0#
    
    'row index for recording data
    Dim j As Long
    j = 2
    
    'loop through all sheets
    For Each ws In Worksheets
        'initialize variables
        ticker = Range("A2").Value
        totalVolume = 0
        openPrice = ws.Cells(2, 3).Value
        closingPrice = ws.Cells(2, 6).Value
        change = 0#
        percentChange = 0#
        j = 2
        IncTicker = ws.Range("A2").Value
        DecTicker = ws.Range("A2").Value
        VolTicker = ws.Range("A2").Value
        numRows = ws.UsedRange.Rows.Count
        
        'format headers for new output
        'stock summation
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
        'loop through all rows
        For i = 2 To numRows
            'Check if new ticker
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1)) Then
                'record ticker
                ws.Cells(j, 9).Value = ws.Cells(i, 1)
                'calulcate and record change
                closingPrice = ws.Cells(i, 6).Value
                If (openPrice = 0 And closingPrice = 0) Then
                    change = 0#
                    percentChange = 0#
                Else
                    change = closingPrice - openPrice
                    percentChange = change / openPrice
                End If
                ws.Cells(j, 10).Value = change
                ws.Cells(j, 11).Value = percentChange
                'format change columns
                If (change > 0#) Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                ElseIf (change < 0#) Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
                ws.Cells(j, 11).NumberFormat = "0.00%"
                'record total volume
                ws.Cells(j, 12).Value = totalVolume + ws.Cells(i, 7).Value
                'increment j
                j = j + 1
                'change ticker
                ticker = ws.Cells(i + 1, 1).Value
                'reset totalVolume
                totalVolume = 0
                'get beginning open price
                openPrice = ws.Cells(i + 1, 3).Value
                'make sure open price is not zero
                'if so, search for a non-zero value (assuming we are not at end of list)
                If ((openPrice = 0) And (ticker <> "")) Then
                    'we know the i + 1 row has zero in the open price column
                    Dim index As Long
                    index = i + 2
                    'search for openPrice within same ticker
                    Do While (openPrice = 0 And ws.Cells(index, 1) = ticker)
                        openPrice = ws.Cells(index, 3).Value
                        index = index + 1
                    Loop
                End If
            Else
                'Increase totalVolume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ws.Columns("A:Q").AutoFit
    Next ws
    
    


End Sub