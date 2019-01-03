Sub StockSummerHard()
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
    
    'greatest percent increase/decrease and greatest total vol
    Dim greatestInc As Double
    Dim greatestDec As Double
    Dim greatestVol As LongLong
    Dim IncTicker As String
    Dim DecTicker As String
    Dim VolTicker As String
    IncTicker = Range("A2").Value
    DecTicker = Range("A2").Value
    VolTicker = Range("A2").Value
    greatestInc = 0#
    greatestDec = 0#
    greatestVol = 0
    
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
        greatestInc = 0#
        greatestDec = 0#
        greatestVol = 0
        numRows = ws.UsedRange.Rows.Count
        
        'format headers for new output
        'stock summation
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'greatest instances
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
    
    
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
    
        'Counter variable for second loop to find greatest values
        Dim k As Long
        k = 0
        
        'start variables
        greatestInc = 0#
        greatestDec = 0#
        greatestVol = 0#
        IncTicker = ""
        DecTicker = ""
        VolTicker = ""
        'loop through all recorded stocks (to j - 1)
        For k = 2 To (j - 1)
            'get percent change and totalVolume
            ticker = ws.Cells(k, 9).Value
            percentChange = ws.Cells(k, 11).Value
            totalVolume = ws.Cells(k, 12).Value
            'see if greatest increase or decrease
            If (percentChange > 0# And percentChange > greatestInc) Then
                greatestInc = percentChange
                IncTicker = ticker
            ElseIf (percentChange < 0# And percentChange < greatestDec) Then
                greatestDec = percentChange
                DecTicker = ticker
            End If
            'see if total volume is greatest total volume
            If (totalVolume > greatestVol) Then
                greatestVol = totalVolume
                VolTicker = ticker
            End If
            
        Next k
        
        'record greatest amounts
        ws.Cells(2, 16).Value = IncTicker
        ws.Cells(2, 17).Value = greatestInc
        ws.Cells(3, 16).Value = DecTicker
        ws.Cells(3, 17).Value = greatestDec
        ws.Cells(4, 16).Value = VolTicker
        ws.Cells(4, 17).Value = greatestVol
        
        ws.Columns("A:Q").AutoFit
    Next ws
    
    


End Sub