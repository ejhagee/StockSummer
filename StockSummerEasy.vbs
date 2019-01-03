Sub StockSummerEasy()
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
    
    'row index for recording data
    Dim j As Long
    j = 2
    
    'loop through all sheets
    For Each ws In Worksheets
        'initialize variables
        ticker = Range("A2").Value
        totalVolume = 0
        j = 2
        numRows = ws.UsedRange.Rows.Count
        
        'format headers for new output
        'stock summation
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
    
        'loop through all rows
        For i = 2 To numRows
            'Check if new ticker
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1)) Then
                'record ticker
                ws.Cells(j, 9).Value = ws.Cells(i, 1)
                ws.Cells(j, 10).Value = totalVolume + ws.Cells(i, 7).Value
                'increment j
                j = j + 1
                'change ticker
                ticker = ws.Cells(i + 1, 1).Value
                'reset totalVolume
                totalVolume = 0
            Else
                'Increase totalVolume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ws.Columns("A:Q").AutoFit
    Next ws
    
    


End Sub