Sub stockChanges():

For Each ws In Worksheets

    ' name worksheet
    
    ws.Range("I1") = "Ticker Symbol"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Volume"
    'Set an initial variables for ticker, yearly change, table row
    Dim ticker As String
    Dim yearlyChange As Double
    
    'keep track of each ticker symbol and info
    Dim Summary_Table_Row As Long
        Summary_Table_Row = 2

    'keep track of ticker prices
    Dim ticker_first_open As Double
    Dim ticker_last_closed As Double
    
    ticker_first_open = Range("C2")

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' set percent change
    Dim percentChange As Double
    ' percentChange = 0
    
    'set total stock volume
    Dim totalStockVolume As Double
    ' totalStockVolume = 0

    'Loop through all ticker changes
        For i = 2 To LastRow
            
        'track stock voulme
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            
        ' Check if we are still within the same ticker symbol, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                If totalStockVolume = 0 Then
                    ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & Summary_Table_Row).Value = ticker
                    ws.Range("J" & Summary_Table_Row).Value = 0
                    ws.Range("K" & Summary_Table_Row).Value = 0
                    ws.Range("L" & Summary_Table_Row).Value = 0
                    
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                 
                'Set Next Open Price
                 ticker_first_open = ws.Cells(i + 1, 3).Value
                 
                    Else
                    
                ' Set the Ticker symbol
                ticker = ws.Cells(i, 1).Value
            
                ' Print the ticker symbol in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = ticker
                
                'select last closing price
                ticker_last_closed = ws.Cells(i, 6).Value
                        
                ' calculate yearly change
                yearlyChange = ticker_last_closed - ticker_first_open

                ' Print the yearlychange to the Summary Table and format
                ws.Range("J" & Summary_Table_Row).Value = yearlyChange
                
                If yearlyChange < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
                
                ' calculate percent change
                If ticker_first_open = 0 Then
                
                Else
                    percentChange = CDbl(yearlyChange / ticker_first_open)
                    
                End If
                
               ' Print the yearlychange to the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Format(percentChange, "Percent")
                          
                ' Print the yearlychange to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Format(totalStockVolume, "General Number")
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the yearlyChange
                yearlyChange = 0

                ' Reset the total volume
                totalStockVolume = 0
                 
                 'Set Next Open Price
                 ticker_first_open = ws.Cells(i + 1, 3).Value
     
                End If
                
            End If
            
        Next i
    
Next ws

End Sub

