VBA Code for automated analysis of stock data

Sub WallStreet()
    
    'Loop through each worksheet
    For Each ws In Worksheets
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Set an initial variable for holding the ticker
    Dim Ticker As String
    
    'Set an initial variable for holding the total amount of volume
    Dim VolumeTotal As Double
    VolumeTotal = 0
    
    'Keep Track of the location for each ticker in the summary table
    Summary_Table_Row = 2
    
        'GRAB TOTAL AMOUNT OF VOLUME EACH STOCK HAD OVER THE YEAR

            'Set Header Name
            ws.Range("J1").Value = "Total Stock Volume"
        
            'Sum the volume of all the vol column values that have the same ticker
            
            'Loop through all cells
            For i = 2 To LastRow
                'Check if we are still within the same ticker, if we are not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the Ticker name
                Ticker = ws.Cells(i, 1).Value
                
                'Add to the VolumeTotal
                VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                
                'Print the Ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                'Print the VolumeTotal in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = VolumeTotal
                
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset the VolumeTotal
                VolumeTotal = 0
                
                'If the cell immediately following a row is the same ticker...
                Else
                
                'Add to the Volume Total
                VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                
                End If
                
            Next i
        'DISPLAY THE TICKER SYMBOL TO COINCIDE WITH THE TOTAL VOLUME
            
            'Set Header Name
            ws.Range("I1").Value = "Ticker"
    Next ws
End Sub
