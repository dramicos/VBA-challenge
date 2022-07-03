Attribute VB_Name = "Module1"
Sub StockTracker()
    
    'Loop through each tab
    For Each ws In Worksheets
    
        'Label the new columns for the stock listing summary and make the cells fit
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Volume"
        
        'create a variable to find the maximum % increase and initialize
        Dim MaxIncrease As Double
        MaxIncrease = 0
        Dim MaxIncreaseTicker As String
        MaxIncreaseTicker = "AAB" 'set to first stock
        
        'create a variable to find the maximum % decrease and initialize
        Dim MaxDecrease As Double
        MaxDecrease = 0
        Dim MaxDecreaseTicker As String
        MaxDecreaseTicker = "AAB" 'set to first stock
        
        'create a variable to find the maximum % increase and initialize
        Dim MaxVolume As Double
        MaxVolume = 0
        Dim MaxVolumeTicker As String
        MaxVolumeTicker = "AAB" 'set to first stock

        
        'find the last row to no when to end loop
        Dim LastRow As Long
    
        'find last row of tickers
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'variable to collecting stock tickers to put in stocks listing summary
        Dim StockTick As String
    
        'variable for Yearly change and initialize
        Dim StockChange As Double
        StockChange = 0
    
        'variable to add up stock volume and initialize
        Dim Volume As Double
        Volume = 0
    
        'variable for holding the stock ticker listing number and initialize for starting on row 2
        Dim StockNumber As Double
        StockNumber = 2
    
        'Variables for starting and ending stock value and initialize
        Dim StockStart As Double
        Dim StockEnd As Double
        StockStart = 0
        StockEnd = 0
        
               
        'Sort through the data and populate the stock listing summary with the data
        For i = 2 To LastRow
    
            'check that were at first line of stock data and get the start price
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                StockStart = ws.Cells(i, 3).Value
                
            End If
                   
               
            'check that we are still on the same stock and get data for summary
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                'If we have reached the last occurance of the stock in the list we...
            
                'Set the stock name
                StockTick = ws.Cells(i, 1).Value
            
                'get the final stock price
                StockEnd = ws.Cells(i, 6).Value
            
                'Set the final Stock Volume value
                Volume = Volume + ws.Cells(i, 7)
            
                'put the stock ticker name in the summary list
                ws.Cells(StockNumber, 9).Value = StockTick
            
                'Put the total stock volume in the list summary
                ws.Cells(StockNumber, 12).Value = Volume
            
                'Calculate the yearly change in stock price and put the value in the list
                StockChange = StockEnd - StockStart
                ws.Cells(StockNumber, 10).Value = StockChange
                
                                               
                'Color the stock change field depending on the value
                If StockChange >= 0 Then
                    ws.Cells(StockNumber, 10).Interior.ColorIndex = 4 'light green
                Else
                    ws.Cells(StockNumber, 10).Interior.ColorIndex = 3 'red
                End If
                
                'Show The change as a percentage
                ws.Cells(StockNumber, 11).Value = StockChange / StockStart
                ws.Cells(StockNumber, 11).NumberFormat = "0.00%"
                
                'Increment the summary index to the next stock
                StockNumber = StockNumber + 1
            
                'Reset Volume for next stock
                Volume = 0
            Else
                'This is what to do while still in the same stock
            
                'Add up the total stock volume
                Volume = Volume + ws.Cells(i, 7).Value
                
                          
            End If
            
            
            
        Next i
        'Find the last row of the summary list
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Get the biggest Increase
        For j = 2 To LastRow2
            'check to see if the current stock is bigger
            If ws.Cells(j, 11).Value > MaxIncrease Then
                MaxIncrease = ws.Cells(j, 11).Value
                MaxIncreaseTicker = ws.Cells(j, 9).Value
            End If
        Next j
        
        'Get the biggest Decrease
        For j = 2 To LastRow2
            'check to see if the current stock is bigger
            If ws.Cells(j, 11).Value < MaxDecrease Then
                MaxDecrease = ws.Cells(j, 11).Value
                MaxDecreaseTicker = ws.Cells(j, 9).Value
            End If
        Next j
        
        'Get the biggest Volume
        For j = 2 To LastRow2
            'check to see if the current stock is bigger
            If ws.Cells(j, 12).Value > MaxVolume Then
                MaxVolume = ws.Cells(j, 12).Value
                MaxVolumeTicker = ws.Cells(j, 9).Value
            End If
        Next j
        
        ws.Range("P2").Value = MaxIncreaseTicker
        ws.Range("Q2").Value = MaxIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
   
        ws.Range("P3").Value = MaxDecreaseTicker
        ws.Range("Q3").Value = MaxDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
    
        ws.Range("P4").Value = MaxVolumeTicker
        ws.Range("Q4").Value = MaxVolume
        
        'autofit the columns in the newly poplulated columns
        ws.Range("I:Q").Columns.AutoFit
    Next ws
    
    

End Sub


