Sub ticker_macro()

'Set variable for worksheet
Dim ws As Worksheet

'Loop through all data in each worksheet
For Each ws In Worksheets
    
   'Set variable for stock ticker
    Dim ticker As String
    ticker = " "
    
    'Set variable for total for every stock ticker
    Dim total_ticker As Double
    total_ticker = 0
    
    
    'Set variable for change in price
    Dim price_change As Double
    price_change = 0
    
    'Set variable for percent of change
    Dim percent_change As Double
    percent_change = 0
    
    'Set variable for open stock price
    Dim open_price As Double
    open_price = 0
    
    'Set variable for close stock price
    Dim close_price As Double
    close_price = 0
    
    'Set variable for minimum stock ticker name
    Dim min_ticker As String
    min_ticker = " "
    
    'Set variable for maximum stock ticker name
    Dim max_ticker As String
    max_ticker = " "
    
    'Set variable for maximum volume
    Dim max_volume As Double
    max_volume = 0
    
    'Set maximum volume ticker variable
    Dim max_ticker_vol As String
    max_ticker_vol = " "
    
    'Set minimum percent variable
    Dim min_percent As Double
    min_percent = 0
    
    'Set maximum percent variable
    Dim max_percent As Double
    max_percent = 0
    
    'Set parameters for summary table
    Dim summary_table As Long
    summary_table = 2
    
    'Set variable for last row and row numer to start
    Dim Lastrow As Long
    Dim i As Long
    
    'Set parameter for last row in worksheet
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        ' Set headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Set headers for other table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        
        'Start open price for first cell location
        open_price = ws.Cells(2, 3).Value
        
        'Loop entire worksheet until last row
        For i = 2 To Lastrow
        
            'Determine of the next cell is the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set ticker to current cell
                ticker = ws.Cells(i, 1).Value
                
                'Calculate price change and percent
                close_price = ws.Cells(i, 6).Value
                price_change = close_price - open_price
                
                'Verify open price does not equal to zero
                If open_price <> 0 Then
                    percent_change = (price_change / open_price) * 100
                    
                Else
                    MsgBox (Continue)
                    
                End If
                
                'Adding up stock ticker total volume with current cell value to total
                total_ticker = total_ticker + ws.Cells(i, 7).Value
                
                'Insert ticker in summary table
                ws.Range("I" & summary_table).Value = ticker
                
                'Insert price change in summary table
                ws.Range("J" & summary_table).Value = price_change
                
                'Fill cell green if greater than zero
                If (price_change > 0) Then
                    ws.Range("J" & summary_table).Interior.ColorIndex = 4
                   
                'Fill cell red if less than zero
                ElseIf (change_price <= 0) Then
                    ws.Range("J" & summary_table).Interior.ColorIndex = 3
                    
                End If
                
                'Insert value for percent change and total_ticker in column "K" and column "L"
                ws.Range("K" & summary_table).Value = (CStr(percent_change) & "%")
                ws.Range("L" & summary_table).Value = total_ticker
                
                'Add 1 to the row in summary table
                summary_table = summary_table + 1
                
                'Reset change_price and percent_change back to zero
                close_price = 0
                price_change = 0
                
                'Adding open price value to next cell value
                open_price = ws.Cells(i + 1, 3).Value
                
                'Determine if current value is max value
                If (percent_change > max_percent) Then
                    
                    max_percent = percent_change
                    
                    max_ticker = ticker
                    
                'Determine if current value is minimum percent
                ElseIf (percent_change < min_percent) Then
                        
                    min_percent = percent_change
                        
                    min_ticker = ticker
                    
                End If
                
                'Determine if current value is max volume
                If (total_ticker > max_volume) Then
                
                    max_volume = total_ticker
                    
                    max_ticker_vol = ticker
                    
                End If
                
                'Reset counters to zero
                percent_change = 0
                
                total_ticker = 0
                
        
        'If next row is the same ticker add to total volume
        Else
            
            total_ticker = total_ticker + ws.Cells(i, 7).Value
            
        End If
        
    Next i
     'Insert values into table for max percent, min percent, max ticker, min ticker, max volume, and max ticker volume
            ws.Range("Q2").Value = (CStr(max_percent) & "%")
            ws.Range("Q3").Value = (CStr(min_percent) & "%")
            ws.Range("P2").Value = max_ticker
            ws.Range("P3").Value = min_ticker
            ws.Range("Q4").Value = max_volume
            ws.Range("P4").Value = max_ticker_vol
    
    
       
Next ws
    
End Sub
