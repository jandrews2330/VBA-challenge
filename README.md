# VBA-challenge

'I worked with several class members including Forrest Margulies, Haley Armenta, and Megan Iyer, as well as used ChatGPT and Learning Assistant to validate my code and identify errors. 


Sub StockLoop()
'Create variables
Dim ws As Worksheet
Dim i As Double
Dim ticker_name As String
Dim QChange As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim PercentChange As Double
Dim FirstOpen As Double
Dim LastClose As Double
Dim Summary_row As Long
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotalVolume As Double
Dim IncreaseTicker As String
Dim DecreaseTicker As String
Dim TotalVolume As String
Dim Stock_total As Double

'Loop through all sheets
For Each ws In Worksheets

'Create counters
QChange = 0
Stock_total = 0
GreatestIncrease = 0
GreatestDecrease = 0
GreatestTotalVolume = 0
Summary_row = 2

'Find the last row of each worksheet
Dim LastRow As Long
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'Headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Reset first open price
FirstOpen = ws.Cells(2, 3).Value

    'Loop through each row in worksheet
    For i = 2 To LastRow
         
        'Check if ticker symbol is the same
        If i = LastRow Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Set ticker name
            ticker_name = ws.Cells(i, 1).Value
            'Set last close price
            LastClose = ws.Cells(i, 6).Value
            'Quarterly change for each ticker
            QChange = LastClose - FirstOpen
            'Add to the stock total
            Stock_total = Stock_total + ws.Cells(i, 7).Value
            'Percent change for each ticker
            If FirstOpen <> 0 Then
                PercentChange = ((LastClose - FirstOpen) / FirstOpen)
            Else
                PercentChange = 0
            End If
            
            'Print the ticker symbol to column I in the Q1 tab
            ws.Range("I" & Summary_row).Value = ticker_name
            'Print the quarterly change to the output
            ws.Range("J" & Summary_row).Value = QChange
            'Print the percentage change to the output
            ws.Range("K" & Summary_row).Value = PercentChange
            'Format to percentage
            ws.Range("K" & Summary_row).NumberFormat = "0.00%"
            'Print the volume amount to the output
            ws.Range("L" & Summary_row).Value = Stock_total
            
                'Conditional formatting for Qchange
                'Green
                If QChange > 0 Then
                ws.Range("J" & Summary_row).Interior.ColorIndex = 4
                'Red
                ElseIf QChange < 0 Then
                ws.Range("J" & Summary_row).Interior.ColorIndex = 3
                'No fill
                Else
                ws.Range("J" & Summary_row).Interior.ColorIndex = xlNone
                End If
                                    
            'Greatest percent change and volume total
            If Stock_total > GreatestTotalVolume Then
            GreatestTotalVolume = Stock_total
            TotalVolume = ticker_name
            End If
            
            If PercentChange > GreatestIncrease Then
                GreatestIncrease = PercentChange
                IncreaseTicker = ticker_name
            End If
                        
            If PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                DecreaseTicker = ticker_name
            End If
                        
            'Add one to the ticker counter row
            Summary_row = Summary_row + 1
                        
            'Reset the FirstOpen, Qchange, and stock total
            If i < LastRow Then
                FirstOpen = ws.Cells(i + 1, 3).Value
                QChange = 0
                Stock_total = 0
            End If
                
        Else
            'Add to the stock total
            Stock_total = Stock_total + ws.Cells(i, 7).Value
        End If
    Next i
    
    'Print greatest increase, decrease, and total volume
    'Header
    ws.Range("O2").Value = "Greatest % Increase"
    'Output values
    ws.Range("P2").Value = IncreaseTicker
    ws.Range("Q2").Value = GreatestIncrease
    'Format to %
    ws.Range("Q2").NumberFormat = "0.00%"
    
    'Header
    ws.Range("O3").Value = "Greatest % Decrease"
    'Output values
    ws.Range("P3").Value = DecreaseTicker
    ws.Range("Q3").Value = GreatestDecrease
    'Format to %
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'Header
    ws.Range("O4").Value = "Greatest Total Volume"
    'Output values
    ws.Range("P4").Value = Stock_total
    ws.Range("Q4").Value = GreatestTotalVolume
    
Next ws

End Sub