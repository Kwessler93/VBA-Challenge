Sub StockAnalysis()

'Define Worksheet Variables

Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long

'Define Variables for Calculations

Dim ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim Volume As Double

'Define Output Variables

Dim QuarterlyPriceChange As Double
Dim QuarterlyPercentChange As Double
Dim TotalVolume As LongLong

'Define other varibles for calculation
   Dim GreatestPercentIncrease As Double
   Dim GreatestPercentDecrease As Double
   Dim GreatestTotalVolume As Double
        
    Dim IncreaseTickerName As String
    Dim DecraseTickerName As String
    Dim VolumeTickerName As String
        

'Assign headers for output

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

'Loop through each ticker and calculate quarterly change between price, percent change from opening to close, and total volume of stock

    SummaryRow = 2
    TotalVolume = 0
    GreatestPercentIncrease = 0
    GreatestPercentDecrease = 0
    GreatestTotalVolume = 0


    ticker = ws.Cells(2, 1).Value
    OpenPrice = ws.Cells(2, 3).Value

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

'Calculate total volume and changes

        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
        If ws.Cells(i + 1, 1).Value <> ticker Then
    
            ClosePrice = ws.Cells(i, 6).Value
        
        QuarterlyPriceChange = ClosePrice - OpenPrice
        
        If OpenPrice <> 0 Then
        
            QuarterlyPercentChange = QuarterlyPriceChange / OpenPrice
            
        Else
        
             QuarterlyPercentChange = 0
            
        End If
    
        ws.Cells(SummaryRow, 9).Value = ticker
        ws.Cells(SummaryRow, 10).Value = (QuarterlyPriceChange)
        ws.Cells(SummaryRow, 11) = QuarterlyPercentChange
        ws.Cells(SummaryRow, 12).Value = TotalVolume
        
        'conditional formatting
        ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
        ws.Cells(SummaryRow, 10).NumberFormat = "0.00;-0.00"
        
        If QuarterlyPriceChange >= 0 Then
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
        End If
    
        'add greatest percent increase, greatest, precent decrease, and greatest total volume
        
        If QuarterlyPercentChange > GreatestPercentIncrease Then
               GreatestPercentIncrease = QuarterlyPercentChange
               IncreaseTickerName = ticker
               End If
               
        If QuarterlyPercentChange < GreatestPercentDecrease Then
            GreatestPercentDecrease = QuarterlyPercentChange
            DecreaseTickerName = ticker
            End If
            
        If TotalVolume > GreatestTotalVolume Then
            GreatestTotalVolume = TotalVolume
            VolumeTickerName = ticker
            End If
            
    'Continue to Loop
    SummaryRow = SummaryRow + 1
    TotalVolume = 0
    ticker = ws.Cells(i + 1, 1).Value
    OpenPrice = ws.Cells(i + 1, 3).Value
        
    End If
      
Next i
        
        ws.Cells(2, 15).Value = " Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 16).Value = IncreaseTickerName
        ws.Cells(2, 17).Value = GreatestPercentIncrease
        ws.Cells(3, 16).Value = DecreaseTickerName
        ws.Cells(3, 17).Value = GreatestPercentDecrease
        ws.Cells(4, 16).Value = VolumeTickerName
        ws.Cells(4, 17).Value = GreatestTotalVolume
        
        'conditional formatting
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("I:Q").AutoFit

Next ws


End Sub