Attribute VB_Name = "Module1"
Sub Stock_Data()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    Dim Ticker As String
    
    Dim QuarterlyChange As Double
    
    Dim PercentageChange As Double
    
    Dim OpeningPrice As Double
    
    Dim ClosingPrice As Double
    
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
    Dim LastRow As Long
    
    SummaryRowTable = 2
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Ticker = ws.Cells(i, 1).Value
        
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        OpeningPrice = ws.Cells(i, 3).Value
        ClosingPrice = ws.Cells(i, 6).Value
        QuarterlyChange = ClosingPrice - OpeningPrice
        PercentageChange = (QuarterlyChange / OpeningPrice) * 100
        
        If OpeningPrice = 0 Then
            PercentageChange = (QuarterlyChange / OpeningPrice) * 100
        Else
            PercentageChange = 0
        End If
        
        ws.Range("I" & SummaryRowTable).Value = Ticker
        ws.Range("J" & SummaryRowTable).Value = QuarterlyChange
        ws.Range("K" & SummaryRowTable).Value = PercentageChange
        ws.Range("L" & SummaryRowTable).Value = TotalStockVolume
        
        SummaryRowTable = SummaryRowTable + 1
        
        TotalStockVolume = 0
        
        Else
    
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        If ws.Cells(SummaryRowTable, 10).Value > 0 Then
            ws.Cells(SummaryRowTable, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(SummaryRowTable, 10).Interior.ColorIndex = 3
        End If
        
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim GreatestPercentIncreaseTicker As String
    Dim GreatestPercentDecreaseTicker As String
    Dim GreatestTotalVolumeTicker As String

    ws.Cells(2, 16).Value = GreatestPercentIncrease
    ws.Cells(3, 16).Value = GreatestPercentDecrease
    ws.Cells(4, 16).Value = GreatestTotalVolume
    ws.Cells(2, 15).Value = GreatestPercentIncreaseTicker
    ws.Cells(3, 15).Value = GreatestPercentDecreaseTicker
    ws.Cells(4, 15).Value = GreatestTotalVolumeTicker
    
    GreatestPercentIncrease = 0
    GreatestPercentDecrease = 0
    GreatestTotalVolume = 0
    
        If PercentageChange > GreatestPercentIncrease Then
            GreatestPercentIncrease = PercentageChange
            GreatestPercentIncreaseTicker = Ticker
        ElseIf PercentageChange < GreatestPercentDecrease Then
            GreatestPercentDecrease = PercentageChange
            GreatestPercentDecreaseTicker = Ticker
        End If
    
        If TotalStockVolume > GreatestTotalVolume Then
            GreatestTotalVolume = TotalStockVolume
            GreatestTotalVolumeTicker = Ticker
        End If


    
    End If
    
    Next i
    
    Next ws
    
End Sub

