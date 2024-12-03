Attribute VB_Name = "Module1"
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TotaVolume As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim SummaryRow As Integer
    Dim LastRow As Long
    Dim StartRow As Long
    Dim GreatestIncrease As Double, GreatestDecrease As Double, GreatestVolume As Double
    Dim GreatestIncreaseTicker As String, GreatestDecreaseTicker As String, GreatestVolumeTicker As String
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' ISet sresults headers
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Result"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Percentage"
        
        ' Initialize tracking variables
        
        TotalVolume = 0
        SummaryRow = 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        StartRow = 2
        
        ' Loop through rows
        For i = 2 To LastRow
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ' Check if next ticker is different
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                'open  price with
                ' Dim StartRow As Long
                
                If StartRow < 2 Or StartRow > LastRow Then
                     MsgBox "Error: StartRow (" & StartRow & ") is out of bounds.", vbCritical
                     Exit Sub
                End If
        
                OpenPrice = ws.Cells(StartRow, 3).Value
                
                'Close Price
                If i < 2 Or i > LastRow Then
                     MsgBox "Error: ClosePrice row (" & i & ") is out of bounds.", vbCritical
                     Exit Sub
                End If

                ClosePrice = ws.Cells(i, 6).Value
                
                
                QuarterlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (QuarterlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If
                
                ' Output data to summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = QuarterlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                
                ' Highlight changes
                If QuarterlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Track greatest values
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestIncreaseTicker = Ticker
                End If
                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestDecreaseTicker = Ticker
                End If
                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
                End If
                
                ' Move to next summary row
                SummaryRow = SummaryRow + 1
                TotalVolume = 0
                StartRow = i + 1
            End If
        Next i
        
        ' Output greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = GreatestIncreaseTicker
        ws.Cells(2, 17).Value = GreatestIncrease
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = GreatestDecreaseTicker
        ws.Cells(3, 17).Value = GreatestDecrease
        
        ws.Cells(4, 15).Value = "Greatest  TotalVolume"
        ws.Cells(4, 16).Value = GreatestVolumeTicker
        ws.Cells(4, 17).Value = GreatestVolume
    Next ws
End Sub


