VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock()
    Dim ws As Worksheet
    Dim ticker As String
    Dim volume As Double
    Dim year_open As Double, year_close As Double, yearly_change As Double, percent_change As Double
    Dim summary_table As Integer
    Dim i As Long, j As Long, k As Long
    

    For Each ws In ThisWorkbook.Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Volume"
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Cells(2, 15).NumberFormat = "0.00%"
        ws.Cells(3, 15).NumberFormat = "0.00%"
        ws.Cells(2, 13).Value = "Greatest % Increase"
        ws.Cells(3, 13).Value = "Greatest % Decrease"
        ws.Cells(4, 13).Value = "Greatest Volume Total"
        summary_table = 2
          
        'For Tickers & Values"
        For i = 2 To ws.UsedRange.Rows.Count
            If i = 2 Then
                year_open = ws.Cells(i, 3).Value
            End If
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                volume = volume + ws.Cells(i, 7).Value
                year_close = ws.Cells(i, 6).Value
                yearly_change = (year_close - year_open)
                percent_change = (yearly_change / year_open)
                
                ws.Cells(summary_table, 9).Value = ticker
                ws.Cells(summary_table, 10).Value = yearly_change
                ws.Cells(summary_table, 11).Value = percent_change
                ws.Cells(summary_table, 12).Value = volume
                summary_table = summary_table + 1
                
                year_open = ws.Cells(i + 1, 3).Value
                volume = 0
            Else
                volume = volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        'For Color'
        For j = 2 To ws.UsedRange.Rows.Count
            If ws.Cells(j, 11).Value < 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(j, 11).Interior.ColorIndex = 4
            End If
        Next j
        
        'For Increase, Decrease, & Volume"
        Dim large As String, small As String, great_vol As String
        Dim increase As Double, decrease As Double, large_volume As Double
        
        
        increase = 0
        decrease = 0
        large_volume = 0
        
        For k = 2 To ws.UsedRange.Rows.Count
            If ws.Cells(k, 11).Value > increase Then
                increase = ws.Cells(k, 11).Value
                large = ws.Cells(k, 9).Value
            End If
            
            If ws.Cells(k, 11).Value < decrease Then
                decrease = ws.Cells(k, 11).Value
                small = ws.Cells(k, 9).Value
            End If
            
            If ws.Cells(k, 12).Value > large_volume Then
                large_volume = ws.Cells(k, 12).Value
                great_vol = ws.Cells(k, 9).Value
            End If
        Next k
        
        ws.Cells(2, 14).Value = large
        ws.Cells(3, 14).Value = small
        ws.Cells(4, 14).Value = great_vol
    
        ws.Cells(2, 15).Value = increase
        ws.Cells(3, 15).Value = decrease
        ws.Cells(4, 15).Value = large_volume
    
        
    Next ws
End Sub

