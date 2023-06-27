Attribute VB_Name = "Module1"
Sub Stocks():
Dim Ticker_Name As String
Dim Volume_Total As Double
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Dim Summary_Table_Row As Integer
Dim open_price, close_price As Double
Dim start_date As String
Dim max_percent, min_percent, max_volume As Double


For Each ws In Worksheets
Summary_Table_Row = 2
Volume_Total = 0
start_date = ws.Range("B2").Value
'max_percent = WorksheetFunction.Max(Range("K2:K" & last_row))
'min_percent = WorksheetFunction.Min(Range("K2:K" & last_row))
'max_volume = WorksheetFunction.Max(Range("L2:L" & last_row))
    For i = 2 To last_row
        If ws.Cells(i, 2).Value = start_date Then open_price = ws.Cells(i, 3).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Name = ws.Cells(i, 1).Value
        close_price = ws.Cells(i, 6).Value
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
        ws.Range("J" & Summary_Table_Row).Value = close_price - open_price
        ws.Range("K" & Summary_Table_Row).Value = ((close_price - open_price) / open_price) * 100
          
        ws.Range("L" & Summary_Table_Row).Value = Volume_Total
        Summary_Table_Row = Summary_Table_Row + 1
        Volume_Total = 0
        
        Else
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
        
        End If
        
     Next i
    max_percent = WorksheetFunction.Max(Range("K2:K" & last_row))
    min_percent = WorksheetFunction.Min(Range("K2:K" & last_row))
    max_volume = WorksheetFunction.Max(Range("L2:L" & last_row))
    For i = 2 To last_row
    If ws.Cells(i, 11).Value = max_percent Then
        ws.Range("Q2").Value = max_percent
        ws.Range("P2").Value = ws.Cells(i, 9).Value
    ElseIf ws.Cells(i, 11).Value = min_percent Then
        ws.Range("Q3").Value = min_percent
        ws.Range("P3").Value = ws.Cells(i, 9).Value
    ElseIf ws.Cells(i, 12).Value = max_volume Then
        ws.Range("Q4").Value = max_volume
        ws.Range("P4").Value = ws.Cells(i, 9).Value
       End If
    Next i
Next ws
        
End Sub
