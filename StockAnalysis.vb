Sub StockAnalysis()
    
  ' initiate variables
  Dim ticker, tickers(3) As String
  Dim oPirce, cPrice As Double
  Dim tVol, nRow, lastRow As Long
  Dim nCol As Integer
  ' output headers
  Dim header() As Variant: header = Array("ticker", "quarterly_change", "percent_change", "total_volume")
  Dim headerMax() As Variant: headerMax = Array("ticker", "value")
  Dim headerRow() As Variant: headerRow = Array("greatest_%increase", "greatest_%decrease", "greatest_total_volume")
  Dim headerLen As Integer: headerLen = UBound(header) - LBound(header) + 1
  
  ' iterate through each sheet
  For Each ws In Worksheets
    ' initialize
    ticker = ws.Cells(2, 1).Value
    oPrice = ws.Cells(2, 3).Value
    cPrice = ws.Cells(2, 6).Value
    Dim maxs() As Variant: maxs = Array(0, 0, 0)
    ' total volume
    tVol = 0
    ' output row/column
    nRow = 2
    nCol = 9
    ' print output header
    For i = 0 To 3
        ws.Cells(1, nCol + i).Value = header(i)
    Next i
    For i = 0 To 1
        ws.Cells(1, nCol + headerLen + 3 + i).Value = headerMax(i)
    Next i
    ' do analysis
    lastRow = ws.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    'iterate through the quarter
    For i = 2 To lastRow
        'check if ticker has changed
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Dim percent As Double: percent = (ws.Cells(i, 6).Value - oPrice) / oPrice
            ' output data
            tVol = tVol + ws.Cells(i, 7).Value
            ws.Cells(nRow, nCol).Value = ticker
            ws.Cells(nRow, nCol + 1).Value = ws.Cells(i, 6).Value - oPrice
            If ws.Cells(i, 6).Value - oPrice > 0 Then
                ws.Cells(nRow, nCol + 1).Interior.Color = "&HB8D7C4"
            Else
                ws.Cells(nRow, nCol + 1).Interior.Color = "&H8FBFFA"
            End If
            ws.Cells(nRow, nCol + 2).Value = percent
            ws.Cells(nRow, nCol + 2).NumberFormat = "0.00%"
            ws.Cells(nRow, nCol + 3).Value = tVol
            ' set maxs
            If percent > maxs(0) Then maxs(0) = percent: tickers(0) = ticker
            If percent < maxs(1) Then maxs(1) = percent: tickers(1) = ticker
            If tVol > maxs(2) Then maxs(2) = tVol: tickers(2) = ticker
            ' reset for next ticker
            nRow = nRow + 1
            ticker = ws.Cells(i + 1, 1).Value
            oPrice = ws.Cells(i + 1, 3).Value
            tVol = 0
        'if ticker is same, continue adding to total volume
        Else
            tVol = tVol + Int(ws.Cells(i, 7).Value)
        End If
    Next i
    
    'print analysis
    For i = 0 To 2
        ws.Cells(2 + i, nCol + headerLen + 2).Value = headerRow(i)
        ws.Cells(2 + i, nCol + headerLen + 3).Value = tickers(i)
        ws.Cells(2 + i, nCol + headerLen + 4).Value = maxs(i)
    Next i
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
  Next ws
End Sub