# VBA_MarketPredictor
# VBA Analysis of Stocks by ticker for quick analysis 


Sub Predictor():

' Establish Variables
    Dim Ticker As String
    Dim i, sumi As Integer
    Dim Yearly_change As Integer
    Dim Percent_change As Integer
    Dim ttl_stock As Long
    Dim row As LongLong
    Dim total As Long
    Dim ticker_total As LongLong
    Dim ticker_open As Integer
    Dim ws As Worksheet
    
For Each ws In Worksheets

  '  Dim Summary_Table_Row As Integer
'  Summary_Table_Row = 2
    
    ' Create Column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

' Fill in "Ticker" column with ticker symbols and create downstream variables
' row = 2
ticker_total = ws.Cells(Rows.Count, 1).End(xlUp).row
ticker_high = ws.Cells(4, 1).End(xlDown).row
ticker_low = ws.Cells(5, 1).End(xlDown).row
ticker_close = ws.Cells(6, 1).End(xlDown).row
ticker_volume = ws.Cells(7, 1).End(xlDown).row



'' BEST OPTION SO FAR - reset variables at each new ws
total = 1
ttl_stock = 0
row = 2
Summary = 2
year_chng = 0
prcnt_chng = 0
'ttl_stock = Cells(row, 7).Value + ttl_stock
For i = 2 To ticker_total
        
'recognize ticker change
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    
    ws.Cells(Summary, 9).Value = ws.Cells(i, 1).Value
    ttl_stock = ttl_stock + ws.Cells(row, 7).Value
    ws.Cells(Summary, 12).Value = ttl_stock
    ticker_openp = ws.Cells(i + 1, 3).Value
    

    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'ticker_openp = Cells(i + 1, 3).Value
    ticker_closep = ws.Cells(i, 6).Value

    year_chng = ticker_closep - ticker_openp
    'Cells(Summary, 10).Value = ticker_closep - ticker_openp
    ws.Cells(Summary, 10).Value = year_chng
    prcnt_chng = (year_chng / ticker_openp) * 100
    ws.Cells(Summary, 11).Value = prcnt_chng
    ws.Cells(Summary, 11).NumberFormat = "0.00%"
    Summary = Summary + 1

    End If

    If (ticker_closep - ticker_openp) < 0 Then
        ws.Cells(Summary, 10).Interior.ColorIndex = 3
        ws.Cells(Summary, 11).Interior.ColorIndex = 3
    ElseIf (ticker_closep - ticker_openp) > 0 Then
        ws.Cells(Summary, 10).Interior.ColorIndex = 4
        ws.Cells(Summary, 11).Interior.ColorIndex = 4
    End If
    Next i
    Next ws
    
End Sub

