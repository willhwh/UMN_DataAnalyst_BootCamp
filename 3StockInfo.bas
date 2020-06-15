Attribute VB_Name = "cStockInfo"
    'global variales for YearlyChange() usage
    Global tickerslist() As String
    Global Lengths As Long
    Global rows As Long
'Ticker
Sub StockInfo():
    
    'Define variables for iterating
    Dim r As Variant
    Dim l As Variant
    
    
    'Define a name array for all tickers
    Dim tickers() As String
    
    'Give the first ticker name to the array
    ReDim Preserve tickers(0)
    tickers(0) = Cells(2, 1).Value
    
    'Give the number of total rows in sheet(1) to totalrow variable
    totalrow = ActiveSheet.UsedRange.rows.Count
    
    'Assign rows as totalrow for futrue usage
    rows = totalrow
    
    'Make sure the data is ordered by <ticker>,and <'date'>
    'Loop the rows in column A to get unique ticker name
    For r = 2 To totalrow:
        'if name unique
        If Cells(r, 1) <> Cells(r + 1, 1) Then
        'tickers array index+1
        ReDim Preserve tickers(UBound(tickers) + 1)
        'add unique name into tickers array
        tickers(UBound(tickers)) = Cells(r + 1, 1).Value
        End If
    Next r
    
    'assign tickerlist() as tickers() for future usage
    tickerslist() = tickers()
    
   
    'loop the tickers array to give unique name to column J
    Lengths = UBound(tickers())
    For l = 0 To Lengths:
        Cells(l + 2, 10) = tickers(l)
    Next l
    
End Sub

