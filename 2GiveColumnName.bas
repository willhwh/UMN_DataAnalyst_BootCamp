Attribute VB_Name = "bGiveColumnName"
'ColumnName
Sub ColumnName():
    
    Cells(1, 10) = "Tickers"
    Cells(1, 11) = "YearlyChange"
    Cells(1, 12) = "PercentChange"
    Cells(1, 13) = "TotalStockVolumn"
    Cells(2, 15) = "Great % Increase"
    Cells(3, 15) = "greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    Debug.Print ("Column names have been given")
End Sub

