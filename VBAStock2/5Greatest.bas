Attribute VB_Name = "eGreatest"
' the total number of unique tickers for usage in Greatest(), Smallest(), and GreatestVolumn() usage
Global StockInfoLength As Variant

'Get the value of the greatest % increase for ticker and value
Sub Greatest():
    
    'Define variables for iterating
    Dim i As Variant
    
    
    'Calculate the total number of unique tickers
    StockInfoLength = Range("J1").End(xlDown).Row
    
    
    'Define variable
    Dim Max As String 'name of the greatest % increase value
    Dim M As String 'name of ticker thathas the greatest % increase
    
    'Calculate the value of the greatest % increase
    Max = Application.WorksheetFunction.Max(Range(Cells(2, 12), Cells(StockInfoLength, 12)))
    
    
    'Loop the unique name to find the one has the greatest % increase
    For i = 2 To StockInfoLength:
        If Cells(i, 12) = Max Then
            'Assign the ticker to the greatest % increase ticker cell
            M = Cells(i, 10).Value
            Cells(2, 16) = M
        End If
    Next i
    
    
    'Change to 100% format
    'Assign the ticker to the greatest % decrease ticker cell
    Cells(2, 17) = Str(Max * 100) + "%"
    
    
    
End Sub

'Get the value of the greatest % decrease for ticker and value
Sub Smallest():


    'Define variables for iterating
    Dim i As Variant
    
    
    'Define variable
    Dim Min As String 'name of the greatest % decrease value
    Dim S As String 'name of ticker thathas the greatest % decrease
    
    
    'calculate the value of the greatest % decrease
    Min = Application.WorksheetFunction.Min(Range(Cells(2, 12), Cells(StockInfoLength, 12)))
    
    
    'loop the unique name to find the one has the greatest % decrease
     For i = 2 To StockInfoLength:
        If Cells(i, 12) = Min Then
            'Assign the ticker to the greatest % decrease ticker cell
            S = Cells(i, 10).Value
            Cells(3, 16) = S
        End If
    Next i
    
    
    'Change to 100% format
    'Assign the value to the greatest % decrease value cell
    Cells(3, 17) = Str(Min * 100) + "%"
    
End Sub

'Get value of the greatest % total volume for ticker and value
Sub GreatestVolumn():

    'Define variables for iterating
    Dim i As Variant

    'Define variable
    Dim GV As String 'name of the greatest% total value
    Dim G As String 'name of the ticker that has the greatest % total volume
    
    'Loop the unique name to find the one has the greatest % total volume
    GV = Application.WorksheetFunction.Max(Range(Cells(2, 13), Cells(StockInfoLength, 13)))
    
    
    'Loop the unique name to find the one has the greatest % total volume
    For i = 2 To StockInfoLength:
        If Cells(i, 13) = GV Then
            'Assign the ticker to the greatest % total volume ticker cell
            G = Cells(i, 10).Value
            Cells(4, 16) = G
        End If
    Next i
    
    
    'Assign the value to  the greatest % total volume ticker cell
    Cells(4, 17) = GV
    
    
End Sub




