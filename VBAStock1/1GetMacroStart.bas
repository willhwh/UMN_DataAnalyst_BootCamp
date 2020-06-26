Attribute VB_Name = "aGetMarcoStart"


'main
Sub marcomain():
    ' Declare Current as a worksheet object variable.
    Dim Current As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each Current In Worksheets
        Call ColumnName
        Call StockInfo
        Call YearlyChange
        Call Greatest
        Call Smallest
        Call GreatestVolumn
        Debug.Print ("All subs finished processing")
         
    Next
    
End Sub


