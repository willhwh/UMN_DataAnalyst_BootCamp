Attribute VB_Name = "dYearlyChange"


'Yearly Change Info
'try use loop to calculate yearly change, percent change and total stock volumn one by one for each ticker at once

Sub YearlyChange():
    
    
    'Define variables for iterating
    Dim r As Variant
    Dim l As Variant
    
    
    'Define variables for calculating YearlyChange elements
    Dim start As Variant
    Dim final As Variant
    Dim p1 As Variant
    Dim p2 As Variant
    Dim change As Variant
    Dim totalvolume As Variant
    Dim percentage As Variant
        
    'Loop all the unique tickers' index
    For l = 0 To Lengths:
        
        
        'Assign default value for all the elements
        p1 = 0
        p2 = 0
        change = 0
        totalvolum = 0
        start = 0
        final = 0
        
        
        'Beginning and closing price for beginning and end of that year
        'Loop all the data row index
        For r = 2 To rows:
            'Get the opening price at the beginning of a given year
            If p1 = 0 Then
                If Cells(r, 1) = tickerslist(l) Then
                    p1 = Cells(r, 3).Value
                    'Get the beginning index of a given year
                    start = r
                End If
            End If
            
            
            'Get the the closing price at the end of that year.
            If p1 <> 0 And p2 = 0 Then
                If Cells(r, 1) <> Cells(r + 1, 1) Then
                    p2 = Cells(r, 6).Value
                    'Get the closing index at the end of that year
                    final = r
                End If
            End If
        Next r
        
        
        'Yearly change value
        'Calculate the yearly change
        change = p2 - p1
        'Sssign the value to yearly chagne column
        Cells(l + 2, 11) = change
    
    
        'If the change is postive
        If Cells(l + 2, 11).Value > 0 Then
            'If positive fill up with bright green
            Cells(l + 2, 11).Interior.ColorIndex = 4
        ElseIf Cells(l + 2, 11).Value < 0 Then
            'Others fillup withbrigt red
            Cells(l + 2, 11).Interior.ColorIndex = 3
        End If

       
        'Calculate the percentage change
        'If p1 =0 then shows error
        'Difine an erro message
        Dim errormessage As String
        
        
        If p1 = 0 Then
            errormessage = "error cause beginning price was 0 "
            Cells(l + 2, 12) = errormessage
            Cells(l + 2, 13) = errormessage
        'Else calculate the percentage
        Else:
            percentage = (p2 - p1) / p1
            'Round up the percentage change to 4decimal
            percentage = Application.WorksheetFunction.RoundUp(percentage, 4)
            'Make it 100% format
            'Assign the value to percentage chagne column
            Cells(l + 2, 12) = Str(percentage * 100) + "%"
            'Calculate the sum of the total stock volume
            totalvolume = Application.sum(Range(Cells(start, 7), Cells(final, 7)))
            'Assign the value to total stock volume column
            Cells(l + 2, 13) = totalvolume
        End If
    Next l

    
End Sub
