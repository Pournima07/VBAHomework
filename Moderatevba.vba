Sub ModerateCountTicker()

    Dim num_sheets As Integer
    Dim current_sheet As Integer
    Dim sheetname As String
    
    Dim reading_rownbr As Long
    Dim writing_rownbr As Long
    
    Dim current_volume  As Double
    Dim total_vol As Double
    Dim opening_price As Double
    Dim closing_price As Double
    Dim difference As Double
    Dim percentage As Double
    
    
    num_sheets = Worksheets.Count
    current_sheet = 1
    
    ' Work on all sheets in the file
    Do While current_sheet <= num_sheets
        
        ' Select current sheet nbr
        Worksheets(current_sheet).Activate
        sheetname = ActiveSheet.Name
        MsgBox "Working on Worksheet: " & sheetname
        
        ' Initialize variables at start of each sheet
        reading_rownbr = 2
        writing_rownbr = 2
        current_ticker = Trim(Cells(2, 1).Value)
        current_total = 0
        opening_price = Cells(2, 3).Value
        
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Open at"
        Cells(1, 11) = "Closed"
        Cells(1, 12) = "Yearly Change"
        Cells(1, 13) = "Percent Change"
        Cells(1, 14) = "Total Stock Volume"

        Do While current_ticker <> ""
        
            ' Read ticker on current reading line
            this_ticker = Trim(Cells(reading_rownbr, 1).Value)

            ' Same ticker, add volume to total and move reading line one down
            If this_ticker = current_ticker Then
                
                current_volume = Cells(reading_rownbr, 7).Value
                current_total = current_total + current_volume
                closing_price = Cells(reading_rownbr, 6).Value
                reading_rownbr = reading_rownbr + 1
            
            Else ' Started reading next ticker
            
                ' Save total of previous ticker to writing column (at writing row nbr)
                difference = closing_price - opening_price
                If opening_price > 0 Then
                    percentage = difference / opening_price
                Else
                    percentage = 0
                End If
                
                Cells(writing_rownbr, 9) = current_ticker
                Cells(writing_rownbr, 10) = opening_price
                Cells(writing_rownbr, 11) = closing_price
                Cells(writing_rownbr, 12) = difference
                Cells(writing_rownbr, 13) = percentage
                Cells(writing_rownbr, 14) = current_total
                
                ' Move writing row nbr onw down
                writing_rownbr = writing_rownbr + 1
                                
                ' Current progress
                Debug.Print current_ticker & " - Volume: " & current_total & ", Opening: " & opening_price & ", closing: " & closing_price & ", difference: " & difference & ", percentage: " & percentage
                
                ' Now start collecting total volume for this new ticker
                current_ticker = this_ticker
                current_total = Cells(reading_rownbr, 7).Value
                opening_price = Cells(reading_rownbr, 6).Value
                closing_price = opening_price
                
                ' And move reading row nbr to one down
                reading_rownbr = reading_rownbr + 1
                
            
            End If
        
        Loop
        
        start_row = 2
        end_row = writing_rownbr - 1
        col_nbr = 12
        
        conditional_format start_row, end_row, col_nbr
        
        current_sheet = current_sheet + 1
    
    Loop
    
    MsgBox "Done"

End Sub


Sub conditional_format(start_row, end_row, col_nbr)

    Dim percent_rg As Range
    Set percent_rg = Range(Cells(start_row, col_nbr + 1), Cells(end_row, col_nbr + 1))
    
    percent_rg.NumberFormat = "0.00%"

    Dim rg As Range
    Dim cond_plus As FormatCondition, cond_minus As FormatCondition, cond_nochange As FormatCondition
    
    Set rg = Range(Cells(start_row, col_nbr), Cells(end_row, col_nbr))
    
    'clear any existing conditional formatting
    rg.FormatConditions.Delete
    
    'define the rule for each conditional format
    Set cond_plus = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set cond_minus = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")
    Set cond_nochange = rg.FormatConditions.Add(xlCellValue, xlEqual, "=0")
    
    'define the format applied for each conditional format
    With cond_plus
    .Interior.Color = vbGreen
    .Font.Color = vbBlack
    End With
    
    With cond_minus
    .Interior.Color = vbRed
    .Font.Color = vbBlack
    End With
    
    With cond_nochange
    .Interior.Color = vbWhite
    .Font.Color = vbBlack
    End With
 
End Sub

