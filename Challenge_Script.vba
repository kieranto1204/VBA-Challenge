Sub challenge()

Dim sheet As Worksheet

    For Each sheet In ActiveWorkbook.Worksheets
    
    sheet.Activate
    
        bottom = sheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Quarterly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"
        Cells(2, 15) = "Greatest % Increase"
        Cells(3, 15) = "Greatest % Decrease"
        Cells(4, 15) = "Greatest Total Volume"
        
        Dim ticker As String
        Dim quartchange As Double
        Dim perchange As Double
        Dim totalstock As Double
        
        Dim tickerrow As Double
        Dim volume As Double
        Dim startprice As Double
        Dim endprice As Double
        
        volume = 0
        Row = 2
        
        startprice = Cells(2, 3).Value
        
        For i = 2 To bottom
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ticker = Cells(i, 1).Value
                Cells(Row, 9).Value = ticker
                
                endprice = Cells(i, 6).Value
                
                quartchange = endprice - startprice
                Cells(Row, 10).Value = quartchange
                
                    If quartchange > 0 Then
                        Cells(Row, 10).Interior.ColorIndex = 4
                    ElseIf Cells(Row, 10) = 0 Then
                        Cells(Row, 10).Interior.ColorIndex = 0
                    ElseIf Cells(Row, 10).Value < 0 Then
                        Cells(Row, 10).Interior.ColorIndex = 3
                    End If
                    
                perchange = quartchange / startprice
                Cells(Row, 11).Value = perchange
                Cells(Row, 11).NumberFormat = "0.00%"
                
                volume = volume + Cells(i, 7).Value
                Cells(Row, 12).Value = volume
                
                Row = Row + 1
                
                startprice = (Cells(i + 1, 3))
                
                volume = 0
            
            Else
                
                volume = volume + Cells(i, 7).Value
                
            End If
        
        Next i
        
        bottom2 = sheet.Cells(Rows.Count, 9).End(xlUp).Row
        
        For j = 2 To bottom2
        
            If Cells(j, 11).Value = Application.WorksheetFunction.Max(Range("k2:k" & bottom2)) Then
                Cells(2, 16).Value = Cells(j, 9).Value
                Cells(2, 17).Value = Cells(j, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
                
            ElseIf Cells(j, 11).Value = Application.WorksheetFunction.Min(Range("k2:k" & bottom2)) Then
                Cells(3, 16).Value = Cells(j, 9).Value
                Cells(3, 17).Value = Cells(j, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
                
            ElseIf Cells(j, 12).Value = Application.WorksheetFunction.Max(sheet.Range("L2:L" & bottom2)) Then
                Cells(4, 16).Value = Cells(j, 9).Value
                Cells(4, 17).Value = Cells(j, 12).Value
            End If
            
        Next j
            
        
        ActiveSheet.Columns("A:Z").AutoFit
    
    Next sheet

End Sub

