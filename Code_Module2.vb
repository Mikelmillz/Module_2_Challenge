Sub Stocks()

   Dim WS As Worksheet
         
    For Each WS In ThisWorkbook.Worksheets
        
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        WS.Cells(2, 15).Value = "Gratest % Increase"
        WS.Cells(3, 15).Value = "Breatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
    
            With WS.Columns("L")
                .ColumnWidth = 20
            End With
        
            With WS.Columns("O")
            .ColumnWidth = 20
            End With
        
            With WS.Columns("Q")
                .ColumnWidth = 10
            End With
        
            With WS.Columns("J")
                .NumberFormat = "0.00"
            End With
        
            With WS.Columns("K")
                .NumberFormat = "0.00%"
            End With

    
    
            Dim Stock_name As String
            Dim Stock_total As Double
            Stock_total = 0
        
            Dim Stock_Year_change As Double
            Dim Stock_Percent_change As Double

    
            Dim Summary_Table As Integer
            Summary_Table = 2
    
    
            Dim LastRow As Long
            LastRow = Cells(Rows.Count, "A").End(xlUp).Row
            j = 2
       
    
    
            Sumary_table = 2
            For I = 2 To LastRow
       
                If WS.Cells(I + 1, 1).Value <> WS.Cells(I, 1).Value Then
                    Stock_name = WS.Cells(I, 1).Value
                    Stock_total = Stock_total + WS.Cells(I, 7).Value
                    Stock_Year_change = WS.Cells(I, 6).Value - WS.Cells(j, 3).Value
                    Stock_Percent_change = WS.Cells(I, 6).Value / WS.Cells(j, 3).Value - 1
                    WS.Range("i" & Summary_Table).Value = Stock_name
                    WS.Range("j" & Summary_Table).Value = Stock_Year_change
                    WS.Range("k" & Summary_Table).Value = Stock_Percent_change
                    WS.Range("l" & Summary_Table).Value = Stock_total
                    Summary_Table = Summary_Table + 1
                    Stock_total = 0
                    j = I + 1
        
    
            Else
                Stock_total = Stock_total + WS.Cells(I, 7).Value
    
            End If
        
        
            Next I
           
         Dim N As Integer
         Dim LastRow2 As Long
            LastRow2 = WS.Cells(Rows.Count, "J").End(xlUp).Row
         For N = 2 To LastRow2
         
        
        If WS.Cells(N, 10).Value > 0 Then
            WS.Cells(N, 10).Interior.ColorIndex = 4
        ElseIf WS.Cells(N, 10).Value < 0 Then
            WS.Cells(N, 10).Interior.ColorIndex = 3
        ElseIf WS.Cells(N, 10).Value = 0 Then
            WS.Cells(N, 10).Interior.ColorIndex = 5
        End If
        
        Next N
        

            
    
        WS.Cells(2, 17).Value = Application.WorksheetFunction.Max(WS.Range("K:K"))
        WS.Cells(2, 17).NumberFormat = "0.00%"
    
    
        WS.Cells(3, 17).Value = Application.WorksheetFunction.Min(WS.Range("K:K"))
        WS.Cells(3, 17).NumberFormat = "0.00%"
    
    
        WS.Cells(4, 17).Value = Application.WorksheetFunction.Max(WS.Range("L:L"))
        Dim LastRow3 As Long
            LastRow3 = WS.Cells(Rows.Count, "I").End(xlUp).Row
    
        For p = 2 To LastRow3
    
        If WS.Cells(2, 17).Value = WS.Cells(p, 11).Value Then
            WS.Cells(2, 16).Value = WS.Cells(p, 9).Value
        
        End If
            
        Next p
    
        For q = 2 To LastRow3
    
        If WS.Cells(3, 17).Value = WS.Cells(q, 11).Value Then
            WS.Cells(3, 16).Value = WS.Cells(q, 9).Value
        
        End If
    
        Next q
            
        For r = 2 To LastRow3
    
        If WS.Cells(4, 17).Value = WS.Cells(r, 12).Value Then
            WS.Cells(4, 16).Value = WS.Cells(r, 9).Value
        
        End If
            
        Next r
    
    Next WS
    
End Sub
