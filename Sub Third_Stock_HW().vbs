Sub Third_Stock_HW()

    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    
    
    
    Dim i As Long
    Dim ticker As String
    Dim count As Long
    
    
    Dim open_price As Double
    Dim close_price As Double
    Dim price_difference As Double
    Dim percent_change As Double
    Dim total_volume As Variant
    Dim start As Long
    Dim new_start_value As Variant
    
    ' to find last row in data
    Dim last_row As Long
    last_row = Cells(Rows.count, 1).End(xlUp).Row
    
    count = 2
    total_volume = 0
    open_price = 0
    close_price = 0
    price_difference = 0
    percent_change = 0
    start = 2
    
    ' set headers
    Range("I1") = "<ticker>"
    Range("J1") = "<price difference>"
    Range("K1") = "<percentage change>"
    Range("L1") = "<total volume>"
    
    
        For i = 2 To last_row
        
            ' If ticker is different
            
            If Cells(i + 1, 1).Value <> Cells(i, 1) Then
                              
            
                ' put ticker in it's new column home
                Cells(count, 9).Value = Cells(i, 1).Value

                ' put volume in it's new column home
                Cells(count, 12) = total_volume
            
            
                ' account for dividing by 0
                
                If Cells(start, 3) = 0 Then
                    
                    ' run through to account to find 0
                    For new_start_value = start To i
                        If Cells(new_start_value, 3) <> 0 Then
                            start = new_start_value
                            Exit For
                        End If
                    Next new_start_value
                End If
                
            ' define open_price
            open_price = Cells(start, 3)
            
                ' calculate price difference
                price_difference = Cells(i, 6) - Cells(start, 3)
                
                ' put price difference in new column home
                Cells(count, 10) = price_difference
                
                
                If open_price <> 0 Then
                
                ' calculate percent change
                percent_change = (price_difference / open_price) * 100
                
                End If
                
                ' put percent change into new column home
                Cells(count, 11) = percent_change & "%"
                
                 ' conditional formatting
            Select Case price_difference
                Case Is > 0
                    Cells(count, 10).Interior.ColorIndex = 4 ' green color
                Case Is < 0
                    Cells(count, 10).Interior.ColorIndex = 3 ' red color
                Case Else
                    Cells(count, 10).Interior.ColorIndex = 0 ' no color
            End Select
                
                
           ' ----------- everything above is "summary table"
                count = count + 1
                total_volume = 0
                start = (i + 1)
            
           
            
            Else
                
                ' define total_volume
                total_volume = total_volume + Cells(i, 7)
                
            
            End If
            
               
        Next i
        
        ' ---------------- Challenge of best/worst performance
        
        ' variables
        Dim best_stock As Variant
        Dim best_value As Variant
        Dim worst_stock As Variant
        Dim worst_value As Variant
        Dim most_volume_stock As Variant
        Dim most_volume_value As Variant
        Dim j As Long
        
        
        ' assign best_value
        best_value = Cells(2, 11)
        
        ' assign worst_value
        worst_value = Cells(2, 11)
        
        ' assign volume
        most_volume_value = Cells(2, 12)
        
        ' headers/titles
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        
        
        Cells(2, 15) = "Greatest % Increase"
        Cells(3, 15) = "Greatest % Decrease"
        Cells(4, 15) = "Greatest Total Volume"
        
        ' to find last_row in data
        last_row = Cells(Rows.count, 9).End(xlUp).Row
        
        ' start loop for new data
        For j = 2 To last_row
        
            'best value
            If Cells(j, 11) > best_value Then
                best_value = Cells(j, 11)
                best_stock = Cells(j, 9)
            End If
            
            If Cells(j, 11) < worst_value Then
                worst_value = Cells(j, 11)
                worst_stock = Cells(j, 9)
            End If
            
            If Cells(j, 11) < worst_value Then
                most_volume_value = Cells(j, 12)
                most_volume_stock = Cells(j, 9)
            End If
        
            ' put data into new column homes
            Cells(2, 17) = best_value
            Cells(2, 16) = best_stock
            Cells(3, 17) = worst_value
            Cells(3, 16) = worst_stock
            Cells(4, 17) = most_volume_value
            Cells(4, 16) = most_volume_stock
        Next j
        
        
Next
starting_ws.Activate 'activate the worksheet that was originally active

End Sub
