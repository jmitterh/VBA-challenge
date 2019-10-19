Sub LoopStock()

'define variables
Dim ws As Worksheet
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim total_stock_volume As Double
Dim percent_change As Double
Dim row_data As Integer

'run through each worksheet
For Each ws In Worksheets
    'set headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'setup integers for loop
    row_data = 2
    previous_i = 1
    total_stock_volume = 0
    
    'LastRow = ws.UsedRange.Rows.Count
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'Summary of yearly change, percent change, and total stock volume for each ticker Loop
        For i = 2 To LastRow
            
            'if ticker row name changes execute below
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'ticker value
            ticker = ws.Cells(i, 1).Value
            
            'initiating variable to start on second row
            previous_i = previous_i + 1
            
            'first day open and last day close of the year
            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value
            
            'sum for total stock volume using vol
            For j = previous_i To i
            
                total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value
                
            Next j
            
            'If there is a zero within open data
            If year_open = 0 Then
            
                percent_change = year_close
                
            Else
                yearly_change = year_close - year_open
                
                percent_change = yearly_change / year_open
                
            End If

            'insert values into summary table
            ws.Cells(row_data, 9).Value = ticker
            ws.Cells(row_data, 10).Value = yearly_change
            ws.Cells(row_data, 11).Value = percent_change
            ws.Cells(row_data, 11).NumberFormat = "0.00%"
            ws.Cells(row_data, 12).Value = total_stock_volume
            
            'increment the row for the data summary when if is executed
            row_data = row_data + 1
            
            'return vartiables back to zero
            total_stock_volume = 0
            yearly_change = 0
            percent_change = 0
            
            'tranfer previews i number to variable previous_i
            previous_i = i
        
        End If

    'finish loop
    Next i
    '
    '
    'last row for column k
    klastrow = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
    'define variables
    Increase = 0
    Decrease = 0
    Greatest = 0
    increase_name = "not avialable"
    decrease_name = "not avialable"
    greatest_name = "not avialable"
    
    
        'find max/min for percentage change and the max volume Loop
        For k = 3 To klastrow
        
            'Define previous increment to check
            last_k = k - 1
                        
            'Define current row for percentage
            current_k = ws.Cells(k, 11).Value
            
            'Define Previous row for percentage
            prevous_k = ws.Cells(last_k, 11).Value
            
            'greatest total volume row
            volume = ws.Cells(k, 12).Value
            
            'Prevous greatest volume row
            prevous_vol = ws.Cells(last_k, 12).Value

            
            'Find the increase
            If Increase > current_k And Increase > prevous_k Then
                
                Increase = Increase
                
                'define name for increase percentage
                'increase_name = ws.Cells(k, 9).Value
                
            ElseIf current_k > Increase And current_k > prevous_k Then
                
                Increase = current_k
                
                'define name for increase percentage
                increase_name = ws.Cells(k, 9).Value
                
            ElseIf prevous_k > Increase And prevous_k > current_k Then
            
                Increase = prevous_k
                
                'define name for increase percentage
                increase_name = ws.Cells(last_k, 9).Value
                
            End If
                
                
            'Find the decrease
            If Decrease < current_k And Decrease < prevous_k Then
                
                Decrease = Decrease
                
                'define name for increase percentage
                'decrease_name = ws.Cells(k, 9).Value
                
            ElseIf current_k < Increase And current_k < prevous_k Then
                
                Decrease = current_k
                
                'define name for increase percentage
                decrease_name = ws.Cells(k, 9).Value
                
            ElseIf prevous_k < Increase And prevous_k < current_k Then
            
                Decrease = prevous_k
                
                'define name for increase percentage
                decrease_name = ws.Cells(last_k, 9).Value
                
            End If
                
           'Find the greatest volume
            If Greatest > volume And Greatest > prevous_vol Then
            
                Greatest = Greatest
                
                'define name for greatest volume
                'greatest_name = ws.Cells(k, 9).Value
            
            ElseIf volume > Greatest And volume > prevous_vol Then
            
                Greatest = volume
                
                'define name for greatest volume
                greatest_name = ws.Cells(k, 9).Value
                
            ElseIf prevous_vol > Greatest And prevous_vol > volume Then
                
                Greatest = prevous_vol
                
                'define name for greatest volume
                greatest_name = ws.Cells(last_k, 9).Value
                
            End If
            
        Next k

    'Column names for greatest and less
    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"
    
    'Increase Decrease and Volume results on excel
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest
    
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"

    'last row for column J
    jlastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
     
        'format columns colors
        For j = 2 To jlastrow
            
            'if greater than or less than zero
            If ws.Cells(j, 10) > 0 Then
            
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            Else
            
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
        Next j

'move to next worksheet
Next ws


End Sub

Sub Clear_All()
'
' Clear_All Macro
'
For Each ws In Worksheets

    ws.Columns("I:P").Clear
    
Next ws
    
End Sub



