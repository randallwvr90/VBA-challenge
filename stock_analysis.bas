Attribute VB_Name = "Module1"
Sub stock_analysis():
    ' declare a variable of type worksheet
    'Dim sheet_a, sheet_b, sheet_c, sheet_d, sheet_e, sheet_f, sheet_p As Worksheet
    ' set the sheet references
    'Set sheet_a = Worksheets("A")
    'Set sheet_b = Worksheets("B")
    'Set sheet_c = Worksheets("C")
    'Set sheet_d = Worksheets("D")
    'Set sheet_e = Worksheets("E")
    'Set sheet_f = Worksheets("F")
    'Set sheet_p = Worksheets("P")
    
    ' the outer loop goes through each sheet
    For Each ws In Worksheets
        
        ' variables:
        Dim this_ticker As String
        Dim next_ticker As String
        Dim first_open_price As Double
        Dim last_close_price As Double
        Dim price_diff As Double
        Dim pct_change As Double
        Dim vol_range As String
        Dim total_volume As LongLong
        Dim first_row As Long ' the first row for a given stock
        Dim output_row As Integer
        
        ' set these to the value in row 2 to start out
        first_open_price = ws.Cells(2, 3)
        first_row = 2
        output_row = 2
        
        ' what is the last row?
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' set up the output with column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        ' the inner loop cycles through the rows
        For Row = 2 To last_row
            
            ' set the current_ticker and next_ticker
            this_ticker = ws.Cells(Row, 1).Value
            next_ticker = ws.Cells(Row + 1, 1).Value
            
            ' now for some if statements I guess...
            ' if current_ticker is different from next_ticker, then this is the last row of that ticker!
            If (this_ticker <> next_ticker) And (ws.Range("C" & Row).Value <> 0) Then
            
                'do the calculations
                last_close_price = ws.Cells(Row, 6)
                price_diff = last_close_price - first_open_price
                If (first_open_price = 0) And (price_diff = 0) Then 'probably all zeros case...
                    pct_change = 0
                ElseIf (first_open_price = 0) Then 'happens to start at zero but isn't all zeros
                    pct_change = 0 'this is wrong but idk...
                Else 'normal case
                    pct_change = price_diff / first_open_price
                End If
                vol_range = "G" & first_row & ":G" & Row
                total_volume = Application.WorksheetFunction.Sum(Range(vol_range))
                
                ' add the stuff to the sheet or whatever...
                ws.Range("I" & output_row).Value = this_ticker 'ticker symbol
                ws.Range("J" & output_row).Value = price_diff 'price difference
                ws.Range("K" & output_row).Value = pct_change 'percent change
                ws.Range("L" & output_row).Value = total_volume 'total volume
                
                ' reset some variables and increment the output_row
                first_open_price = ws.Cells(Row + 1, 3)
                first_row = Row + 1
                output_row = output_row + 1
            ElseIf (this_ticker <> next_ticker) And (ws.Range("C" & Row).Value = 0) Then
                'just
            End If
            
        Next Row
        
        'formatting...
        'what is the last row of the summary section?
        
        For Row = 2 To last_row
            If IsEmpty(Range("J" & Row).Value) Then
                Exit For
            End If
            If ws.Range("J" & Row).Value >= 0 Then
                ws.Range("J" & Row).Interior.ColorIndex = 10
            ElseIf ws.Range("J" & Row).Value < 0 Then
                ws.Range("J" & Row).Interior.ColorIndex = 3
            End If
        Next Row
        
    Next ws
End Sub
