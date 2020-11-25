Sub stockCheckAllSheets()
    'create a wrap that allows the code to run on all worksheets
    For Each ws In Worksheets
    
    'create variables for start and end values
    Dim start_open_value As Double
    'set start value to zero
    start_open_value = 0

    Dim end_close_value As Double

    'create variable for ticker name
    Dim ticker As String

    'create variable for volume total, set it to zero
    'set to longlong b/c long caused error 6 overflow
    Dim vol_total As LongLong
    vol_total = 0

    'create variable for lastRow for loop
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'create variable to move input_row down for every new ticker
    Dim input_row As Long
    input_row = 1

    'create headers for ticker, yearly change, percent change, total volume
    ws.Range("j1").Value = "Ticker"
    'j1 = cells(1, 10)
    ws.Range("k1").Value = "Yearly Change"
    'k1 = cells(1, 11)
    ws.Range("L1").Value = "Percent Change"
    'L1 = cells(1, 12)
    ws.Range("m1").Value = "Total Volume"
    'm1 = cells(1, 13)

    'create loop
    'dim rw as long b/c data set is v. long
    Dim rw As Long
    For rw = 2 To lastRow
        'create if/elseif to account for when rows change
        'record start value, first vol, ticker, increment input_row up
        If (ws.Cells(rw, 1).Value <> ws.Cells((rw - 1), 1).Value) Then
            start_open_value = ws.Cells(rw, 3).Value
            vol_total = ws.Cells(rw, 7).Value
            input_row = input_row + 1
            ticker = ws.Cells(rw, 1).Value
        'create else/if for when rows are the same
        'add row's vol to the vol total
        ElseIf (ws.Cells(rw, 1).Value = ws.Cells((rw + 1), 1).Value) Then
            vol_total = vol_total + ws.Cells(rw, 7).Value
        'when the last row with the ticker occurs, add close value, final vol value
        ElseIf (ws.Cells(rw, 1).Value <> ws.Cells((rw + 1), 1).Value) Then
            end_close_value = ws.Cells(rw, 6).Value
            vol_total = vol_total + ws.Cells(rw, 7).Value
        
        'input ticker and vol total, calc yearly change and percent change
            ws.Cells((input_row), 10).Value = ticker
            ws.Cells((input_row), 11).Value = (end_close_value) - (start_open_value)
            ws.Cells((input_row), 13).Value = vol_total

        'create if/then for percent change
        'if start value is zero, use 1 as denominator; elseif use regular formula
        'format cells in both cases as percent
            If (start_open_value = 0) Then
                ws.Cells((input_row), 12).Value = (((end_close_value) - (start_open_value)) / 1)
                ws.Cells((input_row), 12).NumberFormat = "0.00%"
            ElseIf (start_open_value <> 0) Then
                ws.Cells((input_row), 12).Value = (((end_close_value) - (start_open_value)) / start_open_value)
                ws.Cells((input_row), 12).NumberFormat = "0.00%"
            End If
        End If
    Next rw

'create new lastRow calc for the length of the ychange/%change table
    Dim lastRow1 As Long
    lastRow1 = ws.Cells(Rows.Count, 10).End(xlUp).Row

'Bonus
'_____________________________________________________
'set variables for max increase, max decrease, max vol
 'values themselves - set to zero
        Dim max_inc As Double
        max_inc = 0
        
        Dim max_dec As Double
        max_dec = 0
        
        'longlong b/c vol amounts are very large
        Dim max_vol As LongLong
        max_vol = 0
 'variables for the max val ticker names
        Dim max_inc_name As String
        
        Dim max_dec_name As String
        
        Dim vol_name As String
        
'entering cell ranges for 'header text'
        ws.Range("o2").Value = "Greatest Percent Increase"
        ws.Range("o3").Value = "Greatest Percent Decrease"
        ws.Range("o4").Value = "Greatest Total Volume"
        ws.Range("p1").Value = "Ticker"
        ws.Range("q1").Value = "Value"
    
'formatting text for max_inc and max_dec as percentage
        ws.Range("q2:q3").NumberFormat = "0.00%"
    
'end bonus chunk
'___________________________________________________________________
'create new for loop to do conditional formatting on y/%change list
    Dim rm As Long
    For rm = 2 To lastRow1
        'conditional formatting for yearly change
        'if cell value is > 0, then green (4)
        'if cell value is < 0, then red(3)
        'if anything else, just do nothing
        If ws.Cells(rm, 11).Value > 0 Then
            ws.Cells(rm, 11).Interior.ColorIndex = 4
        ElseIf ws.Cells(rm, 11).Value < 0 Then
            ws.Cells(rm, 11).Interior.ColorIndex = 3
        End If
    
'start bonus chunk
'____________________________________________________________________
 'create if/then for max vol
 'move down cells in y/%change list, see if they are greater than previous value for max_vol; if so, make it the new max_vol
        If (ws.Cells(rm, 13).Value) > max_vol Then
            max_vol = ws.Cells(rm, 13).Value
            max_vol_name = ws.Cells(rm, 10).Value
        End If
 
 'create if/elseif for max increase and decrease
 'move down cells in y/%change list, see if they are greater than previous value for max_inc/dec; if so, make it the new max_inc/dec
        If (ws.Cells(rm, 12).Value) > max_inc Then
            max_inc = ws.Cells(rm, 12).Value
            max_inc_name = ws.Cells(rm, 10).Value
        ElseIf (ws.Cells(rm, 12).Value) < max_dec Then
            max_dec = ws.Cells(rm, 12).Value
            max_dec_name = ws.Cells(rm, 10).Value
        End If
        
    Next rm

'entering values for final max inc/dec and vol
    ws.Range("p2").Value = max_inc_name
    ws.Range("p3").Value = max_dec_name
    ws.Range("p4").Value = max_vol_name

    ws.Range("q2").Value = max_inc
    ws.Range("q3").Value = max_dec
    ws.Range("q4").Value = max_vol

    Next

End Sub
