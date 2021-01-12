Attribute VB_Name = "Module1"
Sub MarketAnalysis()
Dim ws As Worksheet
For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    

Dim tvol As Double
Dim summarytable As Long
Dim openamt As Double
Dim closeamt As Double
Dim yearlychange As Double
Dim percchange As Double
Dim lastrow As Long
    summarytable = 2
    tvol = 0
    openamt = 0
    closeamt = 0
    yearlychange = 0
    percchange = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 
For i = 2 To lastrow
                
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        openamt = ws.Cells(i, 3).Value
    
End If
                
        tvol = tvol + ws.Cells(i, 7)
        
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
        ws.Cells(summarytable, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(summarytable, 12).Value = tvol

        closeamt = ws.Cells(i, 6).Value
    

        yearlychange = closeamt - openamt
        ws.Cells(summarytable, 10).Value = yearlychange
    

                    'Conditional to format to highlight positive or negative change.
    If yearlychange >= 0 Then
        ws.Cells(summarytable, 10).Interior.ColorIndex = 4
Else
        ws.Cells(summarytable, 10).Interior.ColorIndex = 3
End If
    

    If openamt = 0 And closeamt = 0 Then
        percchange = 0
        ws.Cells(summarytable, 11).Value = percchange
        ws.Cells(summarytable, 11).NumberFormat = "0.00%"
ElseIf openamt = 0 Then

Dim percchange_New As String
    percchange_New = "New Stock"
    ws.Cells(summarytable, 11).Value = percchange
Else
    percchange = yearlychange / openamt
    ws.Cells(summarytable, 11).Value = percchange
    ws.Cells(summarytable, 11).NumberFormat = "0.00%"
End If

    summarytable = summarytable + 1
    
tvol = 0
openamt = 0
closeamt = 0
yearlychange = 0
percchange = 0
                    
End If
Next i

    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
 lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    

Dim topstock As String
Dim topval As Double
    
    topval = ws.Cells(2, 11).Value
    

Dim laststock As String
Dim lastval As Double
    
    lastval = ws.Cells(2, 11).Value
    

Dim highvolstock As String
Dim highvolvalue As Double
    

     highvolvalue = ws.Cells(2, 12).Value
    

For j = 2 To lastrow
    
    If ws.Cells(j, 11).Value > topval Then
        topval = ws.Cells(j, 11).Value
        topstock = ws.Cells(j, 9).Value
End If
    
    If ws.Cells(j, 11).Value < lastval Then
        lastval = ws.Cells(j, 11).Value
        laststock = ws.Cells(j, 9).Value
End If

    If ws.Cells(j, 12).Value > highvolvalue Then
        highvolvalue = ws.Cells(j, 12).Value
        highvolstock = ws.Cells(j, 9).Value
End If
    
Next j
        ws.Cells(2, 16).Value = topstock
        ws.Cells(2, 17).Value = topval
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = laststock
        ws.Cells(3, 17).Value = lastval
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = highvolstock
        ws.Cells(4, 17).Value = highvolvalue
        
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit


        Next ws
    

    End Sub


Sub ClearColumns()
Dim ws As Worksheet

For Each ws In Worksheets

    ws.Columns("I:Q").EntireColumn.Delete
    
Next ws
End Sub
