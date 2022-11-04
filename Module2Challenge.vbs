For Each ws in Worksheets
    'create temp variables
    Dim currentRow as Integer
    currentRow = 2
    Dim outputRow As Integer
    outputRow = 2
    'set up output columns
    ws.Cells(1,10).Value = "Ticker"
    ws.Cells(1,11).Value = "Yearly Change"
    ws.Cells(1,12).Value = "% Change"
    ws.Cells(1,13).Value = "Total Volume"
    Dim lastRow As Integer
    lastRow = ws.Cells(Rows.Count,1).End(xlUp).row 'find last row index
    'loop until empty row
    While IsEmpty(ws.Cells(currentRow,1)) = False
        Dim opening As Double
        Dim closing As Double
        Dim yearly As Double
        Dim percent As Double
        Dim volume As Double
        Dim ticker As Double
        ticker = ws.Cells(currentRow,1)
        Dim startRow As Integer
        startRow = currentRow
        Dim endRow As Integer
        For i = currentRow To lastRow
            If Cells(i,1).Value = ticker And (i+1,1).Value <>ticker Then
                endRow = i
                currentRow = i+1
            End If
        Next
        'calc total volume
        For i = startRow to endRow
            volume = volume + ws.Cells(i,7).Value
        Next
        'store opening and closing values
        opening = ws.Cells(startRow,3).Value
        closing = ws.Cells(endRow,6).Value
        'calculate yearly and percent change
        yearly = closing - opening
        percent = yearly/opening
        'output for current ticker
        ws.Cells(1,10).Value = ticker
        ws.Cells(1,11).Value = yearly
        ws.Cells(1,12).Value = percent
        ws.Cells(1,13).Value = volume
        'increment outputRow for next ticker
        outputRow = outputRow + 1
    Wend
    'conditional formatting down yearly column
    'green for positive, red for negative
    For i = 2 To outputRow
        if ws.Cells(i,2)<0 Then
            ws.Cells(i,2).Interior.ColorIndex = 3
        elseif ws.Cells(i,2)>0 Then
            ws.Cells(i,2).Interior.ColorIndex = 4
        End If

'bonus
'find greatest % increase
    'set max
    'iterate through each row, checking max
    'output final max
'find greatest % decrease
    'set min
    'iterate through each row, checking min
    'output final min
'find greatest volume
    'set maxvolume
    'iterate through each row, checking maxvolume
    'output final maxvolume
