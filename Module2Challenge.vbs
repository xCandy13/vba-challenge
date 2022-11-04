Attribute VB_Name = "Module11"
Sub stocks()
    For Each ws In Worksheets
        'clear output columns
        ws.Range("J:Q").ClearContents
        ws.Range("J:Q").ClearFormats
        'create temp variables
        Dim currentRow As Double
        currentRow = 2
        Dim outputRow As Double
        outputRow = 2
        'set up output columns
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "% Change"
        ws.Cells(1, 13).Value = "Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        'Dim lastRow As Integer 'see iterative ticker boundaries search
        'lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        'loop until we reach the first empty row
        While IsEmpty(ws.Cells(currentRow, 1)) = False
            Dim opening As Double
            Dim closing As Double
            Dim yearly As Double
            Dim percent As Double
            Dim volume As Double
            volume = 0
            Dim ticker As String
            ticker = ws.Cells(currentRow, 1)
            Dim startRow As Double
            startRow = currentRow
            Dim endRow As Double
            endRow = startRow + 250 'hard coded days
            currentRow = endRow + 1
            'For i = currentRow To lastRow 'could not get looping to find ticker boundaries to work quickly
             '   If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
            '            endRow = i
             '           currentRow = i + 1
             '   End If
            'Next
            
            'calc total volume
            For i = startRow To endRow
                volume = volume + ws.Cells(i, 7).Value
            Next
            
            'store opening and closing values
            opening = ws.Cells(startRow, 3).Value
            closing = ws.Cells(endRow, 6).Value
            'calculate yearly and percent change
            yearly = closing - opening
            percent = yearly / opening
            'output for current ticker
            ws.Cells(outputRow, 10).Value = ticker
            ws.Cells(outputRow, 11).Value = yearly
            ws.Cells(outputRow, 12).Value = percent
            ws.Cells(outputRow, 13).Value = volume
            'increment outputRow for next ticker
            outputRow = outputRow + 1
        Wend
        'conditional formatting down yearly column
        'green for positive, red for negative
        For i = 2 To outputRow
            If ws.Cells(i, 12) < 0 Then
                ws.Cells(i, 12).Interior.ColorIndex = 3
            ElseIf ws.Cells(i, 12) > 0 Then
                ws.Cells(i, 12).Interior.ColorIndex = 4
            End If
        Next
        
        'bonus
        'find greatest % increase
            'set max
            'iterate through each row, checking max
            'output final max
        Dim max As Double
        max = ws.Cells(2, 12).Value
        For i = 2 To outputRow
            If ws.Cells(i, 12).Value > max Then
                max = ws.Cells(i, 12).Value
                ticker = ws.Cells(i, 10).Value
            End If
        Next
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = ticker
        ws.Cells(2, 17).Value = max
        'find greatest % decrease
            'set min
            'iterate through each row, checking min
            'output final min
        Dim min As Double
        min = ws.Cells(2, 12).Value
        For i = 2 To outputRow
            If ws.Cells(i, 12).Value < min Then
                min = ws.Cells(i, 12).Value
                ticker = ws.Cells(i, 10).Value
            End If
        Next
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = ticker
        ws.Cells(3, 17).Value = min
        'find greatest volume
            'set maxvolume
            'iterate through each row, checking maxvolume
            'output final maxvolume
        Dim maxVolume As Double
        maxVolume = ws.Cells(2, 13).Value
        For i = 2 To outputRow
            If ws.Cells(i, 13).Value > maxVolume Then
                maxVolume = ws.Cells(i, 13).Value
                ticker = ws.Cells(i, 10).Value
            End If
        Next
        ws.Cells(4, 15).Value = "Greatest Trade Volume"
        ws.Cells(4, 16).Value = ticker
        ws.Cells(4, 17).Value = maxVolume
        'format percent columns
        ws.Range("L:L").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
Next
End Sub

