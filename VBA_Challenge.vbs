Attribute VB_Name = "Module11"
Sub StockData()

'Define variables
Dim ws As Worksheet
Dim t As String 'ticker
Dim op As Double 'opening price
Dim cp As Double 'closing price
Dim st As Integer 'summary table data
Dim yc As Double 'yearly change
Dim pc As Double 'percent change
Dim tsv As Double 'total stock value

'Worksheet Loop
For Each ws In Worksheets

    'Column Labels
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Starting values
    tsv = 0
    first = 1
    st = 2
    
    lr = ws.Cells(Rows.count, "A").End(xlUp).Row    'last row

        'Row Loops
        For i = 2 To lr

            'Ticker Change conditional
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            t = ws.Cells(i, 1).Value

            'Next Ticker
            first = first + 1

            'Values for "opening price" and "closing price"
            op = ws.Cells(first, 3).Value
            cp = ws.Cells(i, 6).Value

            'Add TSVs
            For j = first To i
                tsv = tsv + ws.Cells(j, 7).Value
            Next j

            'Opening price conditional
            If op <> 0 Then
                yc = cp - op
                pc = yc / op
            Else
                pc = yc
            End If
        
            'Populate Summary Table
            ws.Cells(st, 9).Value = t
            ws.Cells(st, 10).Value = yc
            ws.Cells(st, 11).Value = pc

            'Summary Table Formatting
            ws.Cells(st, 11).NumberFormat = ".00%"
            ws.Cells(st, 12).Value = tsv

            'New row in Summary Table
            st = st + 1

            'Clear values
            tsv = 0
            yc = 0
            pc = 0

            'Go to next i
            first = i
        End If

    'End Row Loop
    Next i

'Summary Table Conditional Formatting

    nlr = ws.Cells(Rows.count, "J").End(xlUp).Row   'new last row
        For j = 2 To nlr

            'Positve / Negative value Conditional
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
Next ws
End Sub
